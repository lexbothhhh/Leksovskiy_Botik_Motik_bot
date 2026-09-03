import os
import re
import asyncio
import psycopg2
from openpyxl import load_workbook

from flask import Flask, request
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)


# =========================================================
# НАСТРОЙКИ
# =========================================================

BOT_TOKEN = os.environ["BOT_TOKEN"]

ADMIN_ID_1 = int(os.environ["ADMIN_ID_1"])
ADMIN_ID_2 = int(os.environ["ADMIN_ID_2"])

ADMIN_IDS = {ADMIN_ID_1, ADMIN_ID_2}

# НЕ МЕНЯЕМ БЕЗ НЕОБХОДИМОСТИ:
# это секрет для webhook, а не название бота
WEBHOOK_SECRET = "sklad-kustikov-2026-secret-8472"

EXCEL_FILE = "flowers.xlsx"

app = Flask(__name__)

telegram_app = (
    Application.builder()
    .token(BOT_TOKEN)
    .updater(None)
    .build()
)


# =========================================================
# DATABASE
# =========================================================

def get_database_url():
    database_url = os.environ.get("DATABASE_URL")

    if database_url:
        return database_url

    secret_files = [
        "/etc/secrets/DATABASE_URL",
        "/etc/secrets/database_url",
        "/etc/secrets/neon_url",
    ]

    for path in secret_files:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                value = f.read().strip()

            if value:
                return value

    raise RuntimeError("DATABASE_URL не найден")


def get_db():
    return psycopg2.connect(get_database_url())


# =========================================================
# NORMALIZE
# =========================================================

def normalize(text):
    if text is None:
        return ""

    text = str(text).strip().lower()
    text = text.replace("ё", "е")
    text = re.sub(r"\s+", " ", text)

    return text


# =========================================================
# DATABASE INIT
# =========================================================

def init_db():

    conn = get_db()
    cur = conn.cursor()

    # -----------------------------------------------------
    # FLOWERS
    # -----------------------------------------------------

    cur.execute("""
        CREATE TABLE IF NOT EXISTS flowers (
            id SERIAL PRIMARY KEY,
            flower TEXT NOT NULL,
            person TEXT NOT NULL
        )
    """)

    cur.execute("""
        ALTER TABLE flowers
        ADD COLUMN IF NOT EXISTS flower_id TEXT
    """)

    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_flowers_flower
        ON flowers(flower)
    """)

    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_flowers_flower_id
        ON flowers(flower_id)
    """)

    # -----------------------------------------------------
    # PLAYERS
    # -----------------------------------------------------

    cur.execute("""
        CREATE TABLE IF NOT EXISTS players (
            player_id TEXT PRIMARY KEY,
            nickname TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_players_nickname
        ON players(nickname)
    """)

    # -----------------------------------------------------
    # ADMINS
    # -----------------------------------------------------

    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_admins (
            telegram_id BIGINT PRIMARY KEY,
            role TEXT NOT NULL DEFAULT 'admin',
            added_by BIGINT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # -----------------------------------------------------
    # USERS
    # -----------------------------------------------------

    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_users (
            telegram_id BIGINT PRIMARY KEY,
            username TEXT,
            first_name TEXT,
            last_name TEXT,
            chat_id BIGINT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # -----------------------------------------------------
    # SETTINGS
    # -----------------------------------------------------

    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT NOT NULL
        )
    """)

    # -----------------------------------------------------
    # LOGS
    # -----------------------------------------------------

    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_logs (
            id SERIAL PRIMARY KEY,
            admin_id BIGINT,
            action TEXT,
            details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # -----------------------------------------------------
    # GROUP TRIGGER
    # -----------------------------------------------------

    cur.execute("""
        INSERT INTO bot_settings (
            setting_key,
            setting_value
        )
        VALUES ('group_triggers', 'вжух')
        ON CONFLICT (setting_key)
        DO NOTHING
    """)

    # -----------------------------------------------------
    # PERMANENT ADMINS
    # -----------------------------------------------------

    for admin_id in ADMIN_IDS:

        cur.execute("""
            INSERT INTO bot_admins (
                telegram_id,
                role
            )
            VALUES (%s, 'admin')
            ON CONFLICT (telegram_id)
            DO UPDATE SET role = 'admin'
        """, (admin_id,))

    # -----------------------------------------------------
    # MIGRATE OLD FLOWERS -> PLAYERS
    # -----------------------------------------------------

    cur.execute("""
        SELECT DISTINCT
            flower_id,
            person
        FROM flowers
        WHERE flower_id IS NOT NULL
          AND flower_id <> ''
    """)

    rows = cur.fetchall()

    for player_id, nickname in rows:

        if not player_id or not nickname:
            continue

        cur.execute("""
            INSERT INTO players (
                player_id,
                nickname
            )
            VALUES (%s, %s)
            ON CONFLICT (player_id)
            DO UPDATE SET
                nickname = EXCLUDED.nickname,
                updated_at = CURRENT_TIMESTAMP
        """, (
            str(player_id).strip(),
            str(nickname).strip()
        ))

    conn.commit()

    cur.close()
    conn.close()


# =========================================================
# EXCEL IMPORT
# =========================================================

CHECKED_VALUES = {
    "да",
    "д",
    "yes",
    "y",
    "1",
    "true",
    "истина",
    "x",
    "х",
    "✓",
    "✔",
    "☑",
    "☒",
    "+"
}


def import_excel():

    if not os.path.exists(EXCEL_FILE):
        print("Excel файл не найден:", EXCEL_FILE)
        return

    try:

        wb = load_workbook(
            EXCEL_FILE,
            data_only=True
        )

        ws = wb.active

        headers = [
            cell.value
            for cell in ws[1]
        ]

        print("EXCEL HEADERS:", headers)

        name_column = None
        id_column = None

        for index, header in enumerate(headers):

            if header is None:
                continue

            normalized = normalize(header)

            if normalized in {
                "имя",
                "name",
                "игровое имя",
                "ник",
                "никнейм"
            }:
                name_column = index

            if normalized in {
                "id",
                "ид",
                "идентификатор"
            }:
                id_column = index

        if name_column is None:
            print("Не найдена колонка имени")
            return

        flower_columns = []

        for index, header in enumerate(headers):

            if header is None:
                continue

            if index in {
                name_column,
                id_column
            }:
                continue

            flower_columns.append(
                (index, str(header).strip())
            )

        conn = get_db()
        cur = conn.cursor()

        for row in ws.iter_rows(
            min_row=2,
            values_only=True
        ):

            if not row:
                continue

            person = row[name_column]

            if person is None:
                continue

            person = str(person).strip()

            if not person:
                continue

            player_id = None

            if id_column is not None:
                value = row[id_column]

                if value is not None:
                    player_id = str(value).strip()

            # ---------------------------------------------
            # PLAYER
            # ---------------------------------------------

            if player_id:

                cur.execute("""
                    INSERT INTO players (
                        player_id,
                        nickname
                    )
                    VALUES (%s, %s)
                    ON CONFLICT (player_id)
                    DO UPDATE SET
                        nickname = EXCLUDED.nickname,
                        updated_at = CURRENT_TIMESTAMP
                """, (
                    player_id,
                    person
                ))

            # ---------------------------------------------
            # FLOWERS
            # ---------------------------------------------

            for column_index, flower_name in flower_columns:

                if column_index >= len(row):
                    continue

                value = row[column_index]

                if value is None:
                    continue

                checked = normalize(value)

                if checked not in CHECKED_VALUES:
                    continue

                cur.execute("""
                    SELECT 1
                    FROM flowers
                    WHERE flower = %s
                      AND person = %s
                    LIMIT 1
                """, (
                    flower_name,
                    person
                ))

                exists = cur.fetchone()

                if not exists:

                    cur.execute("""
                        INSERT INTO flowers (
                            flower,
                            person,
                            flower_id
                        )
                        VALUES (%s, %s, %s)
                    """, (
                        flower_name,
                        person,
                        player_id
                    ))

        conn.commit()

        cur.close()
        conn.close()

        print("Excel импорт завершён")

    except Exception as e:

        print("ОШИБКА EXCEL:", e)


# =========================================================
# USERS
# =========================================================

def save_user(user, chat_id):

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO bot_users (
            telegram_id,
            username,
            first_name,
            last_name,
            chat_id
        )
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (telegram_id)
        DO UPDATE SET
            username = EXCLUDED.username,
            first_name = EXCLUDED.first_name,
            last_name = EXCLUDED.last_name,
            chat_id = EXCLUDED.chat_id,
            updated_at = CURRENT_TIMESTAMP
    """, (
        user.id,
        user.username,
        user.first_name,
        user.last_name,
        chat_id
    ))

    conn.commit()

    cur.close()
    conn.close()


# =========================================================
# ADMINS
# =========================================================

def is_admin(user_id):

    if user_id in ADMIN_IDS:
        return True

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 1
        FROM bot_admins
        WHERE telegram_id = %s
          AND role = 'admin'
        LIMIT 1
    """, (user_id,))

    result = cur.fetchone()

    cur.close()
    conn.close()

    return result is not None


def is_owner(user_id):

    return user_id in ADMIN_IDS


# =========================================================
# LOGGING
# =========================================================

def log_action(
    admin_id,
    action,
    details=""
):

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO bot_logs (
            admin_id,
            action,
            details
        )
        VALUES (%s, %s, %s)
    """, (
        admin_id,
        action,
        details
    ))

    conn.commit()

    cur.close()
    conn.close()


# =========================================================
# SETTINGS
# =========================================================

def get_setting(key, default=None):

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT setting_value
        FROM bot_settings
        WHERE setting_key = %s
    """, (key,))

    row = cur.fetchone()

    cur.close()
    conn.close()

    if row:
        return row[0]

    return default


def set_setting(key, value):

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO bot_settings (
            setting_key,
            setting_value
        )
        VALUES (%s, %s)
        ON CONFLICT (setting_key)
        DO UPDATE SET
            setting_value = EXCLUDED.setting_value
    """, (
        key,
        value
    ))

    conn.commit()

    cur.close()
    conn.close()


# =========================================================
# KEYBOARDS
# =========================================================

def main_keyboard(user_id):

    buttons = [
        [
            InlineKeyboardButton(
                "🆔 Мой ID",
                callback_data="my_id"
            )
        ]
    ]

    if is_admin(user_id):

        buttons.append([
            InlineKeyboardButton(
                "⚙️ Админ-панель",
                callback_data="admin_panel"
            )
        ])

    return InlineKeyboardMarkup(buttons)


def admin_keyboard():

    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "👤 Люди",
                callback_data="people"
            ),
            InlineKeyboardButton(
                "🌸 Цветы",
                callback_data="flowers"
            )
        ],
        [
            InlineKeyboardButton(
                "👥 Администраторы",
                callback_data="admins"
            ),
            InlineKeyboardButton(
                "⚡ Группа",
                callback_data="group"
            )
        ],
        [
            InlineKeyboardButton(
                "📊 Статистика",
                callback_data="stats"
            )
        ],
        [
            InlineKeyboardButton(
                "❌ Закрыть",
                callback_data="close"
            )
        ]
    ])


def people_keyboard():

    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "➕ Добавить человека",
                callback_data="person_add"
            )
        ],
        [
            InlineKeyboardButton(
                "✏️ Изменить ник",
                callback_data="person_edit"
            )
        ],
        [
            InlineKeyboardButton(
                "🔎 Найти по ID",
                callback_data="person_find"
            )
        ],
        [
            InlineKeyboardButton(
                "🗑 Удалить человека",
                callback_data="person_delete"
            )
        ],
        [
            InlineKeyboardButton(
                "↩️ Назад",
                callback_data="admin_panel"
            )
        ]
    ])


def flowers_keyboard():

    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "➕ Добавить цветок",
                callback_data="flower_add"
            )
        ],
        [
            InlineKeyboardButton(
                "➖ Удалить цветок",
                callback_data="flower_delete"
            )
        ],
        [
            InlineKeyboardButton(
                "🔎 Найти цветок",
                callback_data="flower_find"
            )
        ],
        [
            InlineKeyboardButton(
                "↩️ Назад",
                callback_data="admin_panel"
            )
        ]
    ])


# =========================================================
# COMMAND STATES
# =========================================================

COMMAND_STATES = {}

PLAYER_ID_RE = re.compile(
    r"^[A-Za-z0-9]+$"
)


def state_key(user_id, chat_id):

    return (
        user_id,
        chat_id
    )


# =========================================================
# START
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user
    chat = update.effective_chat

    save_user(
        user,
        chat.id
    )

    await update.message.reply_text(
        "Лексовский Ботик-Мотик\n\n"
        "Напиши название цветка, чтобы узнать, "
        "у кого он есть.",
        reply_markup=main_keyboard(user.id)
    )


# =========================================================
# ADMIN
# =========================================================

async def admin_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user

    if not is_admin(user.id):

        await update.message.reply_text(
            "⛔ У тебя нет доступа к админ-панели."
        )

        return

    await update.message.reply_text(
        "⚙️ Админ-панель",
        reply_markup=admin_keyboard()
    )


# =========================================================
# MY ID
# =========================================================

async def show_my_id(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    await query.message.reply_text(
        f"🆔 Твой Telegram ID:\n\n"
        f"`{query.from_user.id}`",
        parse_mode="Markdown"
    )


# =========================================================
# FLOWER SEARCH
# =========================================================

async def search_flower(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    message = update.message

    if not message:
        return

    chat = update.effective_chat

    # В группах поиск работает только через триггер
    if chat.type in {
        "group",
        "supergroup"
    }:

        text = message.text or ""

        triggers_raw = get_setting(
            "group_triggers",
            "вжух"
        )

        triggers = [
            normalize(x)
            for x in triggers_raw.split(",")
            if normalize(x)
        ]

        text_normalized = normalize(text)

        found_trigger = None

        for trigger in triggers:

            patterns = [
                trigger,
                trigger + " ",
                trigger + ",",
                trigger + ":",
                trigger + " —",
                trigger + " -"
            ]

            for pattern in patterns:

                if text_normalized.startswith(pattern):

                    found_trigger = trigger
                    break

            if found_trigger:
                break

        if not found_trigger:
            return

        flower = text_normalized[
            len(found_trigger):
        ].strip()

        flower = re.sub(
            r"^[,:—-]\s*",
            "",
            flower
        )

        if not flower:
            return

    else:

        flower = normalize(
            message.text
        )

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT person
        FROM flowers
        WHERE LOWER(REPLACE(flower, 'Ё', 'Е'))
            = LOWER(REPLACE(%s, 'Ё', 'Е'))
        ORDER BY person
    """, (
        flower
    ))

    rows = cur.fetchall()

    cur.close()
    conn.close()

    if not rows:

        await message.reply_text(
            "🌸 Такой цветок не найден."
        )

        return

    people = [
        row[0]
        for row in rows
    ]

    result = (
        f"🌸 {flower}\n\n"
        + "\n".join(
            f"• {person}"
            for person in people
        )
    )

    await message.reply_text(
        result
    )


# =========================================================
# DELETE COMMAND
# =========================================================

async def delete_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user

    if not is_admin(user.id):

        await update.message.reply_text(
            "⛔ Нет доступа."
        )

        return

    key = state_key(
        user.id,
        update.effective_chat.id
    )

    COMMAND_STATES[key] = {
        "command": "delete"
    }

    await update.message.reply_text(
        "🗑 Напиши ID человека, которого нужно удалить.\n\n"
        "Или /cancel для отмены."
    )


# =========================================================
# AT COMMAND
# =========================================================

async def at_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user

    if not is_admin(user.id):

        await update.message.reply_text(
            "⛔ Нет доступа."
        )

        return

    key = state_key(
        user.id,
        update.effective_chat.id
    )

    COMMAND_STATES[key] = {
        "command": "at"
    }

    await update.message.reply_text(
        "🆔 Напиши ID игрока.\n\n"
        "Или /cancel для отмены."
    )


# =========================================================
# ADD COMMAND
# =========================================================

async def add_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user

    if not is_admin(user.id):

        await update.message.reply_text(
            "⛔ Нет доступа."
        )

        return

    key = state_key(
        user.id,
        update.effective_chat.id
    )

    COMMAND_STATES[key] = {
        "command": "add"
    }

    await update.message.reply_text(
        "➕ Введи ID и ник через пробел.\n\n"
        "Например:\n"
        "`ABC123 Лекс`",
        parse_mode="Markdown"
    )


# =========================================================
# CANCEL
# =========================================================

async def cancel_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    key = state_key(
        update.effective_user.id,
        update.effective_chat.id
    )

    COMMAND_STATES.pop(
        key,
        None
    )

    await update.message.reply_text(
        "❌ Действие отменено."
    )


# =========================================================
# TEXT HANDLER
# =========================================================

async def text_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    message = update.message

    if not message:
        return

    user = update.effective_user
    chat = update.effective_chat

    save_user(
        user,
        chat.id
    )

    key = state_key(
        user.id,
        chat.id
    )

    text = message.text or ""

    # -----------------------------------------------------
    # COMMAND STATE
    # -----------------------------------------------------

    if key in COMMAND_STATES:

        state = COMMAND_STATES[key]

        command = state.get(
            "command"
        )

        # ---------------------------------------------
        # ADD
        # ---------------------------------------------

        if command == "add":

            if not is_admin(user.id):
                COMMAND_STATES.pop(key, None)
                return

            parts = text.strip().split(
                maxsplit=1
            )

            if len(parts) != 2:

                await message.reply_text(
                    "❌ Формат:\n"
                    "`ID Ник`",
                    parse_mode="Markdown"
                )

                return

            player_id = parts[0].strip()
            nickname = parts[1].strip()

            if not PLAYER_ID_RE.match(
                player_id
            ):

                await message.reply_text(
                    "❌ ID должен содержать "
                    "только латинские буквы и цифры."
                )

                return

            conn = get_db()
            cur = conn.cursor()

            cur.execute("""
                INSERT INTO players (
                    player_id,
                    nickname
                )
                VALUES (%s, %s)
                ON CONFLICT (player_id)
                DO UPDATE SET
                    nickname = EXCLUDED.nickname,
                    updated_at = CURRENT_TIMESTAMP
            """, (
                player_id,
                nickname
            ))

            conn.commit()

            cur.close()
            conn.close()

            COMMAND_STATES.pop(
                key,
                None
            )

            log_action(
                user.id,
                "add_player",
                f"{player_id} {nickname}"
            )

            await message.reply_text(
                f"✅ Человек добавлен.\n\n"
                f"🆔 ID: {player_id}\n"
                f"👤 Ник: {nickname}"
            )

            return

        # ---------------------------------------------
        # AT
        # ---------------------------------------------

        if command == "at":

            player_id = text.strip()

            conn = get_db()
            cur = conn.cursor()

            cur.execute("""
                SELECT nickname
                FROM players
                WHERE player_id = %s
            """, (
                player_id
            ))

            row = cur.fetchone()

            cur.close()
            conn.close()

            COMMAND_STATES.pop(
                key,
                None
            )

            if not row:

                await message.reply_text(
                    "❌ Игрок с таким ID не найден."
                )

                return

            await message.reply_text(
                f"👤 {row[0]}\n"
                f"🆔 {player_id}"
            )

            return

        # ---------------------------------------------
        # DELETE
        # ---------------------------------------------

        if command == "delete":

            player_id = text.strip()

            conn = get_db()
            cur = conn.cursor()

            cur.execute("""
                SELECT nickname
                FROM players
                WHERE player_id = %s
            """, (
                player_id
            ))

            row = cur.fetchone()

            if not row:

                cur.close()
                conn.close()

                COMMAND_STATES.pop(
                    key,
                    None
                )

                await message.reply_text(
                    "❌ Игрок не найден."
                )

                return

            nickname = row[0]

            cur.execute("""
                DELETE FROM players
                WHERE player_id = %s
            """, (
                player_id
            ))

            cur.execute("""
                DELETE FROM flowers
                WHERE flower_id = %s
            """, (
                player_id
            ))

            conn.commit()

            cur.close()
            conn.close()

            COMMAND_STATES.pop(
                key,
                None
            )

            log_action(
                user.id,
                "delete_player",
                f"{player_id} {nickname}"
            )

            await message.reply_text(
                f"🗑 Удалён:\n"
                f"👤 {nickname}\n"
                f"🆔 {player_id}"
            )

            return

    # -----------------------------------------------------
    # GROUP
    # -----------------------------------------------------

    if chat.type in {
        "group",
        "supergroup"
    }:

        await search_flower(
            update,
            context
        )

        return

    # -----------------------------------------------------
    # PRIVATE
    # -----------------------------------------------------

    await search_flower(
        update,
        context
    )


# =========================================================
# CALLBACKS
# =========================================================

async def callback_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    user_id = query.from_user.id

    data = query.data

    # -----------------------------------------------------
    # MY ID
    # -----------------------------------------------------

    if data == "my_id":

        await query.message.reply_text(
            f"🆔 Твой Telegram ID:\n\n"
            f"`{user_id}`",
            parse_mode="Markdown"
        )

        return

    # -----------------------------------------------------
    # ADMIN PANEL
    # -----------------------------------------------------

    if data == "admin_panel":

        if not is_admin(user_id):
            return

        await query.message.reply_text(
            "⚙️ Админ-панель",
            reply_markup=admin_keyboard()
        )

        return

    # -----------------------------------------------------
    # PEOPLE
    # -----------------------------------------------------

    if data == "people":

        if not is_admin(user_id):
            return

        await query.message.reply_text(
            "👤 Люди",
            reply_markup=people_keyboard()
        )

        return

    # -----------------------------------------------------
    # FLOWERS
    # -----------------------------------------------------

    if data == "flowers":

        if not is_admin(user_id):
            return

        await query.message.reply_text(
            "🌸 Цветы",
            reply_markup=flowers_keyboard()
        )

        return

    # -----------------------------------------------------
    # ADMINS
    # -----------------------------------------------------

    if data == "admins":

        if not is_owner(user_id):

            await query.message.reply_text(
                "⛔ Управлять администраторами "
                "могут только два главных администратора."
            )

            return

        conn = get_db()
        cur = conn.cursor()

        cur.execute("""
            SELECT telegram_id, role
            FROM bot_admins
            ORDER BY telegram_id
        """)

        rows = cur.fetchall()

        cur.close()
        conn.close()

        if not rows:

            text = "👥 Администраторов нет."

        else:

            text = "👥 Администраторы:\n\n"

            for telegram_id, role in rows:

                permanent = (
                    " ⭐"
                    if telegram_id in ADMIN_IDS
                    else ""
                )

                text += (
                    f"• `{telegram_id}` — "
                    f"{role}{permanent}\n"
                )

        await query.message.reply_text(
            text,
            parse_mode="Markdown"
        )

        return

    # -----------------------------------------------------
    # GROUP
    # -----------------------------------------------------

    if data == "group":

        if not is_admin(user_id):
            return

        triggers = get_setting(
            "group_triggers",
            "вжух"
        )

        await query.message.reply_text(
            "⚡ Триггеры группы:\n\n"
            f"`{triggers}`\n\n"
            "Бот реагирует только если сообщение "
            "начинается с точного триггера.",
            parse_mode="Markdown"
        )

        return

    # -----------------------------------------------------
    # STATS
    # -----------------------------------------------------

    if data == "stats":

        if not is_admin(user_id):
            return

        conn = get_db()
        cur = conn.cursor()

        cur.execute("""
            SELECT COUNT(*)
            FROM players
        """)

        players_count = cur.fetchone()[0]

        cur.execute("""
            SELECT COUNT(*)
            FROM flowers
        """)

        flowers_count = cur.fetchone()[0]

        cur.execute("""
            SELECT COUNT(DISTINCT flower)
            FROM flowers
        """)

        unique_flowers = cur.fetchone()[0]

        cur.close()
        conn.close()

        await query.message.reply_text(
            "📊 Статистика\n\n"
            f"👤 Людей: {players_count}\n"
            f"🌸 Записей цветов: {flowers_count}\n"
            f"🌺 Уникальных цветов: {unique_flowers}"
        )

        return

    # -----------------------------------------------------
    # CLOSE
    # -----------------------------------------------------

    if data == "close":

        try:
            await query.message.delete()
        except Exception:
            pass

        return

    # -----------------------------------------------------
    # PERSON ADD
    # -----------------------------------------------------

    if data == "person_add":

        if not is_admin(user_id):
            return

        key = state_key(
            user_id,
            query.message.chat_id
        )

        COMMAND_STATES[key] = {
            "command": "add"
        }

        await query.message.reply_text(
            "➕ Введи ID и ник через пробел.\n\n"
            "Например:\n"
            "`ABC123 Лекс`",
            parse_mode="Markdown"
        )

        return

    # -----------------------------------------------------
    # PERSON FIND
    # -----------------------------------------------------

    if data == "person_find":

        if not is_admin(user_id):
            return

        key = state_key(
            user_id,
            query.message.chat_id
        )

        COMMAND_STATES[key] = {
            "command": "at"
        }

        await query.message.reply_text(
            "🔎 Введи ID человека."
        )

        return

    # -----------------------------------------------------
    # PERSON DELETE
    # -----------------------------------------------------

    if data == "person_delete":

        if not is_admin(user_id):
            return

        key = state_key(
            user_id,
            query.message.chat_id
        )

        COMMAND_STATES[key] = {
            "command": "delete"
        }

        await query.message.reply_text(
            "🗑 Введи ID человека, которого "
            "нужно удалить."
        )

        return

    # -----------------------------------------------------
    # BACK
    # -----------------------------------------------------

    if data == "admin_panel":

        await query.message.reply_text(
            "⚙️ Админ-панель",
            reply_markup=admin_keyboard()
        )

        return


# =========================================================
# TELEGRAM HANDLERS
# =========================================================

telegram_app.add_handler(
    CommandHandler(
        "start",
        start
    )
)

telegram_app.add_handler(
    CommandHandler(
        "admin",
        admin_command
    )
)

telegram_app.add_handler(
    CommandHandler(
        "delete",
        delete_command
    )
)

telegram_app.add_handler(
    CommandHandler(
        "at",
        at_command
    )
)

telegram_app.add_handler(
    CommandHandler(
        "add",
        add_command
    )
)

telegram_app.add_handler(
    CommandHandler(
        "cancel",
        cancel_command
    )
)

telegram_app.add_handler(
    CallbackQueryHandler(
        callback_handler
    )
)

telegram_app.add_handler(
    MessageHandler(
        filters.TEXT & ~filters.COMMAND,
        text_handler
    )
)


# =========================================================
# FLASK
# =========================================================

@app.route(
    "/",
    methods=["GET"]
)
def home():

    return "Лексовский Ботик-Мотик работает!"


@app.route("/telegram/" + WEBHOOK_SECRET, methods=["POST"])
async def telegram_webhook():
    try:
        data = request.get_json(force=True)

        if not data:
            return "OK", 200

        update = Update.de_json(
            data,
            telegram_app.bot
        )

        await telegram_app.initialize()

        try:
            if update.effective_user:
                save_user(
                    update.effective_user,
                    update.effective_chat.id
                    if update.effective_chat
                    else None
                )

            await telegram_app.process_update(update)

        finally:
            await telegram_app.shutdown()

        return "OK", 200

    except Exception as e:
        print("WEBHOOK ERROR:", repr(e))
        return "ERROR", 500


# =========================================================
# STARTUP
# =========================================================

init_db()
import_excel()


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port
    )
