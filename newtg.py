
import asyncio
import requests
import tempfile
import os

import logging
logging.basicConfig(level=logging.DEBUG)

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)
from telegram import InlineKeyboardButton, InlineKeyboardMarkup


# ----------------- CONFIG -----------------

BOT_TOKEN = ""
API_URL = "https://check.lexiqo.ru/api/v1/upload/"
STATIC_URL = "https://check.lexiqo.ru/static/docs/"
USER_TOKEN = ""
ADMIN_TOKEN = ""

ALLOWED_STYLES = ["ВШЭ", "МГУ"]
ALLOWED_FORMATS = ["ГОСТ"]
ALLOWED_DICTIONARIES = ["Базовый"]
ALLOWED_SKIPPAGES = ["0", "1", "2"]

DEFAULT_SETTINGS = {
    "style": "",
    "format": "",
    "dictionary": "",
    "skip_pages": ""
}

USER_SETTINGS = {}

# ----------------- HELPERS -----------------

def get_user_settings(user_id: int) -> dict:
    return USER_SETTINGS.get(user_id, DEFAULT_SETTINGS.copy())

def set_user_setting(user_id: int, key: str, value: str):
    settings = get_user_settings(user_id)
    settings[key] = value
    USER_SETTINGS[user_id] = settings

def style_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(text=s, callback_data=f"style:{s}")]
        for s in ALLOWED_STYLES
    ])

def format_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(text=f, callback_data=f"format:{f}")]
        for f in ALLOWED_FORMATS
    ])

def dictionary_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(text=d, callback_data=f"dictionary:{d}")]
        for d in ALLOWED_DICTIONARIES
    ])

def skippages_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(text=d, callback_data=f"page:{d}")]
        for d in ALLOWED_SKIPPAGES
    ])


# ----------------- COMMANDS -----------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет!\n\n"
        "🤖 Я бот-помощник Lexify: помогу тебе проверить твою курсовую работу, диплом или отчет.\n\n"
        "📎 Отправь мне в чат документ Word (.doc или .docx) 😊 (поддерживаем несколько документов сразу)",
        parse_mode="HTML"
    )

async def set_style(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "⚙️ Давай выберем нужные для тебя настройки.\n\n"
        "Сначала выбери стиль оформления документа:",
        parse_mode="HTML",
        reply_markup=style_keyboard()
    )

async def set_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Выбери формат ссылок на источники (в большинстве случаев — это ГОСТ):",
        reply_markup=format_keyboard()
    )

async def set_dictionary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Выбери словарь для проверки орфографии:",
        reply_markup=dictionary_keyboard()
    )

async def set_skippages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Сколько страниц в начале документа пропустить (например, титульник / содержание):",
        reply_markup=skippages_keyboard()
    )

async def settings_com(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
    "🤖 <b>Доступные команды бота:</b>\n\n"
        "/start — 🚀 Начать работу с ботом\n\n"
        "/settings — ⚙ Показать это сообщение\n\n"
        "/set_style — 🎓 Выбрать вуз\n\n"
        "/set_format — 🧪 Выбрать формат ссылок\n\n"
        "/set_dictionary — 📙 Выбрать словарь\n\n"
        "/set_skippages — 📃 Количество пропускаемых страниц\n\n"
        "/view_token — 👀 Посмотреть текущий токен\n\n"
        "/web_check — 🌐 Открыть веб-интерфейс\n\n"
        "📎 Для проверки документа просто отправь мне .doc или .docx файл!",
        parse_mode="HTML"
    )

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    await update.message.reply_text("Мы, роботы, — командные игроки. Вы нам — документ, мы Вам — исправления 🤗. Можешь отправить мне .docx или .doc в чат на проверку")
    

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    action, value = query.data.split(":", 1)

    if action == "style":
        set_user_setting(user_id, "style", value)
        await query.edit_message_text(f"✅ Выбран вуз: <b>{value}</b>.\n\nТеперь выбери формат ссылок на источники литературы (в большинстве случаев — это ГОСТ):", reply_markup=format_keyboard(), parse_mode="HTML")

    elif action == "format":
        set_user_setting(user_id, "format", value)
        await query.edit_message_text(f"✅ Выбран формат: <b>{value}</b>.\n\nВыбери орфографический словарь для проверки слов:", reply_markup=dictionary_keyboard(), parse_mode="HTML")

    elif action == "dictionary":
        set_user_setting(user_id, "dictionary", value)
        await query.edit_message_text(f"✅ Выбран словарь: <b>{value}</b>.\n\nСколько страниц в начале документа пропустить (например, титульник / содержание):", reply_markup=skippages_keyboard(), parse_mode="HTML")
    
    elif action == "page":
        set_user_setting(user_id, "skip_pages", value)
        await query.edit_message_text(f"✅ Пропустим первые несколько страниц ({value}).\n\n🚀 Все настроили!\n\n🔗 Отправляй документы в формате .docx / .doc на проверку в чат:", parse_mode="HTML")


# ----------------- DOCUMENT HANDLER -----------------

async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    document = update.message.document
    original_file_name = document.file_name

    if not document.file_name.endswith((".doc", ".docx")):
        await update.message.reply_text("❌ Только форматы .doc или .docx поддерживаются, можешь отправить мне документ Word.")
        return
    
    #Register user or get token
    BASEURLREG = "https://check.lexiqo.ru/api/v1"

    HEADERS = {
        "Authorization": ADMIN_TOKEN
    }

    r = requests.post(
        f"{BASEURLREG}/admin/users/create/",
        headers=HEADERS,
        json={"email": f"{user_id}@telegram.com"}
    )

    user_current_token = ""

    if r.status_code == 201:
        print("User created:", r.json())
        user_current_token = r.json()['user_token']
    elif r.status_code == 409:

        print("User already exists")
        headers = {"Authorization": ADMIN_TOKEN}

        resp = requests.get(
            "https://check.lexiqo.ru/api/v1/admin/users/view/",
            headers=headers,
            params={"email": f"{user_id}@telegram.com"}
        )

        print(resp.status_code)
        print(resp.text)

        try:
            user_current_token = resp.json()['token']
        except Exception as e:
            msg = await update.message.reply_text(str(e))
    else:
        print("Error:", r.status_code, r.text)
        msg = await update.message.reply_text(str(r.text))
        return
    # ==================================================

    def get_user_settings_db(email):
        headers = {
            "Authorization": ADMIN_TOKEN
        }
        
        params = {
            "email": email
        }
        
        response = requests.get(
            f"{BASEURLREG}/admin/users/settings/",
            headers=headers,
            params=params
        )
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error: {response.status_code}")
            print(response.json())
            return None

    # Usage
    user_settings = get_user_settings_db(f"{user_id}@telegram.com")
    print("settings ", user_settings)

    msg = await update.message.reply_text(f"📤 Загружаю файл...\n\n{original_file_name}")

    tg_file = await document.get_file()

    tmp_dir = tempfile.gettempdir()
    local_path = os.path.join(tmp_dir, document.file_name)

    await tg_file.download_to_drive(local_path)

    settings = get_user_settings(user_id)

    print("session ", settings)

    # ==================================================

    # Default values
    defaults = {
        "style": "ВШЭ",
        "format": "ГОСТ",
        "dictionary": "Базовый",
        "skip_pages": "0"
    }

    # Your data
    settings_check = user_settings

    session_check = settings

    # Priority: 1. Session, 2. Settings, 3. Defaults
    final_settings = {}

    # Apply logic for each field
    fields_to_check = ['style', 'format', 'dictionary', 'skip_pages']

    for field in fields_to_check:
        # First check session (exists and not empty string)
        if field in session_check and session_check[field] != "":
            final_settings[field] = session_check[field]
        # Then check settings (exists and not empty string)
        elif field in settings_check and settings_check[field] != "":
            final_settings[field] = settings_check[field]
        # Finally use defaults
        else:
            final_settings[field] = defaults[field]

    print("Final settings to use:", final_settings)

    # ==================================================

    headers = {
        "Authorization": user_current_token
    }

    # ---- UPLOAD FILE (FILE IS CLOSED PROPERLY) ----
    with open(local_path, "rb") as f:
        files = {
            "file": (
                document.file_name,
                f,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        }

        response = requests.post(
            API_URL,
            headers=headers,
            files=files,
            data=final_settings,
            timeout=60
        )

    if response.status_code != 200:
        await msg.edit_text(
            f"❌ API error\n"
            f"Status: {response.status_code}\n"
            f"{response.text[0:100]}"
        )
        return

    file_token = response.json().get("file_token")
    file_url = f"{STATIC_URL}{file_token}.withNotes.docx"

    # ---- PROGRESS POLLING (5 TRIES) ----
    progress_lines = []

    MAX_TRIES = 11

    for attempt in range(1, MAX_TRIES):
        await asyncio.sleep(5)

        r = requests.head(file_url)

        if r.status_code == 200:
            progress_lines.append(f"Прогресс {attempt}/{MAX_TRIES-1} ✅ Готово")
            await msg.edit_text('✅ Готово\n\n📊 Статистика:\n(soon)\n\n' + f'{original_file_name}\n\n🔗 Файл доступен по ссылке: <a href="{file_url}">открыть</a>',
                                parse_mode="HTML"
            )
            return
        else:
            progress_lines.append(f"Прогресс... {attempt}/{MAX_TRIES-1} ⏳")
            await msg.edit_text(f"📄 Проверяю файл...\n\n{original_file_name}\n\n" + str(progress_lines[-1])
            )

    # ---- FINAL MESSAGE AFTER 5 TRIES ----
    await msg.edit_text(
        '⏳ Файл еще проверяется...\n\n'
        + f'{original_file_name}\n\n🔗 Он скоро будет доступен по ссылке: <a href="{file_url}">открыть</a>',
        parse_mode="HTML"
    )

    # Optional cleanup
    try:
        os.remove(local_path)
    except Exception:
        pass

# ----------------- MAIN -----------------

def main():
    from telegram.ext import CallbackQueryHandler
    app = ApplicationBuilder().token(BOT_TOKEN).read_timeout(60).write_timeout(60).connect_timeout(60).pool_timeout(60).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("settings", set_style))
    #app.add_handler(CommandHandler("set_format", set_format))
    #app.add_handler(CommandHandler("set_dictionary", set_dictionary))
    #app.add_handler(CommandHandler("set_skippages", set_skippages))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, document_handler))

    app.run_polling(
        poll_interval=3.0,            # Polling interval in seconds
        timeout=60,                   # Polling timeout
        drop_pending_updates=False,    # Optional: drop old updates on start
        allowed_updates=None
    )

if __name__ == "__main__":
    main()
