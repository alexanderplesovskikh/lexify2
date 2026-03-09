
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
STATIC_URL = "https://check.lexqio.ru/static/docs/"
USER_TOKEN = ""
ADMIN_TOKEN = ""

ALLOWED_STYLES = ["ВШЭ"]
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
        "🤖 Я бот-помощник: помогу тебе проверить твою курсовую работу, диплом или отчет.\n\n"
        "/check"
    )

async def set_style(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Выбери стиль оформления документа:",
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

    #"🔜 — 🆕 Получить новый токен\n\n"

async def web_check(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Command to send user a link to web interface with their token"""
    user_id = update.effective_user.id
    
    # Get user token from the API
    BASEURLREG = "https://check.lexiqo.ru/api/v1"
    HEADERS = {"Authorization": ADMIN_TOKEN}
    
    try:
        # First, try to get the user info
        response = requests.get(
            f"{BASEURLREG}/admin/users/view/",
            headers=HEADERS,
            params={"email": f"{user_id}@telegram.com"}
        )
        
        if response.status_code == 200:
            user_data = response.json()
            user_token = user_data.get('token', 'Not found')
            
            # Create the check URL with token
            check_url = f"https://check.lexiqo.ru/check?token={user_token}"
            
            await update.message.reply_text(
                f"🌐 <b>Веб-интерфейс для проверки документов</b>\n\n"
                f"🔗 <a href=\"{check_url}\">Нажмите здесь, чтобы открыть веб-интерфейс</a>\n\n"
                f"📝 <b>Что произойдет:</b>\n"
                f"1. Откроется страница check.lexiqo.ru\n"
                f"2. Ваш токен будет автоматически сохранен в браузере\n"
                f"3. Вы сможете загружать документы через удобный интерфейс\n\n"
                f"💡 <b>Важно:</b>\n"
                f"• Ссылка содержит ваш личный токен\n"
                f"• Не делитесь этой ссылкой с другими\n"
                f"• Токен сохранится в localStorage вашего браузера",
                parse_mode="HTML",
                disable_web_page_preview=False
            )
            
        elif response.status_code == 404:
            await update.message.reply_text(
                "❌ Профиль не найден.\n\n"
                "У вас еще нет токена. Отправьте любой .doc/.docx файл для автоматического создания профиля."
            )
        else:
            await update.message.reply_text(
                f"⚠️ Ошибка при получении токена:\n"
                f"Статус: {response.status_code}\n"
                f"Ответ: {response.text[:100]}"
            )
            
    except Exception as e:
        await update.message.reply_text(
            f"❌ Ошибка подключения к серверу:\n{str(e)}"
        )

async def newtoken_com(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
        "New token"
    )

async def viewtoken_com(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    wait_msg = await update.message.reply_text("⏳ Подождите...")
    
    # Get user token from the API
    BASEURLREG = "https://check.lexiqo.ru/api/v1"
    HEADERS = {"Authorization": ADMIN_TOKEN}
    
    try:
        # First, try to get the user info
        response = requests.get(
            f"{BASEURLREG}/admin/users/view/",
            headers=HEADERS,
            params={"email": f"{user_id}@telegram.com"}
        )
        
        if response.status_code == 200:
            user_data = response.json()
            user_token = user_data.get('token', 'Not found')
            
            # Format the token for display (show first and last few characters)
            token_display = user_token
            if len(user_token) > 20:
                token_display = f"{user_token[:8]}...{user_token[-8:]}"
            
            check_url = f"https://check.lexiqo.ru/check?token={user_token}"
            
            await wait_msg.edit_text(
                "👋 <b>Привет!</b>\n\n"
                "🤖 Я бот <b>Lexify</b> — помогу быстро проверить и оформить курсовую, отчёт или диплом.\n\n"
                "🌐 <b>Веб-интерфейс для проверки документов:</b>\n\n"
                f"👉 <a href=\"{check_url}\">Открыть страницу проверки</a> 👈\n\n"
                "📌 На странице ты сможешь:\n"
                "• загрузить файл Word\n"
                "• выбрать стиль и параметры оформления\n"
                "• получить готовый файл с замечаниями\n\n",
                parse_mode="HTML",
                disable_web_page_preview=True
            )
        
        else:
            create_resp = requests.post(
                f"{BASEURLREG}/admin/users/create/",
                headers=HEADERS,
                json={"email": f"{user_id}@telegram.com"}
            )

            # First, try to get the user info
            response = requests.get(
                f"{BASEURLREG}/admin/users/view/",
                headers=HEADERS,
                params={"email": f"{user_id}@telegram.com"}
            )
            
            if response.status_code == 200:
                user_data = response.json()
                user_token = user_data.get('token', 'Not found')
                
                # Format the token for display (show first and last few characters)
                token_display = user_token
                if len(user_token) > 20:
                    token_display = f"{user_token[:8]}...{user_token[-8:]}"

            check_url = f"https://check.lexiqo.ru/check?token={user_token}"
            
            await wait_msg.edit_text(
                "👋 <b>Привет!</b>\n\n"
                "🤖 Я бот <b>Lexify</b> — помогу быстро проверить и оформить курсовую, отчёт или диплом.\n\n"
                "🌐 <b>Веб-интерфейс для проверки документов:</b>\n\n"
                f"👉 <a href=\"{check_url}\">Открыть страницу проверки</a> 👈\n\n"
                "📌 На странице ты сможешь:\n"
                "• загрузить файл Word\n"
                "• выбрать стиль и параметры оформления\n"
                "• получить готовый файл с замечаниями\n\n",
                parse_mode="HTML",
                disable_web_page_preview=True
            )
            
        '''if response.status_code == 404:
            # User doesn't exist yet - offer to create one
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("Создать профиль", callback_data="create_profile")]
            ])
            
            await update.message.reply_text(
                "❌ Профиль не найден.\n\n"
                "У вас еще нет токена, так как вы не создавали профиль.\n\n"
                "Отправьте любой .doc/.docx файл для автоматического создания профиля "
                "или нажмите кнопку ниже:",
                reply_markup=keyboard
            )
            
        else:
            await update.message.reply_text(
                f"⚠️ Ошибка при получении токена:\n"
                f"Статус: {response.status_code}\n"
                f"Ответ: {response.text[:100]}"
            )'''
            
    except Exception as e:
        await wait_msg.edit_text(
            f"❌ Ошибка подключения к серверу:\n{str(e)}"
        )

# ----------------- TEXT HANDLER -----------------

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    await update.message.reply_text("👋 <b>Привет!</b>\n\n/start — 🚀 чтобы начать", parse_mode="HTML")
    

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    action, value = query.data.split(":", 1)

    if action == "style":
        set_user_setting(user_id, "style", value)
        await query.edit_message_text(f"✅ Выбран вуз: <b>{value}</b>.\n\nТеперь выбери формат ссылок на источники (в большинстве случаев — это ГОСТ):", reply_markup=format_keyboard(), parse_mode="HTML")

    elif action == "format":
        set_user_setting(user_id, "format", value)
        await query.edit_message_text(f"✅ Выбран формат: <b>{value}</b>.\n\nВыбери орфографический словарь для проверки слов:", reply_markup=dictionary_keyboard(), parse_mode="HTML")

    elif action == "dictionary":
        set_user_setting(user_id, "dictionary", value)
        await query.edit_message_text(f"✅ Выбран словарь: <b>{value}</b>.\n\nСколько страниц в начале документа пропустить (например, титульник / содержание):", reply_markup=skippages_keyboard(), parse_mode="HTML")
    
    elif action == "page":
        set_user_setting(user_id, "skip_pages", value)
        await query.edit_message_text(f"✅ Пропустим первые несколько страниц ({value}).\n\n🚀 Все готово!\n\n🔗 Отправляй документы в формате .docx / .doc на проверку в чат:", parse_mode="HTML")


# ----------------- DOCUMENT HANDLER -----------------

async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    document = update.message.document

    msg = await update.message.reply_text("👋 <b>Привет!</b>\n\n/start — 🚀 чтобы начать", parse_mode="HTML")


def main():
    from telegram.ext import CallbackQueryHandler, MessageHandler, filters
    app = ApplicationBuilder().token(BOT_TOKEN).read_timeout(30).write_timeout(30).connect_timeout(30).pool_timeout(30).build()

    app.add_handler(CommandHandler("start", viewtoken_com))
    app.add_handler(CommandHandler("settings", settings_com))  # Missing
    app.add_handler(CommandHandler("set_style", set_style))    # Missing
    app.add_handler(CommandHandler("set_format", set_format))  # Missing
    app.add_handler(CommandHandler("set_dictionary", set_dictionary))  # Missing
    app.add_handler(CommandHandler("set_skippages", set_skippages))    # Missing
    app.add_handler(CommandHandler("web_check", web_check))    # Missing
    app.add_handler(CommandHandler("view_token", viewtoken_com))  # Alias
    
    app.add_handler(CallbackQueryHandler(button_handler))
    
    # DOCUMENT HANDLER - THIS IS MISSING!
    app.add_handler(MessageHandler(filters.Document.ALL, document_handler))
    
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    
    app.run_polling(
        poll_interval=3.0,
        timeout=30,
        drop_pending_updates=True,
        allowed_updates=None
    )

if __name__ == "__main__":
    main()