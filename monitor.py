import asyncio
import requests
from telegram import Bot
from dotenv import load_dotenv
import os

load_dotenv()

BOT_TOKEN = ""
CHAT_ID = "1216918422"

CHECK_INTERVAL = 30 * 60  # 30 минут

SITES = {
    "app.lexiqo.ru": "https://app.lexiqo.ru",
    "beta.lexiqo.ru": "https://beta.lexiqo.ru",
}

def check_site(url: str) -> bool:
    try:
        response = requests.get(url, timeout=10)
        return response.status_code < 400
    except requests.RequestException:
        return False


async def send_status(bot: Bot):
    messages = []

    for name, url in SITES.items():
        status = "Доступен!" if check_site(url) else "Недоступен!"
        messages.append(f"{name}: {status}")

    text = "Статус сайтов:\n\n" + "\n".join(messages)
    await bot.send_message(chat_id=CHAT_ID, text=text)


async def main():
    bot = Bot(token=BOT_TOKEN)

    while True:
        await send_status(bot)
        await asyncio.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    asyncio.run(main())