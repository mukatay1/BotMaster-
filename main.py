import asyncio
import logging
import sys
import os
from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.fsm.storage.memory import MemoryStorage
from dotenv import load_dotenv
from handlers import register_handlers, date_router
from database import init_db
from utils.schedule import schedule_jobs

load_dotenv()
PRODUCT = os.getenv('PRODUCT')

if int(PRODUCT):
    API_TOKEN = os.getenv('BOT_TOKEN')
else:
    API_TOKEN = os.getenv('TEST_BOT_TOKEN')

storage = MemoryStorage()

async def main() -> None:
    bot = Bot(token=API_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    register_handlers(dp)
    dp.include_router(date_router)

    init_db()
    schedule_jobs(bot)

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())
