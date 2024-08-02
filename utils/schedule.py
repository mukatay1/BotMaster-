import schedule
import asyncio
from aiogram import Bot
from database import SessionLocal, Employee

async def send_morning_message(bot: Bot) -> None:
    db = SessionLocal()
    employees = db.query(Employee).all()
    for employee in employees:
        try:
            message_text = (
                f"<b>🕰️ Доброе утро, {employee.fio}!</b>\n\n"
                f"<b>Напоминание:</b> <i>Не забудьте отметить ваше прибытие сегодня.</i>"
            )
            await bot.send_message(employee.telegram_id, message_text)
        except Exception as e:
            print(f"Не удалось отправить сообщение сотруднику {employee.telegram_id}: {e}")
    db.close()


async def send_evening_message(bot: Bot) -> None:

    db = SessionLocal()
    employees = db.query(Employee).all()
    for employee in employees:
        try:
            message_text = (
                f"<b>🌆 Добрый вечер, {employee.fio}!</b>\n\n"
                f"<b>Напоминание:</b> <i>Не забудьте отметить ваше время ухода сегодня.</i>"
            )
            await bot.send_message(employee.telegram_id, message_text)
            print(message_text)
        except Exception as e:
            print(f"Не удалось отправить сообщение сотруднику {employee.telegram_id}: {e}")
    db.close()

async def run_schedule(bot: Bot) -> None:
    while True:
        schedule.run_pending()
        await asyncio.sleep(1)

def schedule_jobs(bot: Bot) -> None:
    schedule.every().monday.at("08:50").do(lambda: asyncio.create_task(send_morning_message(bot)))
    schedule.every().tuesday.at("08:50").do(lambda: asyncio.create_task(send_morning_message(bot)))
    schedule.every().wednesday.at("08:50").do(lambda: asyncio.create_task(send_morning_message(bot)))
    schedule.every().thursday.at("08:50").do(lambda: asyncio.create_task(send_morning_message(bot)))
    schedule.every().friday.at("08:50").do(lambda: asyncio.create_task(send_morning_message(bot)))

    schedule.every().monday.at("19:00").do(lambda: asyncio.create_task(send_evening_message(bot)))
    schedule.every().tuesday.at("19:00").do(lambda: asyncio.create_task(send_evening_message(bot)))
    schedule.every().wednesday.at("19:00").do(lambda: asyncio.create_task(send_evening_message(bot)))
    schedule.every().thursday.at("19:00").do(lambda: asyncio.create_task(send_evening_message(bot)))
    schedule.every().friday.at("19:00").do(lambda: asyncio.create_task(send_evening_message(bot)))

    loop = asyncio.get_event_loop()
    loop.create_task(run_schedule(bot))
