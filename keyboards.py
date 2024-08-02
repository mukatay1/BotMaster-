from aiogram.types import ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram import types
from datetime import datetime, timedelta
from aiogram.types import KeyboardButton


def get_reply_keyboard(is_admin: bool) -> ReplyKeyboardMarkup:
    if is_admin:
        kb = [
            [KeyboardButton(text="Пришел")],
            [KeyboardButton(text="Ушел")],
            [KeyboardButton(text="Отъезд")],
            [KeyboardButton(text="Отчет")],
            [KeyboardButton(text="Опоздуны")],
        ]
    else:
        kb = [
            [KeyboardButton(text="Пришел")],
            [KeyboardButton(text="Ушел")],
            [KeyboardButton(text="Отъезд")],
        ]

    keyboard = ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
        input_field_placeholder="Нажмите на кнопку"
    )
    return keyboard


def create_date_keyboard() -> InlineKeyboardMarkup:
    kb = []
    today = datetime.today()

    for i in range(7):
        date = today - timedelta(days=i)
        date_str = date.strftime("%Y-%m-%d")
        button = InlineKeyboardButton(text=date_str, callback_data=f"report_{date_str}")
        kb.append([button])

    keyboard = InlineKeyboardMarkup(inline_keyboard=kb)
    return keyboard


def get_reply_type_keyboard() -> InlineKeyboardMarkup:
    kb = [
        [InlineKeyboardButton(text="Объект", callback_data="type_object")],
        [InlineKeyboardButton(text="Личный", callback_data="type_personal")],
    ]
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=kb
    )
    return keyboard

def get_supervisor_keyboard() -> InlineKeyboardMarkup:
    supervisors = ["Кексель Кристина", "Тайбупенова Шолпан"]
    kb = [[InlineKeyboardButton(text=supervisor, callback_data=f"supervisor_{i}")] for i, supervisor in enumerate(supervisors)]

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=kb
    )
    return keyboard

def get_return_keyboard() -> InlineKeyboardMarkup:
    kb = [
        [InlineKeyboardButton(text="Приезд", callback_data="return")],
    ]
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=kb
    )
    return keyboard