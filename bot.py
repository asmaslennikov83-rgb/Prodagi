import asyncio
import os
import re
import tempfile
from datetime import date, datetime, timedelta

from aiogram import Bot, Dispatcher, F
from aiogram import BaseMiddleware
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery, FSInputFile, Message
from aiogram.utils.keyboard import InlineKeyboardBuilder

from config import (
    TELEGRAM_BOT_TOKEN, WB_API_KEY_1, WB_API_KEY_2,
    CABINET_1_NAME, CABINET_2_NAME, ALLOWED_TELEGRAM_IDS,
)
from report_builder import apply_orders, build_products, make_excel, merge_products
from wb_client import WBApiError, WBClient


class ReportState(StatesGroup):
    custom_dates = State()


bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


class AccessMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, data):
        user = data.get('event_from_user')
        if user and user.id not in ALLOWED_TELEGRAM_IDS:
            text = (
                '⛔ Доступ к боту запрещён.\n\n'
                f'Ваш Telegram ID: <code>{user.id}</code>\n'
                'Передайте этот ID администратору бота.'
            )
            if isinstance(event, CallbackQuery):
                await event.answer('⛔ Нет доступа', show_alert=True)
            elif isinstance(event, Message):
                await event.answer(text, parse_mode='HTML')
            return None
        return await handler(event, data)


dp.message.outer_middleware(AccessMiddleware())
dp.callback_query.outer_middleware(AccessMiddleware())


def start_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text='📊 Сформировать отчёт', callback_data='report')
    return kb.as_markup()


def cabinets_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text=f'1️⃣ {CABINET_1_NAME}', callback_data='cab:1')
    kb.button(text=f'2️⃣ {CABINET_2_NAME}', callback_data='cab:2')
    kb.button(text='🔗 Оба кабинета', callback_data='cab:both')
    kb.adjust(1)
    return kb.as_markup()


def periods_kb():
    kb = InlineKeyboardBuilder()
    for n in (7, 14, 30, 60, 90):
        kb.button(text=f'{n} дней', callback_data=f'days:{n}')
    kb.button(text='📅 Свои даты', callback_data='custom')
    kb.adjust(3, 2, 1)
    return kb.as_markup()


@dp.message(CommandStart())
async def start(m: Message):
    await m.answer('📊 Отчёты Wildberries по заказам FBO + FBS', reply_markup=start_kb())


@dp.callback_query(F.data == 'report')
async def choose_cab(c: CallbackQuery, state: FSMContext):
    await state.clear()
    await c.message.edit_text('Выберите кабинет:', reply_markup=cabinets_kb())
    await c.answer()


@dp.callback_query(F.data.startswith('cab:'))
async def choose_period(c: CallbackQuery, state: FSMContext):
    await state.update_data(cabinet=c.data.split(':', 1)[1])
    await c.message.edit_text('Выберите период. Сегодняшний день в быстрые периоды не входит:', reply_markup=periods_kb())
    await c.answer()


@dp.callback_query(F.data.startswith('days:'))
async def quick_period(c: CallbackQuery, state: FSMContext):
    n = int(c.data.split(':', 1)[1])
    to_d = date.today() - timedelta(days=1)
    from_d = to_d - timedelta(days=n - 1)
    await c.answer()
    await generate(c.message, state, from_d, to_d)


@dp.callback_query(F.data == 'custom')
async def custom(c: CallbackQuery, state: FSMContext):
    await state.set_state(ReportState.custom_dates)
    await c.message.edit_text('Введите период в формате:\n01.09.2026 - 12.09.2026\n\nОбе даты включаются.')
    await c.answer()


@dp.message(ReportState.custom_dates)
async def custom_value(m: Message, state: FSMContext):
    parts = re.split(r'\s+[—–-]\s+', (m.text or '').strip())
    try:
        if len(parts) != 2:
            raise ValueError
        from_d = datetime.strptime(parts[0], '%d.%m.%Y').date()
        to_d = datetime.strptime(parts[1], '%d.%m.%Y').date()
        if from_d > to_d:
            raise ValueError
        if to_d >= date.today():
            await m.answer(f'Конечная дата должна быть не позднее {(date.today()-timedelta(days=1)):%d.%m.%Y}.')
            return
        if (to_d - from_d).days + 1 > 90:
            await m.answer('Максимальный период — 90 дней.')
            return
    except ValueError:
        await m.answer('Неверный формат. Пример: 01.09.2026 - 12.09.2026')
        return
    await generate(m, state, from_d, to_d)


async def cabinet_data(token: str, cabinet_name: str, from_d: date, to_d: date):
    client = WBClient(token)
    cards_task = asyncio.create_task(client.get_all_cards())
    orders_task = asyncio.create_task(client.get_orders(from_d, to_d))
    cards, (orders, source) = await asyncio.gather(cards_task, orders_task)
    products = build_products(cards, cabinet_name)
    unmatched = apply_orders(products, orders, source)
    return products, source, unmatched


async def generate(m: Message, state: FSMContext, from_d: date, to_d: date):
    data = await state.get_data()
    cabinet = data.get('cabinet')
    status = await m.answer('⏳ Получаю товары и заказы Wildberries…')
    try:
        if cabinet == '1':
            products, src, unmatched = await cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d)
            name = CABINET_1_NAME
            sources = {src}
        elif cabinet == '2':
            products, src, unmatched = await cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d)
            name = CABINET_2_NAME
            sources = {src}
        else:
            r1, r2 = await asyncio.gather(
                cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d),
                cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d),
            )
            products = merge_products(r1[0], r2[0])
            sources = {r1[1], r2[1]}
            unmatched = r1[2] + r2[2]
            name = 'Оба кабинета'

        filename = f'WB_orders_{from_d:%Y-%m-%d}_{to_d:%Y-%m-%d}.xlsx'
        path = os.path.join(tempfile.gettempdir(), filename)
        make_excel(products, path, name, f'{from_d:%d.%m.%Y} — {to_d:%d.%m.%Y}')
        caption = (
            f'✅ Отчёт готов\n\n🏢 {name}\n📅 {from_d:%d.%m.%Y} — {to_d:%d.%m.%Y}\n'
            f'📦 Строк товаров/размеров: {len(products)}\n'
            f'🔎 Не сопоставлено заказов: {unmatched}'
        )
        if 'order-feed' in sources:
            caption += '\n⚠️ Использован резервный Order Feed (доступен максимум за 31 день).'
        await m.answer_document(FSInputFile(path, filename=filename), caption=caption)
        try:
            await status.delete()
        except Exception:
            pass
        try:
            os.remove(path)
        except OSError:
            pass
    except WBApiError as e:
        await status.edit_text(f'❌ Ошибка WB API:\n{e}')
    except Exception as e:
        await status.edit_text(f'❌ Ошибка: {type(e).__name__}: {e}')
    finally:
        await state.clear()


async def main():
    missing = []
    if not TELEGRAM_BOT_TOKEN: missing.append('TELEGRAM_BOT_TOKEN')
    if not WB_API_KEY_1: missing.append('WB_API_KEY_1')
    if not WB_API_KEY_2: missing.append('WB_API_KEY_2')
    if not ALLOWED_TELEGRAM_IDS: missing.append('ALLOWED_TELEGRAM_IDS')
    if missing:
        raise RuntimeError('Не заполнены переменные .env: ' + ', '.join(missing))
    await dp.start_polling(bot)


if __name__ == '__main__':
    asyncio.run(main())
