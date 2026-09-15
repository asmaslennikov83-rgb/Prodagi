import asyncio
import os
import re
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

from aiogram import BaseMiddleware, Bot, Dispatcher, F
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
from procurement import (
    calculate_procurement, load_bundles_file, load_stock_file, make_procurement_excel,
)
from report_builder import apply_orders, build_products, merge_products
from wb_client import WBApiError, WBClient


class PurchaseState(StatesGroup):
    custom_dates = State()
    stock_file = State()
    bundles_file = State()


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
    kb.button(text='📦 Рассчитать закупку', callback_data='purchase')
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


def stock_days_kb():
    kb = InlineKeyboardBuilder()
    for weeks in (1, 2, 3, 4):
        days = weeks * 7
        kb.button(text=f'{weeks} нед. ({days} дн.)', callback_data=f'stockdays:{days}')
    kb.adjust(2)
    return kb.as_markup()


def coefficient_kb():
    kb = InlineKeyboardBuilder()
    for label, value in [('0,5', '0.5'), ('0,75', '0.75'), ('1', '1'), ('1,25', '1.25'), ('1,5', '1.5')]:
        kb.button(text=label, callback_data=f'coef:{value}')
    kb.adjust(3, 2)
    return kb.as_markup()


@dp.message(CommandStart())
async def start(m: Message):
    await m.answer(
        '📦 Бот расчёта закупок Wildberries\n\n'
        'Учитывает заказы FBO + FBS, остатки на вашем складе и комплекты.',
        reply_markup=start_kb(),
    )


@dp.callback_query(F.data == 'purchase')
async def choose_cab(c: CallbackQuery, state: FSMContext):
    await cleanup_state_files(state)
    await state.clear()
    await c.message.edit_text('Выберите кабинет:', reply_markup=cabinets_kb())
    await c.answer()


@dp.callback_query(F.data.startswith('cab:'))
async def choose_period(c: CallbackQuery, state: FSMContext):
    await state.update_data(cabinet=c.data.split(':', 1)[1])
    await c.message.edit_text(
        'Выберите период продаж для расчёта средней скорости.\n'
        'Сегодняшний день в быстрые периоды не входит:',
        reply_markup=periods_kb(),
    )
    await c.answer()


@dp.callback_query(F.data.startswith('days:'))
async def quick_period(c: CallbackQuery, state: FSMContext):
    n = int(c.data.split(':', 1)[1])
    to_d = date.today() - timedelta(days=1)
    from_d = to_d - timedelta(days=n - 1)
    await state.update_data(from_d=from_d.isoformat(), to_d=to_d.isoformat(), analysis_days=n)
    await c.message.edit_text('На сколько дней должен хватать товарный запас?', reply_markup=stock_days_kb())
    await c.answer()


@dp.callback_query(F.data == 'custom')
async def custom(c: CallbackQuery, state: FSMContext):
    await state.set_state(PurchaseState.custom_dates)
    await c.message.edit_text('Введите период в формате:\n01.09.2026 - 12.09.2026\n\nОбе даты включаются.')
    await c.answer()


@dp.message(PurchaseState.custom_dates)
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
        analysis_days = (to_d - from_d).days + 1
        if analysis_days > 90:
            await m.answer('Максимальный период анализа — 90 дней.')
            return
    except ValueError:
        await m.answer('Неверный формат. Пример: 01.09.2026 - 12.09.2026')
        return
    await state.update_data(from_d=from_d.isoformat(), to_d=to_d.isoformat(), analysis_days=analysis_days)
    await state.set_state(None)
    await m.answer('На сколько дней должен хватать товарный запас?', reply_markup=stock_days_kb())


@dp.callback_query(F.data.startswith('stockdays:'))
async def choose_stock_days(c: CallbackQuery, state: FSMContext):
    target_days = int(c.data.split(':', 1)[1])
    await state.update_data(target_days=target_days)
    await c.message.edit_text('Выберите коэффициент закупки:', reply_markup=coefficient_kb())
    await c.answer()


@dp.callback_query(F.data.startswith('coef:'))
async def choose_coefficient(c: CallbackQuery, state: FSMContext):
    coefficient = float(c.data.split(':', 1)[1])
    await state.update_data(coefficient=coefficient)
    await state.set_state(PurchaseState.stock_file)
    await c.message.edit_text(
        '📎 Отправьте файл остатков на вашем складе.\n\n'
        'Поддерживаются <b>.xls</b> и <b>.xlsx</b>.\n'
        'В файле должны быть колонки <b>Код</b> и <b>Доступно</b>.',
        parse_mode='HTML',
    )
    await c.answer()


async def save_uploaded_excel(m: Message, prefix: str) -> str | None:
    if not m.document:
        return None
    name = m.document.file_name or ''
    suffix = Path(name).suffix.casefold()
    if suffix not in {'.xls', '.xlsx'}:
        return None
    fd, path = tempfile.mkstemp(prefix=prefix, suffix=suffix)
    os.close(fd)
    await bot.download(m.document, destination=path)
    return path


@dp.message(PurchaseState.stock_file)
async def receive_stock_file(m: Message, state: FSMContext):
    path = await save_uploaded_excel(m, 'wb_stock_')
    if not path:
        await m.answer('❌ Отправьте файл остатков в формате .xls или .xlsx.')
        return
    try:
        stocks, errors = load_stock_file(path)
    except Exception as exc:
        try: os.remove(path)
        except OSError: pass
        await m.answer(f'❌ Не удалось прочитать файл остатков:\n{exc}')
        return
    if not stocks:
        try: os.remove(path)
        except OSError: pass
        await m.answer('❌ В файле остатков не найдено ни одного товара.')
        return
    await state.update_data(stock_path=path, stock_parse_errors=errors, stock_rows=len(stocks))
    await state.set_state(PurchaseState.bundles_file)
    await m.answer(
        f'✅ Остатки приняты: {len(stocks)} уникальных баркодов.\n\n'
        'Теперь отправьте <b>шаблон комплектов</b> (.xls или .xlsx).',
        parse_mode='HTML',
    )


@dp.message(PurchaseState.bundles_file)
async def receive_bundles_file(m: Message, state: FSMContext):
    bundle_path = await save_uploaded_excel(m, 'wb_bundles_')
    if not bundle_path:
        await m.answer('❌ Отправьте шаблон комплектов в формате .xls или .xlsx.')
        return
    try:
        bundles, bundle_errors = load_bundles_file(bundle_path)
    except Exception as exc:
        try: os.remove(bundle_path)
        except OSError: pass
        await m.answer(f'❌ Не удалось прочитать шаблон комплектов:\n{exc}')
        return
    if not bundles:
        try: os.remove(bundle_path)
        except OSError: pass
        await m.answer('❌ В шаблоне не найдено ни одного комплекта.')
        return
    await state.update_data(bundle_path=bundle_path, bundle_parse_errors=bundle_errors)
    await generate_purchase(m, state)


async def cabinet_data(token: str, cabinet_name: str, from_d: date, to_d: date):
    client = WBClient(token)
    cards_task = asyncio.create_task(client.get_all_cards())
    orders_task = asyncio.create_task(client.get_orders(from_d, to_d))
    cards, (orders, source) = await asyncio.gather(cards_task, orders_task)
    products = build_products(cards, cabinet_name)
    unmatched = apply_orders(products, orders, source)
    return products, source, unmatched


async def generate_purchase(m: Message, state: FSMContext):
    data = await state.get_data()
    status = await m.answer('⏳ Получаю заказы FBO/FBS и рассчитываю закупку…')
    output_path = None
    try:
        from_d = date.fromisoformat(data['from_d'])
        to_d = date.fromisoformat(data['to_d'])
        analysis_days = int(data['analysis_days'])
        target_days = int(data['target_days'])
        coefficient = float(data['coefficient'])
        cabinet = data['cabinet']

        stocks, stock_errors_now = load_stock_file(data['stock_path'])
        bundles, bundle_errors_now = load_bundles_file(data['bundle_path'])
        initial_errors = list(data.get('stock_parse_errors') or []) + list(data.get('bundle_parse_errors') or [])
        # Avoid duplicates if the file was parsed twice.
        initial_errors += [e for e in stock_errors_now if e not in initial_errors]
        initial_errors += [e for e in bundle_errors_now if e not in initial_errors]

        if cabinet == '1':
            products, src, unmatched = await cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d)
            cabinet_name = CABINET_1_NAME
            sources = {src}
        elif cabinet == '2':
            products, src, unmatched = await cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d)
            cabinet_name = CABINET_2_NAME
            sources = {src}
        else:
            r1, r2 = await asyncio.gather(
                cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d),
                cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d),
            )
            products = merge_products(r1[0], r2[0])
            sources = {r1[1], r2[1]}
            unmatched = r1[2] + r2[2]
            cabinet_name = 'Оба кабинета'

        result = calculate_procurement(
            products=products,
            stocks=stocks,
            bundles=bundles,
            analysis_days=analysis_days,
            target_days=target_days,
            coefficient=coefficient,
            initial_errors=initial_errors,
        )

        filename = f'WB_purchase_{from_d:%Y-%m-%d}_{to_d:%Y-%m-%d}.xlsx'
        output_path = os.path.join(tempfile.gettempdir(), filename)
        make_procurement_excel(
            result, output_path, cabinet_name,
            f'{from_d:%d.%m.%Y} — {to_d:%d.%m.%Y}',
            analysis_days, target_days, coefficient,
        )
        purchase_total = sum(r.purchase_qty for r in result.rows)
        purchase_positions = sum(1 for r in result.rows if r.purchase_qty > 0)
        caption = (
            f'✅ Расчёт закупки готов\n\n'
            f'🏢 {cabinet_name}\n'
            f'📅 Продажи: {from_d:%d.%m.%Y} — {to_d:%d.%m.%Y}\n'
            f'📦 Запас: {target_days} дней\n'
            f'✖️ Коэффициент: {coefficient:g}\n'
            f'🛒 К закупке: {purchase_total} шт. / {purchase_positions} позиций\n'
            f'🔎 Не сопоставлено заказов WB: {unmatched}\n'
            f'⚠️ Записей на листе «Ошибки»: {len(result.errors)}'
        )
        if 'order-feed' in sources:
            caption += '\n⚠️ Для части данных использован резервный Order Feed.'
        await m.answer_document(FSInputFile(output_path, filename=filename), caption=caption)
        try:
            await status.delete()
        except Exception:
            pass
    except WBApiError as exc:
        await status.edit_text(f'❌ Ошибка WB API:\n{exc}')
    except Exception as exc:
        await status.edit_text(f'❌ Ошибка: {type(exc).__name__}: {exc}')
    finally:
        if output_path:
            try: os.remove(output_path)
            except OSError: pass
        await cleanup_state_files(state)
        await state.clear()


async def cleanup_state_files(state: FSMContext):
    try:
        data = await state.get_data()
    except Exception:
        return
    for key in ('stock_path', 'bundle_path'):
        path = data.get(key)
        if path:
            try:
                os.remove(path)
            except OSError:
                pass


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
