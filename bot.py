import asyncio
import html
import os
import re
import shutil
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


class ManualState(StatesGroup):
    barcode = State()


class TemplateState(StatesGroup):
    bundles_file = State()


bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

DATA_DIR = Path(os.getenv('BOT_DATA_DIR', 'data'))
DATA_DIR.mkdir(parents=True, exist_ok=True)


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
    kb.button(text='🔎 Ручная проверка по ШК', callback_data='manual_check')
    kb.button(text='📦 Загрузить шаблон комплектов', callback_data='upload_bundles_template')
    kb.adjust(1)
    return kb.as_markup()


def _last_file_path(user_id: int, kind: str) -> Path | None:
    for suffix in ('.xlsx', '.xls'):
        p = DATA_DIR / f'{user_id}_{kind}{suffix}'
        if p.exists():
            return p
    return None


def _save_last_file(user_id: int, kind: str, source_path: str) -> Path:
    suffix = Path(source_path).suffix.casefold()
    for old_suffix in ('.xlsx', '.xls'):
        old = DATA_DIR / f'{user_id}_{kind}{old_suffix}'
        if old.exists() and old.suffix.casefold() != suffix:
            try:
                old.unlink()
            except OSError:
                pass
    dest = DATA_DIR / f'{user_id}_{kind}{suffix}'
    shutil.copy2(source_path, dest)
    return dest


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


def stock_mode_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text='🏢 Только FBO', callback_data='stockmode:fbo')
    kb.button(text='📦 Только FBS', callback_data='stockmode:fbs')
    kb.button(text='➕ FBO + FBS', callback_data='stockmode:both')
    kb.adjust(1)
    return kb.as_markup()


@dp.message(CommandStart())
async def start(m: Message):
    await m.answer(
        '📦 Бот расчёта закупок Wildberries\n\n'
        'Учитывает заказы FBO + FBS, остатки на вашем складе и комплекты.',
        reply_markup=start_kb(),
    )




@dp.callback_query(F.data == 'upload_bundles_template')
async def upload_bundles_template_start(c: CallbackQuery, state: FSMContext):
    await cleanup_state_files(state)
    await state.clear()
    await state.set_state(TemplateState.bundles_file)
    current = _last_file_path(c.from_user.id, 'bundles')
    current_text = (
        '\n\n✅ Сейчас сохранён шаблон: <code>' + html.escape(current.name) + '</code>'
        if current else
        '\n\n⚠️ Сейчас сохранённого шаблона нет.'
    )
    await c.message.edit_text(
        '📦 <b>Загрузка шаблона комплектов</b>\n\n'
        'Отправьте файл <b>.xls</b> или <b>.xlsx</b>. Новый файл заменит предыдущий шаблон для вашего Telegram ID.'
        + current_text,
        parse_mode='HTML',
    )
    await c.answer()


@dp.message(TemplateState.bundles_file)
async def upload_bundles_template_file(m: Message, state: FSMContext):
    bundle_path = await save_uploaded_excel(m, 'wb_bundles_manual_')
    if not bundle_path:
        await m.answer('❌ Отправьте шаблон комплектов в формате .xls или .xlsx.')
        return
    try:
        bundles, bundle_errors = load_bundles_file(bundle_path)
    except Exception as exc:
        try:
            os.remove(bundle_path)
        except OSError:
            pass
        await m.answer(f'❌ Не удалось прочитать шаблон комплектов:\n{exc}')
        return
    if not bundles:
        try:
            os.remove(bundle_path)
        except OSError:
            pass
        await m.answer('❌ В шаблоне не найдено ни одного комплекта.')
        return
    saved = _save_last_file(m.from_user.id, 'bundles', bundle_path)
    try:
        os.remove(bundle_path)
    except OSError:
        pass
    await state.clear()
    text = f'✅ Шаблон комплектов сохранён. Найдено комплектов: {len(bundles)}.'
    if bundle_errors:
        text += f'\n⚠️ Предупреждений при чтении: {len(bundle_errors)}.'
    text += '\n\nТеперь его можно использовать в «🔎 Ручной проверке по ШК» без полноценного отчёта.'
    await m.answer(text, reply_markup=start_kb())


@dp.callback_query(F.data == 'manual_check')
async def manual_check_start(c: CallbackQuery, state: FSMContext):
    await cleanup_state_files(state)
    await state.clear()
    await state.update_data(flow='manual', cabinet='both')
    await state.set_state(ManualState.barcode)
    await c.message.edit_text(
        '🔎 <b>Ручная проверка по ШК</b>\n\n'
        'Введите один баркод товара:',
        parse_mode='HTML',
    )
    await c.answer()


@dp.message(ManualState.barcode)
async def manual_barcode_value(m: Message, state: FSMContext):
    barcode = re.sub(r'\s+', '', (m.text or '').strip())
    if not barcode or not re.fullmatch(r'\d{5,30}', barcode):
        await m.answer('❌ Введите один числовой баркод без пробелов.')
        return
    await state.update_data(flow='manual', cabinet='both', manual_barcode=barcode)
    await state.set_state(None)
    await m.answer(
        'Выберите период продаж для расчёта средней скорости.\n'
        'Сегодняшний день в быстрые периоды не входит:',
        reply_markup=periods_kb(),
    )


@dp.callback_query(F.data == 'purchase')
async def choose_cab(c: CallbackQuery, state: FSMContext):
    await cleanup_state_files(state)
    await state.clear()
    await state.update_data(flow='purchase')
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
    await c.message.edit_text(
        'Какие текущие остатки учитывать при расчёте закупки?',
        reply_markup=stock_mode_kb(),
    )
    await c.answer()


@dp.callback_query(F.data.startswith('stockmode:'))
async def choose_stock_mode(c: CallbackQuery, state: FSMContext):
    stock_mode = c.data.split(':', 1)[1]
    await state.update_data(stock_mode=stock_mode)
    data = await state.get_data()
    if data.get('flow') == 'manual':
        await c.answer()
        await generate_manual_check(c.message, state, c.from_user.id)
        return
    if stock_mode == 'fbo':
        # Local/FBS stock file is intentionally skipped in FBO-only mode.
        await state.update_data(stock_path=None, stock_parse_errors=[], stock_rows=0)
        await state.set_state(PurchaseState.bundles_file)
        await c.message.edit_text(
            '✅ Будут учтены только текущие остатки FBO на складах Wildberries.\n\n'
            'Теперь отправьте <b>шаблон комплектов</b> (.xls или .xlsx).',
            parse_mode='HTML',
        )
    else:
        await state.set_state(PurchaseState.stock_file)
        label = 'только FBS' if stock_mode == 'fbs' else 'FBO + FBS'
        await c.message.edit_text(
            f'Выбрано: <b>{label}</b>.\n\n'
            '📎 Отправьте файл остатков на вашем складе (FBS).\n'
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
    _save_last_file(m.from_user.id, 'stock', path)
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
    _save_last_file(m.from_user.id, 'bundles', bundle_path)
    await state.update_data(bundle_path=bundle_path, bundle_parse_errors=bundle_errors)
    await generate_purchase(m, state)


async def cabinet_data(token: str, cabinet_name: str, from_d: date, to_d: date, need_fbo_stock: bool = False, need_fbs_stock: bool = False):
    client = WBClient(token)
    cards_task = asyncio.create_task(client.get_all_cards())
    orders_task = asyncio.create_task(client.get_orders(from_d, to_d))
    cards, (orders, source) = await asyncio.gather(cards_task, orders_task)
    products = build_products(cards, cabinet_name)
    unmatched = apply_orders(products, orders, source)
    if need_fbo_stock:
        nm_ids = [p.nm_id for p in products if p.nm_id]
        fbo_by_chrt = await client.get_fbo_stocks(nm_ids)
        for p in products:
            if p.chrt_id:
                p.fbo_stock = fbo_by_chrt.get(p.chrt_id, 0)
    if need_fbs_stock:
        chrt_ids = [p.chrt_id for p in products if p.chrt_id]
        fbs_by_chrt = await client.get_fbs_stocks(chrt_ids)
        for p in products:
            if p.chrt_id:
                p.fbs_stock = fbs_by_chrt.get(p.chrt_id, 0)
    return products, source, unmatched


def _find_product_by_barcode(products, barcode: str):
    return next((p for p in products if barcode in p.barcodes), None)


def _find_row_for_product(result, product):
    if product is None:
        return None
    return next((r for r in result.rows if r.product.barcodes & product.barcodes), None)


def _fmt_num(value: float, digits: int = 2) -> str:
    text = f'{value:.{digits}f}'.rstrip('0').rstrip('.')
    return text.replace('.', ',')


async def generate_manual_check(m: Message, state: FSMContext, user_id: int):
    data = await state.get_data()
    barcode = data.get('manual_barcode', '')
    stock_mode = data.get('stock_mode', 'fbs')
    bundle_path = _last_file_path(user_id, 'bundles')

    if bundle_path is None:
        await m.answer(
            '❌ Нет сохранённого шаблона комплектов.\n\n'
            'На главном экране нажмите «📦 Загрузить шаблон комплектов», загрузите файл и повторите проверку.'
        )
        await state.clear()
        return
    status = await m.answer(f'⏳ Проверяю ШК {barcode} по двум кабинетам…')
    try:
        from_d = date.fromisoformat(data['from_d'])
        to_d = date.fromisoformat(data['to_d'])
        analysis_days = int(data['analysis_days'])
        target_days = int(data['target_days'])
        coefficient = float(data['coefficient'])

        # Для ручной проверки FBS-остатки всегда получаем напрямую из WB API.
        stocks = {}
        bundles, _ = load_bundles_file(bundle_path)

        r1, r2 = await asyncio.gather(
            cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d, True, True),
            cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d, True, True),
        )
        products1, products2 = r1[0], r2[0]
        merged = merge_products(products1, products2)

        # Bundle SKU itself is not purchased; explain this explicitly.
        bundle = next((b for b in bundles if b.barcode == barcode), None)
        if bundle:
            p1 = _find_product_by_barcode(products1, barcode)
            p2 = _find_product_by_barcode(products2, barcode)
            fbo1, fbs1 = (p1.fbo, p1.fbs) if p1 else (0, 0)
            fbo2, fbs2 = (p2.fbo, p2.fbs) if p2 else (0, 0)
            components = ', '.join(bundle.components)
            await status.edit_text(
                f'🧩 <b>{html.escape(bundle.name)}</b>\n'
                f'ШК: <code>{barcode}</code>\n\n'
                f'<b>{html.escape(CABINET_1_NAME)}</b>: FBO {fbo1} / FBS {fbs1} / всего {fbo1+fbs1}\n'
                f'<b>{html.escape(CABINET_2_NAME)}</b>: FBO {fbo2} / FBS {fbs2} / всего {fbo2+fbs2}\n'
                f'<b>Всего заказов набора:</b> {fbo1+fbs1+fbo2+fbs2}\n'
                f'<b>Компоненты:</b> {html.escape(components)}\n\n'
                'Этот ШК является комплектом. Сам комплект не закупается — закупка рассчитывается по его компонентам.',
                parse_mode='HTML',
            )
            await state.clear()
            return

        merged_product = _find_product_by_barcode(merged, barcode)
        if merged_product is None:
            await status.edit_text(f'❌ ШК <code>{barcode}</code> не найден ни в одном кабинете WB.', parse_mode='HTML')
            await state.clear()
            return

        # Calculate the final merged procurement row with the selected stock mode.
        merged_result = calculate_procurement(
            products=merged,
            stocks_fbs=stocks,
            bundles=bundles,
            analysis_days=analysis_days,
            target_days=target_days,
            coefficient=coefficient,
            stock_mode=stock_mode,
        )
        row = _find_row_for_product(merged_result, merged_product)
        if row is None:
            await status.edit_text('❌ Товар найден, но он определён как комплект и не входит в закупку компонентов.')
            await state.clear()
            return

        # Separate cabinet calculations are used only to show sales and bundle consumption per cabinet.
        res1 = calculate_procurement(products1, {}, bundles, analysis_days, target_days, coefficient, 'fbo')
        res2 = calculate_procurement(products2, {}, bundles, analysis_days, target_days, coefficient, 'fbo')
        p1 = next((p for p in products1 if p.barcodes & merged_product.barcodes or (p.vendor_code and p.vendor_code == merged_product.vendor_code and p.size == merged_product.size)), None)
        p2 = next((p for p in products2 if p.barcodes & merged_product.barcodes or (p.vendor_code and p.vendor_code == merged_product.vendor_code and p.size == merged_product.size)), None)
        rr1 = _find_row_for_product(res1, p1)
        rr2 = _find_row_for_product(res2, p2)

        fbo1, fbs1 = (p1.fbo, p1.fbs) if p1 else (0, 0)
        fbo2, fbs2 = (p2.fbo, p2.fbs) if p2 else (0, 0)
        bundle1 = rr1.bundle_sales_units if rr1 else 0
        bundle2 = rr2.bundle_sales_units if rr2 else 0
        direct1, direct2 = fbo1 + fbs1, fbo2 + fbs2
        total_direct = direct1 + direct2
        total_bundle = bundle1 + bundle2
        mode_label = {'fbo': 'только FBO', 'fbs': 'только FBS', 'both': 'FBO + FBS'}.get(stock_mode, stock_mode)

        barcodes_text = ', '.join(sorted(merged_product.barcodes))
        text = (
            f'🔎 <b>Ручная проверка по ШК</b>\n\n'
            f'<b>ШК:</b> <code>{barcode}</code>\n'
            f'<b>Все ШК товара:</b> {barcodes_text}\n'
            f'<b>Артикул продавца:</b> {html.escape(merged_product.vendor_code or "—")}\n'
            f'<b>Бренд:</b> {html.escape(merged_product.brand or "—")}\n'
            f'<b>Товар:</b> {html.escape(merged_product.name or "—")}\n'
            f'<b>Размер:</b> {html.escape(merged_product.size or "0")}\n\n'
            f'📅 <b>Период:</b> {from_d:%d.%m.%Y} — {to_d:%d.%m.%Y} ({analysis_days} дн.)\n\n'
            f'🏢 <b>{html.escape(CABINET_1_NAME)}</b>\n'
            f'FBO: {fbo1} | FBS: {fbs1} | напрямую: {direct1}\n'
            f'Через наборы: {bundle1} | общий расход: {direct1 + bundle1}\n\n'
            f'🏢 <b>{html.escape(CABINET_2_NAME)}</b>\n'
            f'FBO: {fbo2} | FBS: {fbs2} | напрямую: {direct2}\n'
            f'Через наборы: {bundle2} | общий расход: {direct2 + bundle2}\n\n'
            f'📊 <b>Итого</b>\n'
            f'Заказы товара напрямую: {total_direct}\n'
            f'Расход через наборы: {total_bundle}\n'
            f'<b>Общий расход: {row.total_consumption}</b>\n'
            f'Среднее: <b>{_fmt_num(row.average_per_day)} шт./день</b>\n\n'
            f'📦 <b>Остатки</b>\n'
            f'FBO: {row.fbo_physical_stock} шт.\n'
            f'FBS (WB API): {row.fbs_physical_stock} шт.\n'
            f'Для расчёта ({mode_label}): <b>{row.physical_stock} шт.</b>\n'
            f'Хватит примерно на: <b>{_fmt_num(row.days_remaining, 1)} дн.</b>\n\n'
            f'🛒 <b>Закупка</b>\n'
            f'Целевой запас: {target_days} дн.\n'
            f'Коэффициент: {str(coefficient).replace(".", ",")}\n'
            f'Необходимо иметь: {row.required_stock} шт.\n'
            f'<b>К закупке: {row.purchase_qty} шт.</b>'
        )
        await status.edit_text(text, parse_mode='HTML')
    except WBApiError as exc:
        await status.edit_text(f'❌ Ошибка WB API:\n{exc}')
    except Exception as exc:
        await status.edit_text(f'❌ Ошибка: {type(exc).__name__}: {exc}')
    finally:
        await state.clear()


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
        stock_mode = data.get('stock_mode', 'fbs')
        cabinet = data['cabinet']
        need_fbo_stock = stock_mode in {'fbo', 'both'}

        if stock_mode in {'fbs', 'both'}:
            stocks, stock_errors_now = load_stock_file(data['stock_path'])
        else:
            stocks, stock_errors_now = {}, []
        bundles, bundle_errors_now = load_bundles_file(data['bundle_path'])
        initial_errors = list(data.get('stock_parse_errors') or []) + list(data.get('bundle_parse_errors') or [])
        # Avoid duplicates if the file was parsed twice.
        initial_errors += [e for e in stock_errors_now if e not in initial_errors]
        initial_errors += [e for e in bundle_errors_now if e not in initial_errors]

        if cabinet == '1':
            products, src, unmatched = await cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d, need_fbo_stock)
            cabinet_name = CABINET_1_NAME
            sources = {src}
        elif cabinet == '2':
            products, src, unmatched = await cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d, need_fbo_stock)
            cabinet_name = CABINET_2_NAME
            sources = {src}
        else:
            r1, r2 = await asyncio.gather(
                cabinet_data(WB_API_KEY_1, CABINET_1_NAME, from_d, to_d, need_fbo_stock),
                cabinet_data(WB_API_KEY_2, CABINET_2_NAME, from_d, to_d, need_fbo_stock),
            )
            products = merge_products(r1[0], r2[0])
            sources = {r1[1], r2[1]}
            unmatched = r1[2] + r2[2]
            cabinet_name = 'Оба кабинета'

        result = calculate_procurement(
            products=products,
            stocks_fbs=stocks,
            bundles=bundles,
            analysis_days=analysis_days,
            target_days=target_days,
            coefficient=coefficient,
            stock_mode=stock_mode,
            initial_errors=initial_errors,
        )

        filename = f'WB_purchase_{from_d:%Y-%m-%d}_{to_d:%Y-%m-%d}.xlsx'
        output_path = os.path.join(tempfile.gettempdir(), filename)
        make_procurement_excel(
            result, output_path, cabinet_name,
            f'{from_d:%d.%m.%Y} — {to_d:%d.%m.%Y}',
            analysis_days, target_days, coefficient, stock_mode,
        )
        purchase_total = sum(r.purchase_qty for r in result.rows)
        purchase_positions = sum(1 for r in result.rows if r.purchase_qty > 0)
        caption = (
            f'✅ Расчёт закупки готов\n\n'
            f'🏢 {cabinet_name}\n'
            f'📅 Продажи: {from_d:%d.%m.%Y} — {to_d:%d.%m.%Y}\n'
            f'📦 Запас: {target_days} дней\n'
            f'✖️ Коэффициент: {coefficient:g}\n'
            f'📊 Остатки: { {"fbo": "только FBO", "fbs": "только FBS", "both": "FBO + FBS"}.get(stock_mode, stock_mode) }\n'
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
