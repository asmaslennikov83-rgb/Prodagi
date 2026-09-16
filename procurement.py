from __future__ import annotations

import math
import re
from decimal import Decimal, ROUND_CEILING
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from report_builder import Product


@dataclass
class Bundle:
    name: str
    barcode: str
    components: list[str]


@dataclass
class PurchaseRow:
    product: Product
    bundle_sales_units: int = 0
    fbo_direct_stock: int = 0
    fbs_direct_stock: int = 0
    fbo_stock_inside_bundles: int = 0
    fbs_stock_inside_bundles: int = 0
    selected_stock: int = 0
    average_per_day: float = 0.0
    required_stock: int = 0
    purchase_qty: int = 0
    note: str = ''

    @property
    def total_consumption(self) -> int:
        return self.product.fbo + self.product.fbs + self.bundle_sales_units

    @property
    def fbo_physical_stock(self) -> int:
        return self.fbo_direct_stock + self.fbo_stock_inside_bundles

    @property
    def fbs_physical_stock(self) -> int:
        return self.fbs_direct_stock + self.fbs_stock_inside_bundles

    @property
    def physical_stock(self) -> int:
        return self.selected_stock


@dataclass
class BundleAuditRow:
    bundle_name: str
    bundle_barcode: str
    fbo: int
    fbs: int
    bundle_orders: int
    bundle_stock_fbo: int
    bundle_stock_fbs: int
    bundle_stock_selected: int
    component_barcode: str
    component_name: str
    qty_in_bundle: int
    component_sales_consumption: int
    component_stock_inside_bundles_fbo: int
    component_stock_inside_bundles_fbs: int
    component_stock_inside_bundles_selected: int


@dataclass
class ProcurementResult:
    rows: list[PurchaseRow]
    bundle_rows: list[BundleAuditRow]
    errors: list[tuple[str, str, str]] = field(default_factory=list)


def _norm_header(value) -> str:
    text = '' if value is None else str(value)
    text = text.strip().casefold().replace('ё', 'е')
    return re.sub(r'[^a-zа-я0-9]+', '', text)


def barcode_text(value) -> str:
    if value is None:
        return ''
    if isinstance(value, bool):
        return ''
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format(value, '.15g').strip()
    text = str(value).strip()
    if not text:
        return ''
    if re.fullmatch(r'\d+\.0+', text):
        return text.split('.', 1)[0]
    return re.sub(r'\s+', '', text)


def int_qty(value) -> int:
    if value in (None, ''):
        return 0
    if isinstance(value, str):
        value = value.replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def _read_xlsx_rows(path: str | Path) -> list[list]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        return [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _read_xls_rows(path: str | Path) -> list[list]:
    try:
        import xlrd  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            'Для чтения .xls не установлен пакет xlrd. Выполните: pip install -r requirements.txt'
        ) from exc
    book = xlrd.open_workbook(str(path), on_demand=True)
    try:
        sh = book.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    finally:
        book.release_resources()


def read_table_rows(path: str | Path) -> list[list]:
    suffix = Path(path).suffix.casefold()
    if suffix == '.xlsx':
        return _read_xlsx_rows(path)
    if suffix == '.xls':
        return _read_xls_rows(path)
    raise ValueError('Поддерживаются только файлы .xls и .xlsx')


def load_stock_file(path: str | Path) -> tuple[dict[str, int], list[tuple[str, str, str]]]:
    rows = read_table_rows(path)
    errors: list[tuple[str, str, str]] = []
    header_idx = code_col = qty_col = None

    for r_idx, row in enumerate(rows[:30]):
        headers = [_norm_header(x) for x in row]
        for c_idx, h in enumerate(headers):
            if h in {'код', 'баркод', 'barcode', 'sku'}:
                code_col = c_idx
            if h in {'доступно', 'количество', 'остаток', 'qty', 'quantity'}:
                qty_col = c_idx
        if code_col is not None and qty_col is not None:
            header_idx = r_idx
            break
        code_col = qty_col = None

    if header_idx is None or code_col is None or qty_col is None:
        raise ValueError('В файле остатков не найдены колонки «Код» и «Доступно».')

    stocks: dict[str, int] = defaultdict(int)
    for row_num, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if code_col >= len(row):
            continue
        bc = barcode_text(row[code_col])
        if not bc:
            continue
        qty = int_qty(row[qty_col] if qty_col < len(row) else 0)
        stocks[bc] += qty
        if qty < 0:
            errors.append(('Остатки', bc, f'Строка {row_num}: отрицательный остаток {qty}.'))
    return dict(stocks), errors


def load_bundles_file(path: str | Path) -> tuple[list[Bundle], list[tuple[str, str, str]]]:
    rows = read_table_rows(path)
    errors: list[tuple[str, str, str]] = []
    if not rows:
        raise ValueError('Шаблон комплектов пуст.')

    header_idx = name_col = bundle_col = None
    component_cols: list[int] = []
    for r_idx, row in enumerate(rows[:30]):
        headers = [_norm_header(x) for x in row]
        name_col = next((i for i, h in enumerate(headers) if h in {'название', 'наименование'}), None)
        bundle_col = next((i for i, h in enumerate(headers) if 'баркод' in h and ('комплект' in h or 'набор' in h)), None)
        if bundle_col is None:
            bundle_col = next((i for i, h in enumerate(headers) if h in {'баркодкомплекта', 'баркоднабора'}), None)
        if bundle_col is not None:
            component_cols = [
                i for i, h in enumerate(headers)
                if i != bundle_col and (h.startswith('баркод') or h.startswith('barcode'))
            ]
        if bundle_col is not None and component_cols:
            header_idx = r_idx
            break

    if header_idx is None or bundle_col is None or not component_cols:
        raise ValueError(
            'В шаблоне комплектов нужны колонки «баркод комплекта» и «баркод1», «баркод2» ...'
        )

    bundles: list[Bundle] = []
    seen_bundle_barcodes: set[str] = set()
    for row_num, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        bundle_bc = barcode_text(row[bundle_col] if bundle_col < len(row) else None)
        if not bundle_bc:
            continue
        name = ''
        if name_col is not None and name_col < len(row):
            name = str(row[name_col] or '').strip()
        components = [
            barcode_text(row[c] if c < len(row) else None)
            for c in component_cols
        ]
        components = [x for x in components if x]
        if not components:
            errors.append(('Комплекты', bundle_bc, f'Строка {row_num}: у комплекта нет компонентов.'))
            continue
        if bundle_bc in seen_bundle_barcodes:
            errors.append(('Комплекты', bundle_bc, f'Строка {row_num}: повторный баркод комплекта.'))
            continue
        seen_bundle_barcodes.add(bundle_bc)
        bundles.append(Bundle(name=name or bundle_bc, barcode=bundle_bc, components=components))
    return bundles, errors


def _product_barcode_index(products: Iterable[Product]) -> dict[str, Product]:
    index: dict[str, Product] = {}
    for p in products:
        for bc in p.barcodes:
            if bc:
                index.setdefault(bc, p)
    return index


def calculate_procurement(
    products: list[Product],
    stocks_fbs: dict[str, int],
    bundles: list[Bundle],
    analysis_days: int,
    target_days: int,
    coefficient: float,
    stock_mode: str = 'fbs',
    initial_errors: list[tuple[str, str, str]] | None = None,
) -> ProcurementResult:
    errors = list(initial_errors or [])
    by_barcode = _product_barcode_index(products)
    bundle_barcodes = {b.barcode for b in bundles}

    # Products that are bundle SKUs are shown on the audit sheet, not in the purchase list.
    bundle_products = {id(by_barcode[bc]) for bc in bundle_barcodes if bc in by_barcode}

    fbs_direct_by_product: dict[int, int] = defaultdict(int)
    fbo_direct_by_product: dict[int, int] = defaultdict(int)
    for p in products:
        fbs_direct_by_product[id(p)] = sum(stocks_fbs.get(bc, 0) for bc in p.barcodes)
        fbo_direct_by_product[id(p)] = int(p.fbo_stock or 0)

    bundle_sales_by_product: dict[int, int] = defaultdict(int)
    fbs_stock_in_bundles_by_product: dict[int, int] = defaultdict(int)
    fbo_stock_in_bundles_by_product: dict[int, int] = defaultdict(int)
    bundle_rows: list[BundleAuditRow] = []

    def selected_stock(fbo: int, fbs: int) -> int:
        if stock_mode == 'fbo':
            return fbo
        if stock_mode == 'both':
            return fbo + fbs
        return fbs

    for bundle in bundles:
        bundle_product = by_barcode.get(bundle.barcode)
        if bundle_product is None:
            errors.append(('Комплекты', bundle.barcode, 'Баркод комплекта не найден среди товаров Wildberries.'))
            fbo = fbs = 0
            bundle_stock_fbo = 0
        else:
            fbo, fbs = bundle_product.fbo, bundle_product.fbs
            bundle_stock_fbo = int(bundle_product.fbo_stock or 0)
        bundle_orders = fbo + fbs
        bundle_stock_fbs = stocks_fbs.get(bundle.barcode, 0)
        bundle_stock_selected = selected_stock(bundle_stock_fbo, bundle_stock_fbs)

        matched_counts: dict[int, int] = defaultdict(int)
        matched_product: dict[int, Product] = {}
        representative_bc: dict[int, str] = {}
        unknown_counts: Counter[str] = Counter()
        for comp_bc in bundle.components:
            comp_product = by_barcode.get(comp_bc)
            if comp_product is None:
                unknown_counts[comp_bc] += 1
                continue
            pid = id(comp_product)
            matched_counts[pid] += 1
            matched_product[pid] = comp_product
            representative_bc.setdefault(pid, comp_bc)

        for comp_bc, mult in unknown_counts.items():
            errors.append((
                'Комплекты', comp_bc,
                f'Компонент не найден среди товаров WB. Комплект: {bundle.name} ({bundle.barcode}), кратность {mult}.'
            ))
            bundle_rows.append(BundleAuditRow(
                bundle.name, bundle.barcode, fbo, fbs, bundle_orders,
                bundle_stock_fbo, bundle_stock_fbs, bundle_stock_selected,
                comp_bc, 'НЕ НАЙДЕН', mult, bundle_orders * mult,
                bundle_stock_fbo * mult, bundle_stock_fbs * mult, bundle_stock_selected * mult,
            ))

        for pid, mult in matched_counts.items():
            p = matched_product[pid]
            if pid in bundle_products:
                errors.append((
                    'Комплекты', representative_bc[pid],
                    f'Компонент сам является комплектом ({p.name}). Вложенные комплекты не разворачиваются автоматически.'
                ))
            sales_use = bundle_orders * mult
            stock_use_fbo = bundle_stock_fbo * mult
            stock_use_fbs = bundle_stock_fbs * mult
            stock_use_selected = bundle_stock_selected * mult
            bundle_sales_by_product[pid] += sales_use
            fbo_stock_in_bundles_by_product[pid] += stock_use_fbo
            fbs_stock_in_bundles_by_product[pid] += stock_use_fbs
            bundle_rows.append(BundleAuditRow(
                bundle.name, bundle.barcode, fbo, fbs, bundle_orders,
                bundle_stock_fbo, bundle_stock_fbs, bundle_stock_selected,
                representative_bc[pid], p.name, mult, sales_use,
                stock_use_fbo, stock_use_fbs, stock_use_selected,
            ))

    rows: list[PurchaseRow] = []
    for p in products:
        if id(p) in bundle_products:
            continue
        row = PurchaseRow(product=p)
        row.bundle_sales_units = bundle_sales_by_product[id(p)]
        row.fbo_direct_stock = fbo_direct_by_product[id(p)]
        row.fbs_direct_stock = fbs_direct_by_product[id(p)]
        row.fbo_stock_inside_bundles = fbo_stock_in_bundles_by_product[id(p)]
        row.fbs_stock_inside_bundles = fbs_stock_in_bundles_by_product[id(p)]
        row.selected_stock = selected_stock(row.fbo_physical_stock, row.fbs_physical_stock)
        total = row.total_consumption
        row.average_per_day = total / analysis_days if analysis_days > 0 else 0.0
        raw_required = (Decimal(total) / Decimal(analysis_days) * Decimal(target_days) * Decimal(str(coefficient))) if analysis_days > 0 else Decimal(0)
        row.required_stock = int(raw_required.to_integral_value(rounding=ROUND_CEILING))
        row.purchase_qty = max(0, row.required_stock - row.physical_stock)
        if stock_mode in {'fbs', 'both'} and not any(bc in stocks_fbs for bc in p.barcodes) and row.fbs_stock_inside_bundles == 0:
            row.note = 'Нет в файле остатков FBS'
        rows.append(row)

    known_barcodes = set(by_barcode) | bundle_barcodes
    for bc, qty in stocks_fbs.items():
        if bc not in known_barcodes:
            errors.append(('Остатки', bc, f'Баркод из файла остатков не найден в WB/шаблоне комплектов. Остаток: {qty}.'))

    rows.sort(key=lambda r: (-r.purchase_qty, r.product.name.casefold(), r.product.size.casefold()))
    bundle_rows.sort(key=lambda r: (r.bundle_name.casefold(), r.component_name.casefold()))
    return ProcurementResult(rows=rows, bundle_rows=bundle_rows, errors=errors)


def make_procurement_excel(
    result: ProcurementResult,
    path: str | Path,
    cabinet_name: str,
    period_text: str,
    analysis_days: int,
    target_days: int,
    coefficient: float,
    stock_mode: str = 'fbs',
):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Закупка'

    navy = '1F4E78'
    blue = 'D9EAF7'
    green = 'E2F0D9'
    yellow = 'FFF2CC'
    red = 'FCE4D6'
    thin = Side(style='thin', color='D9E2F3')

    ws.merge_cells('A1:R1')
    ws['A1'] = 'Расчёт закупки Wildberries'
    ws['A1'].font = Font(bold=True, size=16, color='FFFFFF')
    ws['A1'].fill = PatternFill('solid', fgColor=navy)
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A2:R2')
    ws['A2'] = f'Кабинет: {cabinet_name} | Период продаж: {period_text} ({analysis_days} дн.)'
    ws.merge_cells('A3:R3')
    mode_label = {'fbo': 'Только FBO', 'fbs': 'Только FBS', 'both': 'FBO + FBS'}.get(stock_mode, stock_mode)
    ws['A3'] = f'Целевой запас: {target_days} дн. | Коэффициент: {coefficient:g} | Учитываемые остатки: {mode_label}'

    headers = [
        'Баркод(ы)', 'Артикул продавца', 'Бренд', 'Наименование', 'Размер', 'Заказы FBO', 'Заказы FBS',
        'Продажи комплектами', 'Общий расход', 'Среднее/день', 'Запас, дней',
        'Коэффициент', 'Необходимо иметь', 'Остаток FBO', 'Остаток FBS', 'Остаток учтён', 'К закупке', 'Примечание'
    ]
    header_row = 5
    for col, value in enumerate(headers, 1):
        c = ws.cell(header_row, col, value)
        c.fill = PatternFill('solid', fgColor=navy)
        c.font = Font(color='FFFFFF', bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for r_idx, r in enumerate(result.rows, start=header_row + 1):
        values = [
            ', '.join(sorted(r.product.barcodes)), r.product.vendor_code, r.product.brand, r.product.name, r.product.size or '0',
            r.product.fbo, r.product.fbs, r.bundle_sales_units, r.total_consumption,
            r.average_per_day, target_days, coefficient, r.required_stock,
            r.fbo_physical_stock, r.fbs_physical_stock, r.physical_stock, r.purchase_qty, r.note,
        ]
        for c_idx, value in enumerate(values, 1):
            cell = ws.cell(r_idx, c_idx, value)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical='center', wrap_text=c_idx in (1, 4, 18))
        ws.cell(r_idx, 10).number_format = '0.00'
        ws.cell(r_idx, 12).number_format = '0.00'
        if r.purchase_qty > 0:
            ws.cell(r_idx, 17).fill = PatternFill('solid', fgColor=green)
            ws.cell(r_idx, 17).font = Font(bold=True)
        if r.note:
            ws.cell(r_idx, 18).fill = PatternFill('solid', fgColor=yellow)

    widths = [32, 24, 20, 48, 12, 13, 13, 20, 14, 14, 12, 13, 18, 14, 14, 16, 14, 26]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = 'A6'
    if result.rows:
        ws.auto_filter.ref = f'A5:R{header_row + len(result.rows)}'
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[5].height = 36

    # Bundle audit sheet
    bs = wb.create_sheet('Комплекты')
    b_headers = [
        'Название', 'Баркод комплекта', 'Заказы FBO', 'Заказы FBS', 'Всего заказов',
        'Остаток комплектов FBO', 'Остаток комплектов FBS', 'Остаток комплектов учтён',
        'Баркод компонента', 'Наименование компонента', 'Кол-во в комплекте', 'Расход через продажи',
        'Компонентов в FBO-комплектах', 'Компонентов в FBS-комплектах', 'Компонентов учтено'
    ]
    for i, h in enumerate(b_headers, 1):
        c = bs.cell(1, i, h)
        c.fill = PatternFill('solid', fgColor=navy)
        c.font = Font(color='FFFFFF', bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for r_idx, r in enumerate(result.bundle_rows, 2):
        vals = [
            r.bundle_name, r.bundle_barcode, r.fbo, r.fbs, r.bundle_orders,
            r.bundle_stock_fbo, r.bundle_stock_fbs, r.bundle_stock_selected,
            r.component_barcode, r.component_name, r.qty_in_bundle, r.component_sales_consumption,
            r.component_stock_inside_bundles_fbo, r.component_stock_inside_bundles_fbs, r.component_stock_inside_bundles_selected,
        ]
        for c_idx, value in enumerate(vals, 1):
            cell = bs.cell(r_idx, c_idx, value)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical='center', wrap_text=c_idx in (1, 10))
    for i, width in enumerate([24, 22, 12, 12, 14, 18, 18, 20, 22, 42, 18, 20, 24, 24, 22], 1):
        bs.column_dimensions[get_column_letter(i)].width = width
    bs.freeze_panes = 'A2'
    if result.bundle_rows:
        bs.auto_filter.ref = f'A1:O{1 + len(result.bundle_rows)}'
    bs.row_dimensions[1].height = 42

    es = wb.create_sheet('Ошибки')
    for i, h in enumerate(['Источник', 'Баркод', 'Описание'], 1):
        c = es.cell(1, i, h)
        c.fill = PatternFill('solid', fgColor=navy)
        c.font = Font(color='FFFFFF', bold=True)
    if result.errors:
        for r_idx, (source, bc, desc) in enumerate(result.errors, 2):
            es.cell(r_idx, 1, source)
            es.cell(r_idx, 2, bc)
            es.cell(r_idx, 3, desc)
            for c_idx in range(1, 4):
                es.cell(r_idx, c_idx).fill = PatternFill('solid', fgColor=red if source == 'Комплекты' else yellow)
                es.cell(r_idx, c_idx).border = Border(left=thin, right=thin, top=thin, bottom=thin)
                es.cell(r_idx, c_idx).alignment = Alignment(vertical='top', wrap_text=True)
    else:
        es['A2'] = 'Ошибок не обнаружено'
        es['A2'].fill = PatternFill('solid', fgColor=green)
    es.column_dimensions['A'].width = 18
    es.column_dimensions['B'].width = 24
    es.column_dimensions['C'].width = 90
    es.freeze_panes = 'A2'

    wb.save(path)
