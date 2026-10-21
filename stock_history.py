import csv
import io
import re
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation


class StockHistory(defaultdict):
    def __init__(self):
        super().__init__(lambda: defaultdict(int))
        self.issues = defaultdict(list)
        self.missing = defaultdict(list)


def normalize(value):
    return re.sub(r'[^a-zа-я0-9]', '', str(value).strip().casefold().replace('ё', 'е'))


def parse_day(value):
    text = str(value).strip()
    for pattern in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(text[:10], pattern).date()
        except ValueError:
            pass
    return None


def read_stock_history(payload, date_from, date_to):
    totals = StockHistory()
    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        files = [entry for entry in archive.infolist() if entry.filename.casefold().endswith('.csv')]
        if not files or sum(entry.file_size for entry in files) > 200 * 1024 * 1024:
            raise ValueError('CSV истории отсутствует или превышает 200 МБ после распаковки.')
        for entry in files:
            raw = archive.read(entry)
            try:
                text = raw.decode('utf-8-sig')
            except UnicodeDecodeError:
                text = raw.decode('cp1251')
            try:
                dialect = csv.Sniffer().sniff(text[:16384], delimiters=',;\t')
            except csv.Error as exc:
                raise ValueError('Не удалось определить разделитель CSV истории.') from exc
            reader = csv.DictReader(io.StringIO(text), delimiter=dialect.delimiter, quotechar='"', doublequote=True, strict=True)
            headers = {normalize(header): header for header in reader.fieldnames or []}
            def column(*aliases):
                return next((headers[normalize(alias)] for alias in aliases if normalize(alias) in headers), None)
            chrt_column = column('ChrtID', 'ID размера', 'Идентификатор размера')
            nm_column = column('NmID', 'Артикул WB', 'Артикул ВБ')
            size_column = column('SizeName', 'TechSize', 'Size', 'Размер')
            day_column = column('Date', 'Dt', 'Дата', 'Day')
            quantity_column = column('Quantity', 'Stock', 'StockCount', 'StockQuantity', 'Остаток', 'Остатки', 'Остаток шт')
            date_columns = {header: parse_day(header) for header in reader.fieldnames or [] if parse_day(header)}
            if not (chrt_column or (nm_column and size_column)) or not ((day_column and quantity_column) or date_columns):
                raise ValueError('Неизвестные колонки CSV истории: ' + ', '.join(reader.fieldnames or []))
            for line_number, row in enumerate(reader, 2):
                if None in row or any(value is None for value in row.values()):
                    raise ValueError(f'{entry.filename}: число полей в записи {line_number} не соответствует заголовку CSV.')
                chrt = str(row.get(chrt_column) or '').strip()
                nm = str(row.get(nm_column) or '').strip()
                size = str(row.get(size_column) or '0').strip()
                key = ('chrt', chrt) if chrt and chrt != '0' else ('size', nm, size)
                if key[0] == 'size' and not (nm and size_column):
                    raise ValueError(f'Не указан идентификатор размера в строке {line_number}.')
                entries = [(parse_day(row.get(day_column)), row.get(quantity_column))] if day_column and quantity_column else [(day, row.get(header)) for header, day in date_columns.items()]
                for day, value in entries:
                    if day is None:
                        totals.issues[key].append(
                            f'{entry.filename}: строка {line_number}, дата {str(row.get(day_column))[:60]!r} не распознана.'
                        )
                        continue
                    if not date_from <= day <= date_to:
                        continue
                    if str(value).strip() in {'', '-', '—', '–'}:
                        totals.missing[key].append(day.isoformat())
                        continue
                    try:
                        quantity = Decimal(re.sub(r'\s+', '', str(value)).replace(',', '.'))
                        if not quantity.is_finite() or quantity != quantity.to_integral_value() or quantity < 0:
                            raise InvalidOperation
                    except InvalidOperation:
                        field = quantity_column if day_column and quantity_column else next(header for header, field_day in date_columns.items() if field_day == day)
                        totals.issues[key].append(
                            f'{entry.filename}: строка {line_number}, {day.isoformat()}, '
                            f'поле {field!r}, остаток {str(value)[:60]!r} не является неотрицательным целым числом.'
                        )
                        continue
                    totals[key][day.isoformat()] += int(quantity)
    return totals


def apply_stock_history(products, payload, date_from, date_to, stock_mode='both'):
    history = read_stock_history(payload, date_from, date_to)
    expected = {(date_from + timedelta(days=offset)).isoformat() for offset in range((date_to - date_from).days + 1)}
    for product in products:
        product.availability_dates = None
        product.history_warning = ''
        product.history_status = ''
        product.availability_mode = 'fbs' if stock_mode == 'fbs' else 'both'
        key = ('chrt', str(product.chrt_id))
        if key not in history and key not in history.issues and key not in history.missing:
            key = ('size', str(product.nm_id), product.size)
        quantities = history.get(key)
        issues = history.issues.get(key)
        if issues:
            details = '; '.join(issues[:3])
            if len(issues) > 3:
                details += f'; ещё ошибок: {len(issues) - 3}'
            product.history_warning = f'История этого товара содержит ошибки, расчёт по календарным дням: {details}'
            product.history_status = 'invalid_data'
            continue
        if quantities is None or set(quantities) != expected or history.missing.get(key):
            known_days = len((set(quantities or {}) & expected) - set(history.missing.get(key, [])))
            if quantities is None and key not in history.missing:
                reason = f'WB не вернул историю этого товара/размера (ID размера {product.chrt_id})'
                product.history_status = 'not_returned'
            else:
                missing_dates = sorted(history.missing.get(key, []))
                reason = f'История наличия неполная: числовые остатки за {known_days} из {len(expected)} дней'
                if missing_dates:
                    reason += f'; WB не указал остаток за {len(set(missing_dates))} дней (прочерк или пустое значение)'
                product.history_status = 'missing_values' if missing_dates else 'incomplete_period'
            product.history_warning = reason + '. Скорость рассчитана по календарным дням.'
            continue
        order_dates = product.fbs_order_dates if stock_mode == 'fbs' else product.order_dates
        product.availability_dates = {day for day, quantity in quantities.items() if quantity > 0} | (order_dates & expected)
        product.history_status = 'complete'
