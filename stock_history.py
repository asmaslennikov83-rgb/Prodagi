import csv
import io
import re
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation


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
    totals = defaultdict(lambda: defaultdict(int))
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
            reader = csv.DictReader(io.StringIO(text), dialect=dialect)
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
                chrt = str(row.get(chrt_column) or '').strip()
                nm = str(row.get(nm_column) or '').strip()
                size = str(row.get(size_column) or '0').strip()
                key = ('chrt', chrt) if chrt and chrt != '0' else ('size', nm, size)
                if key[0] == 'size' and not (nm and size_column):
                    raise ValueError(f'Не указан идентификатор размера в строке {line_number}.')
                entries = [(parse_day(row.get(day_column)), row.get(quantity_column))] if day_column and quantity_column else [(day, row.get(header)) for header, day in date_columns.items()]
                for day, value in entries:
                    if day is None:
                        raise ValueError(f'Некорректная дата в строке {line_number}.')
                    if not date_from <= day <= date_to:
                        continue
                    try:
                        quantity = Decimal(str(value).replace('\xa0', '').replace(' ', '').replace(',', '.'))
                        if not quantity.is_finite() or quantity != quantity.to_integral_value() or quantity < 0:
                            raise InvalidOperation
                    except InvalidOperation as exc:
                        raise ValueError(f'Некорректный остаток в строке {line_number}.') from exc
                    totals[key][day.isoformat()] += int(quantity)
    return totals


def apply_stock_history(products, payload, date_from, date_to):
    history = read_stock_history(payload, date_from, date_to)
    expected = {(date_from + timedelta(days=offset)).isoformat() for offset in range((date_to - date_from).days + 1)}
    for product in products:
        quantities = history.get(('chrt', str(product.chrt_id)))
        if quantities is None:
            quantities = history.get(('size', str(product.nm_id), product.size))
        if quantities is None or set(quantities) != expected:
            product.history_warning = 'История наличия неполная: скорость рассчитана по календарным дням.'
            continue
        product.availability_dates = {day for day, quantity in quantities.items() if quantity > 0} | (product.order_dates & expected)
