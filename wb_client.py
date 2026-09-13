import asyncio
from datetime import date, datetime, time, timedelta, timezone
from typing import Any

import aiohttp

from config import WB_CONTENT_URL, WB_STATISTICS_URL, WB_ANALYTICS_URL

MOSCOW = timezone(timedelta(hours=3))


class WBApiError(RuntimeError):
    pass


class LegacyOrdersUnavailable(WBApiError):
    pass


class WBClient:
    def __init__(self, token: str):
        self.token = token
        self.headers = {'Authorization': token, 'Content-Type': 'application/json'}

    async def _request(self, method: str, url: str, *, json=None, params=None, retries: int = 5):
        timeout = aiohttp.ClientTimeout(total=180)
        async with aiohttp.ClientSession(headers=self.headers, timeout=timeout) as session:
            for attempt in range(retries):
                async with session.request(method, url, json=json, params=params) as r:
                    if r.status == 204:
                        return None
                    if r.status == 429:
                        retry_after = r.headers.get('Retry-After')
                        delay = int(retry_after) if retry_after and retry_after.isdigit() else min(60, 5 * (attempt + 1))
                        await asyncio.sleep(delay)
                        continue
                    text = await r.text()
                    if r.status in (404, 410) and '/api/v1/supplier/orders' in url:
                        raise LegacyOrdersUnavailable(f'Старый Orders API недоступен ({r.status}).')
                    if r.status >= 400:
                        raise WBApiError(f'WB API {r.status}: {text[:1500]}')
                    if not text:
                        return None
                    try:
                        return await r.json(content_type=None)
                    except Exception as e:
                        raise WBApiError(f'WB API вернул не-JSON: {text[:800]}') from e
        raise WBApiError('WB API: превышено число повторных попыток')

    async def get_all_cards(self) -> list[dict]:
        url = f'{WB_CONTENT_URL}/content/v2/get/cards/list'
        result: list[dict] = []
        cursor: dict[str, Any] = {'limit': 100}
        while True:
            payload = {'settings': {'cursor': cursor, 'filter': {'withPhoto': -1}}}
            data = await self._request('POST', url, json=payload) or {}
            cards = data.get('cards') or []
            result.extend(cards)
            rc = data.get('cursor') or {}
            if len(cards) < 100:
                break
            nm_id, updated_at = rc.get('nmID'), rc.get('updatedAt')
            if not nm_id or not updated_at:
                break
            cursor = {'limit': 100, 'nmID': nm_id, 'updatedAt': updated_at}
        return result

    async def get_orders_legacy(self, date_from: date, date_to: date) -> list[dict]:
        """Detailed orders. We fetch changes since date_from, then filter by actual order `date`.
        This endpoint exposes barcode + warehouseType, which is ideal for FBO/FBS split.
        """
        url = f'{WB_STATISTICS_URL}/api/v1/supplier/orders'
        params = {'dateFrom': f'{date_from.isoformat()}T00:00:00', 'flag': 0}
        rows = await self._request('GET', url, params=params) or []
        if not isinstance(rows, list):
            raise WBApiError('Неожиданный формат ответа Orders API')
        out = []
        for row in rows:
            dt = parse_wb_datetime(row.get('date'))
            if dt and date_from <= dt.date() <= date_to:
                out.append(row)
        return out

    async def get_orders_feed(self, date_from: date, date_to: date) -> list[dict]:
        """Fallback for <=31 calendar days. Feed period is based on current status time,
        so we still filter returned rows by createdAt.
        """
        if (date_to - date_from).days + 1 > 31:
            raise WBApiError('Order Feed поддерживает максимум 31 день; для 60/90 дней нужен Orders API.')

        url = f'{WB_ANALYTICS_URL}/api/analytics/v1/order-feed'
        start = datetime.combine(date_from, time.min, tzinfo=MOSCOW)
        end = datetime.combine(date_to, time.max.replace(microsecond=0), tzinfo=MOSCOW)
        snapshot = None
        offset = 0
        limit = 1000
        result: list[dict] = []

        while True:
            pagination = {'offset': offset, 'limit': limit}
            if snapshot:
                pagination['snapshotTime'] = snapshot
            payload = {
                'selectedPeriod': {'start': start.isoformat(), 'end': end.isoformat()},
                'nmIds': [], 'subjectIds': [], 'brandNames': [], 'tagIds': [],
                'pagination': pagination,
            }
            data = await self._request('POST', url, json=payload) or {}
            block = data.get('data') or {}
            snapshot = block.get('snapshotTime') or snapshot
            rows = block.get('orders') or []
            if not rows:
                break
            result.extend(rows)
            if len(rows) < limit:
                break
            offset += len(rows)

        out = []
        for row in result:
            dt = parse_wb_datetime(row.get('createdAt'))
            if dt and date_from <= dt.date() <= date_to:
                out.append(row)
        return out

    async def get_orders(self, date_from: date, date_to: date) -> tuple[list[dict], str]:
        try:
            rows = await self.get_orders_legacy(date_from, date_to)
            return rows, 'orders-api'
        except LegacyOrdersUnavailable:
            rows = await self.get_orders_feed(date_from, date_to)
            return rows, 'order-feed'


def parse_wb_datetime(value: Any) -> datetime | None:
    if not value:
        return None
    s = str(value).strip().replace('Z', '+00:00')
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S'):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                pass
    return None
