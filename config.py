import os
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN', '').strip()
WB_API_KEY_1 = os.getenv('WB_API_KEY_1', '').strip()
WB_API_KEY_2 = os.getenv('WB_API_KEY_2', '').strip()
CABINET_1_NAME = os.getenv('CABINET_1_NAME', 'Кабинет 1').strip() or 'Кабинет 1'
CABINET_2_NAME = os.getenv('CABINET_2_NAME', 'Кабинет 2').strip() or 'Кабинет 2'

WB_CONTENT_URL = 'https://content-api.wildberries.ru'
WB_STATISTICS_URL = 'https://statistics-api.wildberries.ru'
WB_ANALYTICS_URL = 'https://seller-analytics-api.wildberries.ru'
