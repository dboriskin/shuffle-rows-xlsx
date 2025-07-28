import logging
import os
from pathlib import Path
from io import BytesIO
import base64

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from aiogram.enums import ParseMode
from aiogram.types import BufferedInputFile
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
from aiohttp import web
import pandas as pd

# === Config ===
API_TOKEN = os.getenv("BOT_TOKEN")
BASE_WEBHOOK_URL = os.getenv("BASE_WEBHOOK_URL")
WEBHOOK_PATH = "/webhook"
WEBHOOK_URL = BASE_WEBHOOK_URL + WEBHOOK_PATH

# === Новый шаблон XLSX в base64 ===
EXCEL_TEMPLATE_BASE64 = """
... (вставь сюда твой base64)
""".strip()

# === Logging ===
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === Init bot ===
session = AiohttpSession()
bot = Bot(token=API_TOKEN, session=session)
dp = Dispatcher()

# === Колонки-мэппинг ===
COLUMN_MAPPING = {
    "Время подачи ТС на загрузку": "Время подачи ТС на загрузку",
    "Окно выгрузки": "Время подачи ТС на разгрузку",
    "Кост центр": "Внутренний заказчик",
    "Отправитель": "Организация отправитель",
    "Получатель": "Организация получатель",
    "Адрес загрузки": "Адрес загрузки",
    "Адрес разгрузки": "Адрес разгрузки",
    "Объем план, м3 (ТС)": "Объем, куб м",
    "Кол-во мест (ТС)": "Кол-во грузовых мест",
    "Требуется ли ПРР?": "Требуется ли погрузка",
    "Часы работы": "Часы работы",
    "Кол-во единиц товара": "Количество единиц товара",
    "Комментарии": "Комментарии",
    "Тип маршрута": "направление",
}

# === Подготовка шаблона ===
def load_template():
    decoded = base64.b64decode(EXCEL_TEMPLATE_BASE64)
    with BytesIO(decoded) as b:
        return pd.read_excel(b)

def clean_string(value):
    if pd.isna(value):
        return ""
    return str(value).strip().replace('\xa0', '').replace('\u200b', '').replace('  ', ' ')

def parse_driver_and_phone(value):
    clean_val = clean_string(value)
    if not clean_val:
        return ""
    parts = [x.strip() for x in clean_val.split(",")]
    if len(parts) >= 2:
        return f"{parts[0]}, {parts[1]}"
    return parts[0]

@dp.message(F.document)
async def handle_doc(message: Message):
    logger.info(f"📥 Получен файл от {message.from_user.username}")
    file = await bot.download(message.document)
    incoming_df = pd.read_excel(file)
    incoming_df = incoming_df.applymap(clean_string)

    template_df = load_template()
    output_df = pd.DataFrame(columns=template_df.columns)

    for col in output_df.columns:
        if col == "Тип грузовых мест":
            output_df[col] = ""
        elif col == "Контакты":
            output_df[col] = incoming_df.get("Контакты", "").map(parse_driver_and_phone)
        elif col == "Водитель":
            output_df[col] = ""
        elif col in COLUMN_MAPPING:
            source_col = COLUMN_MAPPING[col]
            output_df[col] = incoming_df.get(source_col, "")
        else:
            output_df[col] = ""

    # Сохранение
    with BytesIO() as b:
        output_df.to_excel(b, index=False)
        b.seek(0)
        processed = BufferedInputFile(b.read(), filename="output.xlsx")
        await message.reply_document(processed)

# === Webhook setup ===
async def on_startup(bot: Bot) -> None:
    await bot.set_webhook(WEBHOOK_URL)
    logger.info(f"🚀 Webhook установлен: {WEBHOOK_URL}")

def main():
    app = web.Application()
    SimpleRequestHandler(dispatcher=dp, bot=bot).register(app, path=WEBHOOK_PATH)
    setup_application(app, dp, bot=bot)
    web.run_app(app, port=8080)

if __name__ == "__main__":
    main()