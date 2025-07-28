
import base64
import io
import logging
import os
import pandas as pd
import tempfile
from aiogram import Bot, Dispatcher, types, F
from aiogram.enums import ParseMode
from aiogram.types import FSInputFile
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
from aiohttp import web

API_TOKEN = os.getenv("TOKEN")
WEBHOOK_PATH = "/webhook"
WEBHOOK_URL = os.getenv("WEBHOOK_URL") + WEBHOOK_PATH

bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# === Новый шаблон XLSX в base64 ===
EXCEL_TEMPLATE_BASE64 = """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""".strip()

def write_template_if_missing():
    if not os.path.exists("template.xlsx"):
        with open("template.xlsx", "wb") as f:
            f.write(base64.b64decode(EXCEL_TEMPLATE_BASE64))
        print("✅ Файл template.xlsx создан.")

def clean_string(value):
    if pd.isna(value):
        return ""
    return str(value).strip().replace(' ', '').replace('​', '').replace('  ', ' ')

def parse_driver_data(value):
    clean_val = clean_string(value)
    if not clean_val:
        return "-", "-", "-"
    parts = [part.strip() for part in clean_val.split(",")]
    parts += ["-"] * (3 - len(parts))
    return parts[0], parts[1], parts[2]

def extract_driver_contact(value):
    clean_val = clean_string(value)
    if not clean_val:
        return "", ""
    parts = clean_val.split(",")
    driver = parts[0].strip() if len(parts) > 0 else ""
    contact = parts[1].strip(" )") if len(parts) > 1 else ""
    return driver, contact

COLUMN_MAPPING = {
    "Время подачи ТС на загрузку": "Время подачи ТС на загрузку",
    "Окно выгрузки": "Время подачи ТС на разгрузку",
    "Кост центр": "Внутренний заказчик",
    "Отправитель": "Организация отправитель",
    "Получатель": "Организация получатель",
    "Адрес загрузки": "Адрес загрузки",
    "Адрес разгрузки": "Адрес разгрузки",
    "Тип грузовых мест": "На паллетах",
    "Объем план, м3 (ТС)": "Объем, куб м",
    "Кол-во мест (ТС)": "Кол-во грузовых мест",
    "Требуется ли ПРР?": "Требуется ли ПРР?",
    "Часы работы": "Часы работы",
    "Кол-во единиц товара": "Количество единиц товара",
    "Комментарии": "Комментарии",
}

def process_xlsx(input_bytes: bytes) -> io.BytesIO:
    write_template_if_missing()
    template_df = pd.read_excel("template.xlsx")
    template_df.columns = [clean_string(c) for c in template_df.columns]
    input_df = pd.read_excel(io.BytesIO(input_bytes)).applymap(clean_string)

    if "Данные по водителю и машине" in input_df.columns:
        marka, number, trailer = zip(*input_df["Данные по водителю и машине"].map(parse_driver_data))
        input_df["Марка ТС"] = marka
        input_df["Номер ТС"] = number
        input_df["Номер прицепа"] = trailer

    if "Контакты" in input_df.columns:
        driver, contact = zip(*input_df["Контакты"].map(extract_driver_contact))
        input_df["Водитель"] = driver
        input_df["Контакты"] = contact

    result = pd.DataFrame()
    for col in template_df.columns:
        if col in ["Марка ТС", "Номер ТС", "Номер прицепа", "Контакты"]:
            result[col] = input_df.get(col, "")
        elif col in COLUMN_MAPPING:
            result[col] = input_df.get(COLUMN_MAPPING[col], "")
        else:
            result[col] = ""

    output = io.BytesIO()
    result.to_excel(output, index=False)
    output.seek(0)
    return output

@dp.message(F.document)
async def handle_doc(message: types.Message, bot: Bot):
    document = message.document
    if not document.file_name.endswith(".xlsx"):
        await message.reply("❌ Пришлите файл в формате .xlsx")
        return
    file = await bot.download(document)
    processed = process_xlsx(file.read())
    await message.reply_document(FSInputFile(processed, filename="output.xlsx"))

async def on_startup(bot: Bot):
    await bot.set_webhook(WEBHOOK_URL)

def main():
    app = web.Application()
    SimpleRequestHandler(dispatcher=dp, bot=bot).register(app, path=WEBHOOK_PATH)
    setup_application(app, dp, bot=bot)
    web.run_app(app, port=8080)

if __name__ == "__main__":
    main()