from aiogram import Bot, Dispatcher, types
from aiogram.webhook.aiohttp_server import SimpleRequestHandler
from aiohttp import web
import pandas as pd
import os
from pathlib import Path
import base64

TOKEN = os.getenv("TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL")

TEMPLATE_B64 = """<Твой длинный base64 шаблон>"""  # Сюда подставишь свой TEMPLATE_B64

bot = Bot(token=TOKEN)
dp = Dispatcher()

def clean_string(value):
    if pd.isna(value):
        return ""
    return str(value).strip().replace('\xa0', '').replace('\u200b', '').replace('  ', ' ')

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

def write_template_if_missing():
    if not os.path.exists("стало2.xlsx"):
        with open("стало2.xlsx", "wb") as f:
            f.write(base64.b64decode(TEMPLATE_B64))

async def process_file(file_path):
    write_template_if_missing()
    template_df = pd.read_excel("стало2.xlsx")
    template_df.columns = [clean_string(c) for c in template_df.columns]

    df = pd.read_excel(file_path)
    df = df.applymap(clean_string)

    if "Данные по водителю и машине" in df.columns:
        marka, number, trailer = zip(*df["Данные по водителю и машине"].map(parse_driver_data))
        df["Марка ТС"] = marka
        df["Номер ТС"] = number
        df["Номер прицепа"] = trailer

    if "Контакты" in df.columns:
        driver, contact = zip(*df["Контакты"].map(extract_driver_contact))
        df["Водитель"] = driver
        df["Контакты"] = contact

    for col in template_df.columns:
        if col not in df.columns:
            df[col] = ""

    df = df[template_df.columns]

    out_path = file_path.replace(".xlsx", "_shuffled.xlsx")
    df.to_excel(out_path, index=False)
    return out_path

@dp.message(lambda msg: msg.document and msg.document.file_name.endswith('.xlsx'))
async def handle_file(message: types.Message):
    file_info = await bot.get_file(message.document.file_id)
    file_path = f"/tmp/{message.document.file_name}"
    await bot.download_file(file_info.file_path, file_path)

    out_path = await process_file(file_path)
    await message.answer_document(types.FSInputFile(out_path))

    os.remove(file_path)
    os.remove(out_path)

async def on_startup(app):
    await bot.set_webhook(WEBHOOK_URL + "/webhook")

async def on_shutdown(app):
    await bot.delete_webhook()
    await bot.session.close()

app = web.Application()
SimpleRequestHandler(dispatcher=dp, bot=bot).register(app, path="/webhook")
app.on_startup.append(on_startup)
app.on_shutdown.append(on_shutdown)

if __name__ == "__main__":
    web.run_app(app, port=int(os.getenv("PORT", 8080)))