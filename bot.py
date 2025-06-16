from aiogram import Bot, Dispatcher, types
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
from aiohttp import web
import pandas as pd
import os

TOKEN = os.getenv("TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL")

bot = Bot(token=TOKEN)
dp = Dispatcher()

async def process_file(file_path):
    df = pd.read_excel(file_path)
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
    await bot.session.close()  # ВАЖНО: закрыть клиентскую сессию

app = web.Application()
SimpleRequestHandler(dispatcher=dp, bot=bot).register(app, path="/webhook")
app.on_startup.append(on_startup)
app.on_shutdown.append(on_shutdown)

if __name__ == "__main__":
    web.run_app(app, port=int(os.getenv("PORT", 8080)))