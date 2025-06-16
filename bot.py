
from aiogram import Bot, Dispatcher, types
import asyncio
import pandas as pd
import os

TOKEN = os.getenv("TOKEN", "7934411658:AAG7QE5fDW-bpNxCTvODqjSrsbBcRY86A1A")

async def process_file(file_path):
    df = pd.read_excel(file_path)
    # Здесь можно вставить твою логику обработки
    out_path = file_path.replace(".xlsx", "_shuffled.xlsx")
    df.to_excel(out_path, index=False)
    return out_path

async def handle_file(message: types.Message):
    file = await message.document.download()
    file_path = file.name
    out_path = await process_file(file_path)
    await message.answer_document(types.FSInputFile(out_path))
    os.remove(file_path)
    os.remove(out_path)

async def main():
    bot = Bot(token=TOKEN)
    dp = Dispatcher()
    dp.message.register(handle_file, lambda msg: msg.document and msg.document.file_name.endswith('.xlsx'))
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
