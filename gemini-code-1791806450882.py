import os
from datetime import datetime
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)
import openpyxl

BOT_TOKEN = "BOT_TOKENINGIZNI_SHU_YERGA_YOZING"

EXCEL_FILE = "hisobot.xlsx"
MAHSULOTLAR = [
    "Jitkiy gvoz",
    "Akrilik germetik",
    "Tomchi",
    "Qopqoqlik pena",
    "Professional Pena",
]


# Excel faylini tayyorlash yoki ochish
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Anketalar"
        headers = [
            "Sana",
            "Klient nomi",
            "Mahsulot turi",
            "Brend nomi",
            "Sfati xaqida odzf",
            "Kg",
            "Narxi",
            "Yillik bonus",
            "1 oyda o'rtacha abyom",
            "Qancha ast bilan ishledi",
            "Tavarga garantiya",
            "Firma qulayliklari",
            "Bizdan nima uchun olmaysiz",
            "Shu muammolarni bartaraf qilsak 2027 yildan biz bilan ishlaysizmi",
        ]
        ws.append(headers)
        wb.save(EXCEL_FILE)


init_excel()


# FSM holatlari (savollar ketma-ketligi)
class Form(StatesGroup):
    klient_nomi = State()
    mahsulot_turi = State()
    brend_nomi = State()
    odzf = State()
    kg = State()
    narxi = State()
    yillik_bonus = State()
    oylik_abyom = State()
    ast = State()
    garantiya = State()
    qulayliklar = State()
    nima_uchun_olmaydi = State()
    hamkorlik_2027 = State()


bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "Assalomu alaykum! Yangi anketani to'ldirishni boshlaymiz.\n\n"
        "1. Klient nomini kiriting:",
        reply_markup=ReplyKeyboardRemove(),
    )
    await state.set_state(Form.klient_nomi)


@dp.message(Form.klient_nomi)
async def process_klient(message: Message, state: FSMContext):
    await state.update_data(klient_nomi=message.text)

    # Mahsulotlarni tugma qilib chiqarish
    keyboard = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=m)] for m in MAHSULOTLAR],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await message.answer("Mahsulot turini tanlang:", reply_markup=keyboard)
    await state.set_state(Form.mahsulot_turi)


@dp.message(Form.mahsulot_turi)
async def process_mahsulot(message: Message, state: FSMContext):
    await state.update_data(mahsulot_turi=message.text)
    await message.answer(
        "Brend nomini kiriting:", reply_markup=ReplyKeyboardRemove()
    )
    await state.set_state(Form.brend_nomi)


@dp.message(Form.brend_nomi)
async def process_brend(message: Message, state: FSMContext):
    await state.update_data(brend_nomi=message.text)
    await message.answer("Sifati haqida fikri (otziv):")
    await state.set_state(Form.odzf)


@dp.message(Form.odzf)
async def process_odzf(message: Message, state: FSMContext):
    await state.update_data(odzf=message.text)
    await message.answer("Kg (og'irligi):")
    await state.set_state(Form.kg)


@dp.message(Form.kg)
async def process_kg(message: Message, state: FSMContext):
    await state.update_data(kg=message.text)
    await message.answer("Narxi:")
    await state.set_state(Form.narxi)


@dp.message(Form.narxi)
async def process_narxi(message: Message, state: FSMContext):
    await state.update_data(narxi=message.text)
    await message.answer("Yillik bonus:")
    await state.set_state(Form.yillik_bonus)


@dp.message(Form.yillik_bonus)
async def process_bonus(message: Message, state: FSMContext):
    await state.update_data(yillik_bonus=message.text)
    await message.answer("1 oyda o'rtacha abyom:")
    await state.set_state(Form.oylik_abyom)


@dp.message(Form.oylik_abyom)
async def process_abyom(message: Message, state: FSMContext):
    await state.update_data(oylik_abyom=message.text)
    await message.answer("Qancha ast (otstoy/qoldiq) bilan ishlaydi:")
    await state.set_state(Form.ast)


@dp.message(Form.ast)
async def process_ast(message: Message, state: FSMContext):
    await state.update_data(ast=message.text)
    await message.answer("Tovarga garantiya bormi / qanday?")
    await state.set_state(Form.garantiya)


@dp.message(Form.garantiya)
async def process_garantiya(message: Message, state: FSMContext):
    await state.update_data(garantiya=message.text)
    await message.answer("Firma qulayliklari:")
    await state.set_state(Form.qulayliklar)


@dp.message(Form.qulayliklar)
async def process_qulayliklar(message: Message, state: FSMContext):
    await state.update_data(qulayliklar=message.text)
    await message.answer("Bizdan nima uchun olmaysiz?")
    await state.set_state(Form.nima_uchun_olmaydi)


@dp.message(Form.nima_uchun_olmaydi)
async def process_sabab(message: Message, state: FSMContext):
    await state.update_data(nima_uchun_olmaydi=message.text)
    await message.answer(
        "Shu muammolarni bartaraf qilsak 2027 yildan biz bilan ishlaysizmi?"
    )
    await state.set_state(Form.hamkorlik_2027)


@dp.message(Form.hamkorlik_2027)
async def process_yakun(message: Message, state: FSMContext):
    await state.update_data(hamkorlik_2027=message.text)
    data = await state.get_data()

    # Ma'lumotlarni Excel faylga qo'shish
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    sana_bugun = datetime.now().strftime("%Y-%m-%d %H:%M")

    qator = [
        sana_bugun,
        data.get("klient_nomi"),
        data.get("mahsulot_turi"),
        data.get("brend_nomi"),
        data.get("odzf"),
        data.get("kg"),
        data.get("narxi"),
        data.get("yillik_bonus"),
        data.get("oylik_abyom"),
        data.get("ast"),
        data.get("garantiya"),
        data.get("qulayliklar"),
        data.get("nima_uchun_olmaydi"),
        data.get("hamkorlik_2027"),
    ]
    ws.append(qator)
    wb.save(EXCEL_FILE)

    await state.clear()
    await message.answer(
        "✅ Rahmat! Ma'lumotlar saqlandi.\n\n"
        "Keyingi anketani kiritish uchun /start ni bosing."
    )


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    import asyncio

    asyncio.run(main())