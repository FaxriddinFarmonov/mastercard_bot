# mastercard_bot_telegram.py
import re
import pdfplumber
from openpyxl import Workbook
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import FSInputFile
import asyncio
import os

TOKEN = "7970127157:AAE98YOZf3P2BQ2ScSiqu0TdbYp9MP-Mch0"
bot = Bot(token=TOKEN)
dp = Dispatcher()


# ======================= Helper funksiyalar =======================

def _to_float(s):
    if not s:
        return 0.0
    s = s.strip().replace(",", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except:
        return 0.0


def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text


def _extract_section_totals(block, header_pattern):
    m = re.search(header_pattern, block, re.I | re.S)
    if not m:
        return 0.0, 0.0, 0.0

    sub = block[m.end():]

    total_m = re.search(r"\bTotal:\s*([-\d,().]+)", sub)
    debit_m = re.search(r"Total\s+Debit\s+Fees[:\s]*([-\d,().]+)", sub)
    credit_m = re.search(r"Total\s+Credit\s+Fees[:\s]*([-\d,().]+)", sub)

    total = _to_float(total_m.group(1)) if total_m else 0.0
    debit = _to_float(debit_m.group(1)) if debit_m else 0.0
    credit = _to_float(credit_m.group(1)) if credit_m else 0.0

    return total, debit, credit


def parse_mastercard_report(pdf_text):
    results = {}

    blocks = re.split(r"Business Date:", pdf_text)
    for block in blocks[1:]:
        date_m = re.search(r"(\d{4}\.\d{2}\.\d{2})", block)
        cur_m = re.search(r"Reconcilation Currency:\s*([A-Za-z]+)", block, re.I)

        if not date_m or not cur_m:
            continue

        date = date_m.group(1)
        currency = cur_m.group(1).upper()

        key = (date, currency)

        if key not in results:
            results[key] = {
                "currency": currency,
                "acquiring": {"total": 0.0, "debit": 0.0, "credit": 0.0},
                "issuing": {"total": 0.0, "debit": 0.0, "credit": 0.0},
                "misc": {"total": 0.0, "debit": 0.0, "credit": 0.0},
                "grand_total": 0.0,
                "grand_debit": 0.0,
                "grand_credit": 0.0
            }

        acq_tot, acq_deb, acq_cred = _extract_section_totals(
            block, r"Transaction Function:\s*First Presentment Original/Acquiring"
        )
        iss_tot, iss_deb, iss_cred = _extract_section_totals(
            block, r"Transaction Function:\s*First Presentment Original/Issuing"
        )
        misc_tot, misc_deb, misc_cred = _extract_section_totals(
            block, r"Transaction Function:\s*Fee Collection Original/Miscellaneous"
        )

        grand_m = re.search(r"Grand Total:\s*([-\d,().]+)", block)
        grand_deb_m = re.search(r"Grand Total Debit Fees\s*([-\d,().]+)", block)
        grand_cred_m = re.search(r"Grand Total Credit Fees\s*([-\d,().]+)", block)

        results[key]["acquiring"]["total"] += acq_tot
        results[key]["acquiring"]["debit"] += acq_deb
        results[key]["acquiring"]["credit"] += acq_cred

        results[key]["issuing"]["total"] += iss_tot
        results[key]["issuing"]["debit"] += iss_deb
        results[key]["issuing"]["credit"] += iss_cred

        results[key]["misc"]["total"] += misc_tot
        results[key]["misc"]["debit"] += misc_deb
        results[key]["misc"]["credit"] += misc_cred

        if grand_m:
            results[key]["grand_total"] = _to_float(grand_m.group(1))
        if grand_deb_m:
            results[key]["grand_debit"] = _to_float(grand_deb_m.group(1))
        if grand_cred_m:
            results[key]["grand_credit"] = _to_float(grand_cred_m.group(1))

    return results


def create_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "MasterCard Report"

    ws.append(["Date", "Acquiring", "Issuing", "Miscellaneous", "Grand Total", "Type", "Currency"])
    ws.append([])

    for (date, currency) in sorted(data.keys()):
        row = data[(date, currency)]

        ws.append([date,
                   row["acquiring"]["total"],
                   row["issuing"]["total"],
                   row["misc"]["total"],
                   row["grand_total"],
                   "Total",
                   currency])

        ws.append(["",
                   row["acquiring"]["debit"],
                   row["issuing"]["debit"],
                   row["misc"]["debit"],
                   row["grand_debit"],
                   "Total Debit Fees",
                   ""])

        ws.append(["",
                   row["acquiring"]["credit"],
                   row["issuing"]["credit"],
                   row["misc"]["credit"],
                   row["grand_credit"],
                   "Total Credit Fees",
                   ""])
        ws.append([])

    for r in ws.iter_rows(min_row=3):
        for c in r:
            if isinstance(c.value, (int, float)):
                c.number_format = "0.00"

    wb.save(output_file)


# ======================= Telegram handlerlar =======================

@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    await message.answer("👋 Salom! Menga faqat *MasterCard PDF* yuboring, men uni Excelga aylantirib qaytaraman.")


# 🚫 Rasmlar
@dp.message(F.photo)
async def handle_photo(message: types.Message):
    await message.answer("❗ Rasm qabul qilinmaydi. Faqat MasterCard PDF yuboring.")

# 🚫 Videolar
@dp.message(F.video)
async def handle_video(message: types.Message):
    await message.answer("❗ Video qabul qilinmaydi. Faqat MasterCard PDF yuboring.")

# 🚫 Ovozli xabarlar
@dp.message(F.voice)
async def handle_voice(message: types.Message):
    await message.answer("❗ Ovozli xabar qabul qilinmaydi. Faqat MasterCard PDF yuboring.")


# 📥 Faqat PDF fayllarni qabul qilish
@dp.message(F.document)
async def handle_doc(message: types.Message):
    if message.document.mime_type != "application/pdf":
        await message.answer("❗ Bu fayl PDF emas. Faqat MasterCard PDF yuboring.")
        return

    os.makedirs("downloads", exist_ok=True)
    filename = message.document.file_name or f"{message.document.file_unique_id}.pdf"
    pdf_path = os.path.join("downloads", f"{message.document.file_unique_id}_{filename}")

    try:
        # 📥 Faylni olish
        file = await bot.get_file(message.document.file_id)
        await bot.download_file(file.file_path, pdf_path)

        await message.answer("✅ PDF qabul qilindi. O‘qilmoqda...")

        # 📖 PDF dan text olish
        text = extract_text_from_pdf(pdf_path)

        # 🔑 MasterCard hisobotligini tekshirish
        if "MasterCard Financial Position Details by Transaction Types/ISS And ACQ" not in text:
            await message.answer("❗ Bu PDF MasterCard hisobotiga o‘xshamaydi.")
            return

        # 📊 Parse qilish
        parsed = parse_mastercard_report(text)

        if not parsed:
            await message.answer("❗ Kerakli ma’lumotlar topilmadi. To‘g‘ri MasterCard hisobotini yuboring.")
            return

        # 📊 Excel yaratish
        out_xlsx = os.path.join("downloads", f"{message.document.file_unique_id}.xlsx")
        create_excel(parsed, out_xlsx)

        # 📤 Excelni yuborish
        await message.reply_document(FSInputFile(out_xlsx))

    except Exception as e:
        await message.answer(f"❗ Xatolik: {e}")

    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        if 'out_xlsx' in locals() and os.path.exists(out_xlsx):
            os.remove(out_xlsx)


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

