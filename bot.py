import os
import asyncio
import base64
import json
import re
import logging
from pathlib import Path
from datetime import datetime
from io import BytesIO

import anthropic
from docx import Document
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

TELEGRAM_TOKEN    = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
GOOGLE_SHEET_ID   = os.environ["GOOGLE_SHEET_ID"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]

HEADERS       = ["ФИО", "Телефон", "Возраст", "Город", "Должности", "Источник", "Кто добавил", "Дата"]
CRM_HEADERS   = ["ФИО", "НОМЕР", "ГОРОД", "ВОЗРАСТ", "ПОЗИЦИЯ", "ИСТОЧНИК"]
DATE_FMT      = "%d.%m.%Y %H:%M"

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

pending: dict[int, list] = {}
pending_timers: dict[int, asyncio.Task] = {}
BATCH_WAIT = 4

# ── Google Sheets ──────────────────────────────────────────────────────────────

def get_gc():
    creds_data = json.loads(GOOGLE_CREDS_JSON)
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_data, scopes=scopes)
    return gspread.authorize(creds)

def get_sheet(name="Резюме"):
    gc = get_gc()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(name, rows=5000, cols=20)
        if name == "Резюме":
            ws.append_row(HEADERS)
            ws.format("A1:H1", {"textFormat": {"bold": True}})
        elif name == "Экспорты":
            ws.append_row(["chat_id", "last_export"])
    return ws

def append_rows_batch(rows: list[list]):
    ws = get_sheet("Резюме")
    ws.append_rows(rows, value_input_option="USER_ENTERED")

def get_last_export(chat_id: int) -> datetime | None:
    """Возвращает время последнего экспорта для данного chat_id."""
    try:
        ws = get_sheet("Экспорты")
        records = ws.get_all_records()
        for r in records:
            if str(r.get("chat_id")) == str(chat_id):
                raw = r.get("last_export", "")
                if raw:
                    return datetime.strptime(raw, DATE_FMT)
    except Exception as e:
        log.warning(f"get_last_export error: {e}")
    return None

def set_last_export(chat_id: int, dt: datetime):
    """Сохраняет время экспорта."""
    try:
        ws = get_sheet("Экспорты")
        records = ws.get_all_values()
        # Ищем строку с этим chat_id
        for i, row in enumerate(records[1:], start=2):
            if row and str(row[0]) == str(chat_id):
                ws.update(f"B{i}", [[dt.strftime(DATE_FMT)]])
                return
        # Не нашли — добавляем новую строку
        ws.append_row([str(chat_id), dt.strftime(DATE_FMT)])
    except Exception as e:
        log.warning(f"set_last_export error: {e}")

def get_rows_since(since: datetime | None) -> list[list]:
    """Возвращает строки из листа Резюме начиная с даты since."""
    ws = get_sheet("Резюме")
    all_rows = ws.get_all_values()
    if len(all_rows) <= 1:
        return []
    result = []
    for row in all_rows[1:]:  # пропускаем заголовок
        if len(row) < 8:
            continue
        date_str = row[7].strip()
        if not date_str:
            continue
        try:
            row_dt = datetime.strptime(date_str, DATE_FMT)
        except Exception:
            continue
        if since is None or row_dt > since:
            result.append(row)
    return result

# ── xlsx генерация ─────────────────────────────────────────────────────────────

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "crm_template.xlsx")

def generate_crm_xlsx(rows: list[list]) -> BytesIO:
    """Генерирует xlsx точной копией шаблона CRM."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Очищаем старые данные (оставляем только заголовок)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # Стили — точно как в шаблоне
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    arial_bold = Font(name="Arial", bold=True)
    jetbrains  = Font(name="JetBrains Mono", bold=False)
    white_fill = PatternFill("solid", fgColor="FFFFFFFF")

    # rows колонки: ФИО(0) Телефон(1) Возраст(2) Город(3) Должности(4) Источник(5) Кто(6) Дата(7)
    for r_i, row in enumerate(rows, 2):
        fio      = row[0] if len(row) > 0 else ""
        phone    = str(row[1]).replace(".0", "") if len(row) > 1 else ""
        city     = row[3] if len(row) > 3 else ""
        age      = str(row[2]).replace(".0", "") if len(row) > 2 else ""
        position = row[4] if len(row) > 4 else ""
        source   = row[5] if len(row) > 5 else ""

        vals = [fio, phone, city, age, position, source]
        ws.row_dimensions[r_i].height = 15.75

        for c_i, val in enumerate(vals, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.border = border
            cell.fill = white_fill
            cell.alignment = Alignment(horizontal="left", vertical="bottom")
            if c_i == 6:  # ИСТОЧНИК — JetBrains Mono
                cell.font = jetbrains
            else:
                cell.font = arial_bold
            if c_i == 2:  # НОМЕР — текст, без .0 и E+11
                cell.number_format = "@"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Claude extraction ──────────────────────────────────────────────────────────

SYSTEM_PROMPT = """Ты извлекаешь данные из резюме с сайтов поиска работы.

Верни ТОЛЬКО JSON без пояснений, markdown и дополнительного текста:
{
  "fio": "Фамилия Имя Отчество на языке оригинала",
  "phone": "380XXXXXXXXX (только цифры, начиная с 380, без пробелов тире скобок плюсов)",
  "age": "только цифра, например 23",
  "city": "город на русском языке (Днепр, Киев, Харьков и т.д.)",
  "positions": "должность1, должность2 (на языке оригинала)",
  "source": "work.ua или rabota.ua или olx.ua — определи по логотипу, цветам, стилю сайта. Если не можешь — unknown"
}

Правила для телефона:
- 0XX XXX XX XX → 380XXXXXXXXX
- +38 0XX... → 380XXXXXXXXX
- Уже начинается с 380 — оставь как есть
- Только цифры, никаких символов

Если поле не найдено — пустая строка "".
"""

def _parse(raw: str) -> dict:
    raw = re.sub(r"^```json|^```|```$", "", raw, flags=re.MULTILINE).strip()
    return json.loads(raw)

def extract_from_image(image_bytes: bytes, mime: str = "image/jpeg") -> dict:
    b64 = base64.standard_b64encode(image_bytes).decode()
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001", max_tokens=500, system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
            {"type": "text", "text": "Извлеки данные из этого скриншота резюме."}
        ]}],
    )
    return _parse(msg.content[0].text)

def extract_from_pdf(pdf_bytes: bytes) -> dict:
    b64 = base64.standard_b64encode(pdf_bytes).decode()
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001", max_tokens=500, system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
            {"type": "text", "text": "Извлеки данные из этого резюме."}
        ]}],
    )
    return _parse(msg.content[0].text)

def extract_from_text(text: str) -> dict:
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001", max_tokens=500, system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"Резюме:\n{text}"}],
    )
    return _parse(msg.content[0].text)

def extract_from_docx(docx_bytes: bytes) -> dict:
    doc = Document(BytesIO(docx_bytes))
    text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return extract_from_text(text)

# ── Process single item ────────────────────────────────────────────────────────

async def process_item(item: dict, bot) -> dict | None:
    try:
        kind = item["kind"]
        if kind == "photo":
            file = await bot.get_file(item["file_id"])
            raw = bytes(await file.download_as_bytearray())
            data = extract_from_image(raw, "image/jpeg")
        elif kind == "image_doc":
            file = await bot.get_file(item["file_id"])
            raw = bytes(await file.download_as_bytearray())
            mime = "image/png" if item["fname"].endswith(".png") else "image/jpeg"
            data = extract_from_image(raw, mime)
        elif kind == "pdf":
            file = await bot.get_file(item["file_id"])
            raw = bytes(await file.download_as_bytearray())
            data = extract_from_pdf(raw)
        elif kind == "docx":
            file = await bot.get_file(item["file_id"])
            raw = bytes(await file.download_as_bytearray())
            data = extract_from_docx(raw)
        elif kind == "text":
            data = extract_from_text(item["text"])
        else:
            return None
        data["who"] = item.get("who", "")
        return data
    except Exception as e:
        log.exception(f"Error processing item: {e}")
        return None

# ── Batch processor ────────────────────────────────────────────────────────────

async def run_batch(chat_id: int, status_msg, bot):
    items = pending.pop(chat_id, [])
    if not items:
        return

    total = len(items)
    await status_msg.edit_text(f"⏳ Обрабатываю {total} резюме параллельно...")

    tasks = [process_item(item, bot) for item in items]
    results = await asyncio.gather(*tasks)

    ok, errors = [], []
    rows = []
    now = datetime.now().strftime(DATE_FMT)

    for data in results:
        if data:
            ok.append(data)
            rows.append([
                data.get("fio", ""),
                data.get("phone", ""),
                data.get("age", ""),
                data.get("city", ""),
                data.get("positions", ""),
                data.get("source", ""),
                data.get("who", ""),
                now,
            ])
        else:
            errors.append(1)

    if rows:
        try:
            append_rows_batch(rows)
        except Exception as e:
            log.exception(f"Sheet write error: {e}")
            await status_msg.edit_text(f"❌ Ошибка записи в таблицу: {e}")
            return

    report = f"✅ Готово! Обработано {len(ok)}/{total}"
    if errors:
        report += f"\n❌ Ошибок: {len(errors)}"
    if ok:
        report += "\n\n📋 Добавлено в таблицу:"
        for d in ok[:5]:
            report += (
                f"\n• {d.get('fio') or '—'} | {d.get('phone') or '—'}"
                f" | {d.get('age') or '—'} | {d.get('city') or '—'}"
                f"\n  💼 {d.get('positions') or '—'} | 🌐 {d.get('source') or '—'}"
            )
        if len(ok) > 5:
            report += f"\n\n  ...и ещё {len(ok)-5}"
    report += "\n\n📤 Когда готов выгрузить в CRM — нажми /export"

    await status_msg.edit_text(report)

# ── Export handler ─────────────────────────────────────────────────────────────

async def cmd_export(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    msg = await update.message.reply_text("⏳ Формирую файл для CRM...")

    try:
        last = get_last_export(chat_id)
        rows = get_rows_since(last)

        if not rows:
            period = f"с {last.strftime(DATE_FMT)}" if last else "за всё время"
            await msg.edit_text(f"📭 Нет новых резюме {period}.")
            return

        buf = generate_crm_xlsx(rows)
        now = datetime.now()
        fname = f"crm_import_{now.strftime('%d%m%Y_%H%M')}.xlsx"

        period_str = f"с {last.strftime(DATE_FMT)}" if last else "за всё время"
        caption = f"📊 Экспорт для CRM\n{period_str} → {now.strftime(DATE_FMT)}\nКонтактов: {len(rows)}"

        await ctx.bot.send_document(
            chat_id=chat_id,
            document=buf,
            filename=fname,
            caption=caption,
        )
        set_last_export(chat_id, now)
        await msg.delete()

    except Exception as e:
        log.exception(f"Export error: {e}")
        await msg.edit_text(f"❌ Ошибка при экспорте: {e}")

# ── Telegram handlers ──────────────────────────────────────────────────────────

def get_user_name(update: Update) -> str:
    u = update.effective_user
    return u.full_name or u.username or str(u.id)

async def schedule_batch(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    if chat_id in pending_timers:
        pending_timers[chat_id].cancel()

    count = len(pending.get(chat_id, []))
    if count == 1:
        msg = await update.message.reply_text("📥 Получено 1 резюме, жду ещё...")
        ctx.chat_data["status_msg"] = msg
    else:
        try:
            await ctx.chat_data["status_msg"].edit_text(f"📥 Получено {count} резюме, жду ещё...")
        except Exception:
            pass

    async def delayed():
        await asyncio.sleep(BATCH_WAIT)
        status_msg = ctx.chat_data.get("status_msg")
        if status_msg:
            await run_batch(chat_id, status_msg, ctx.bot)

    pending_timers[chat_id] = asyncio.create_task(delayed())

async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    pending.setdefault(update.effective_chat.id, []).append({
        "kind": "photo", "file_id": photo.file_id, "who": get_user_name(update),
    })
    await schedule_batch(update, ctx)

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    fname = (doc.file_name or "").lower()
    mime = doc.mime_type or ""
    ext = Path(fname).suffix

    if ext in (".jpg", ".jpeg") or "jpeg" in mime:
        kind = "image_doc"
    elif ext == ".png" or "png" in mime:
        kind = "image_doc"
    elif ext == ".pdf" or "pdf" in mime:
        kind = "pdf"
    elif ext in (".docx", ".doc") or "word" in mime or "officedocument" in mime:
        kind = "docx"
    else:
        await update.message.reply_text(f"⚠️ Формат {ext} не поддерживается. Жду: jpg, png, pdf, docx.")
        return

    pending.setdefault(update.effective_chat.id, []).append({
        "kind": kind, "file_id": doc.file_id, "fname": fname, "who": get_user_name(update),
    })
    await schedule_batch(update, ctx)

async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if len(text) < 30:
        return
    pending.setdefault(update.effective_chat.id, []).append({
        "kind": "text", "text": text, "who": get_user_name(update),
    })
    await schedule_batch(update, ctx)

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.effective_user.first_name or ""
    keyboard = ReplyKeyboardMarkup(
        [[KeyboardButton("📤 Экспорт в CRM (/export)")]],
        resize_keyboard=True
    )
    await update.message.reply_text(
        f"👋 Привет, {name}!\n\n"
        "Отправь резюме — извлеку данные и запишу в Google Таблицу.\n\n"
        "📎 Форматы: фото, jpg, png, pdf, docx\n"
        "📦 Можно кидать пачками\n\n"
        "📤 /export — выгрузить xlsx для CRM",
        reply_markup=keyboard,
    )

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("export", cmd_export))
    app.add_handler(MessageHandler(filters.Regex(r"Экспорт в CRM"), cmd_export))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    log.info("Bot started")
    app.run_polling()

if __name__ == "__main__":
    main()
