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
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

TELEGRAM_TOKEN    = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
GOOGLE_SHEET_ID   = os.environ["GOOGLE_SHEET_ID"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]

HEADERS = ["ФИО", "Телефон", "Возраст", "Город", "Должности", "Источник", "Кто добавил", "Дата"]

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

pending: dict[int, list] = {}
pending_timers: dict[int, asyncio.Task] = {}
BATCH_WAIT = 4

# ── Google Sheets ──────────────────────────────────────────────────────────────

def get_sheet():
    creds_data = json.loads(GOOGLE_CREDS_JSON)
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_data, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet("Резюме")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Резюме", rows=5000, cols=len(HEADERS))
        ws.append_row(HEADERS)
        ws.format("A1:H1", {"textFormat": {"bold": True}})
    return ws

def append_rows_batch(rows: list[list]):
    ws = get_sheet()
    ws.append_rows(rows, value_input_option="USER_ENTERED")

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
        model="claude-opus-4-5", max_tokens=500, system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
            {"type": "text", "text": "Извлеки данные из этого скриншота резюме."}
        ]}],
    )
    return _parse(msg.content[0].text)

def extract_from_pdf(pdf_bytes: bytes) -> dict:
    b64 = base64.standard_b64encode(pdf_bytes).decode()
    msg = client.messages.create(
        model="claude-opus-4-5", max_tokens=500, system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
            {"type": "text", "text": "Извлеки данные из этого резюме."}
        ]}],
    )
    return _parse(msg.content[0].text)

def extract_from_text(text: str) -> dict:
    msg = client.messages.create(
        model="claude-opus-4-5", max_tokens=500, system=SYSTEM_PROMPT,
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
    now = datetime.now().strftime("%d.%m.%Y %H:%M")

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

    await status_msg.edit_text(report)

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
        msg = await update.message.reply_text(f"📥 Получено 1 резюме, жду ещё...")
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
    await update.message.reply_text(
        f"👋 Привет, {name}!\n\n"
        "Отправь мне резюме — извлеку данные и запишу в Google Таблицу.\n\n"
        "📎 Форматы: фото, jpg, png, pdf, docx\n"
        "📦 Можно кидать пачками — обработаю всё сразу!"
    )

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    log.info("Bot started")
    app.run_polling()

if __name__ == "__main__":
    main()
    
