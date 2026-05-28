import os
import asyncio
import base64
import json
import re
import logging
from pathlib import Path
from datetime import datetime

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

TELEGRAM_TOKEN   = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
GOOGLE_SHEET_ID  = os.environ["GOOGLE_SHEET_ID"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]

HEADERS = ["ФИО", "Телефон", "Возраст", "Город", "Должности", "Источник", "Кто добавил", "Дата"]

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# Буфер: chat_id -> список файлов, ожидающих обработки
pending: dict[int, list] = {}
pending_timers: dict[int, asyncio.Task] = {}
BATCH_WAIT = 4  # секунд ждём после последнего файла перед запуском

# ── Google Sheets ──────────────────────────────────────────────────────────────

def get_sheet():
    creds_data = json.loads(GOOGLE_CREDS_JSON)
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_data, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet("Резюме")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Резюме", rows=5000, cols=len(HEADERS))
        ws.append_row(HEADERS)
        ws.format("A1:H1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.26, "green": 0.52, "blue": 0.96}
        })
    return ws

def append_rows_batch(rows: list[list]):
    ws = get_sheet()
    ws.append_rows(rows, value_input_option="USER_ENTERED")

# ── Claude extraction ──────────────────────────────────────────────────────────

SYSTEM_PROMPT = """Ты извлекаешь данные из скриншота или файла резюме с сайтов поиска работы.

Верни ТОЛЬКО JSON без пояснений, markdown и дополнительного текста:
{
  "fio": "Фамилия Имя Отчество на языке оригинала",
  "phone": "380XXXXXXXXX (только цифры, начиная с 380, без пробелов тире скобок плюсов)",
  "age": "только цифра, например 23",
  "city": "город на русском языке (Днепр, Киев, Харьков и т.д.)",
  "positions": "должность1, должность2 (на языке оригинала)",
  "source": "work.ua или rabota.ua или olx.ua — определи по логотипу, цветам, стилю сайта на скриншоте. Если не можешь определить — напиши unknown"
}

Правила для телефона:
- Украинские номера: 0XX XXX XX XX → 380XXXXXXXXX
- Если номер начинается с +38 — убери +38 и оставь 380...
- Если номер начинается с 0 — замени 0 на 380
- Если номер уже начинается с 380 — оставь как есть
- Только цифры, никаких символов

Если поле не найдено — пустая строка "".
"""

def _parse(raw: str) -> dict:
    raw = re.sub(r"^```json|^```|```$", "", raw, flags=re.MULTILINE).strip()
    return json.loads(raw)

def extract_from_image(image_bytes: bytes, mime: str = "image/jpeg") -> dict:
    b64 = base64.standard_b64encode(image_bytes).decode()
    msg = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
            {"type": "text", "text": "Извлеки данные из этого скриншота резюме."}
        ]}],
    )
    return _parse(msg.content[0].text)

def extract_from_pdf(pdf_bytes: bytes) -> dict:
    b64 = base64.standard_b64encode(pdf_bytes).decode()
    msg = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
            {"type": "text", "text": "Извлеки данные из этого резюме."}
        ]}],
    )
    return _parse(msg.content[0].text)

def extract_from_text(text: str) -> dict:
    msg = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"Резюме:\n{text}"}],
    )
    return _parse(msg.content[0].text)

# ── Process single item ────────────────────────────────────────────────────────

async def process_item(item: dict, bot) -> dict | None:
    """Обрабатывает один файл/фото. Возвращает dict с данными или None при ошибке."""
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

    # Параллельная обработка
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

    # Записываем в таблицу одним запросом
    if rows:
        try:
            append_rows_batch(rows)
        except Exception as e:
            log.exception(f"Sheet write error: {e}")
            await status_msg.edit_text(f"❌ Ошибка записи в таблицу: {e}")
            return

    # Итоговый отчёт
    report = f"✅ Готово! Обработано {len(ok)}/{total}"
    if errors:
        report += f"\n❌ Ошибок: {len(errors)}"
    if ok:
        report += "\n\n📋 Добавлено в таблицу:"
        for d in ok[:5]:  # Показываем первые 5
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
    name = u.full_name or u.username or str(u.id)
    return name

async def schedule_batch(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Откладывает запуск батча — ждёт BATCH_WAIT сек после последнего файла."""
    chat_id = update.effective_chat.id

    # Отменяем предыдущий таймер если есть
    if chat_id in pending_timers:
        pending_timers[chat_id].cancel()

    # Показываем/обновляем статус
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

    task = asyncio.create_task(delayed())
    pending_timers[chat_id] = task


async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    photo = update.message.photo[-1]
    pending.setdefault(chat_id, []).append({
        "kind": "photo",
        "file_id": photo.file_id,
        "who": get_user_name(update),
    })
    await schedule_batch(update, ctx)


async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
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
    else:
        await update.message.reply_text(f"⚠️ Формат {ext} не поддерживается. Жду: jpg, png, pdf.")
        return

    pending.setdefault(chat_id, []).append({
        "kind": kind,
        "file_id": doc.file_id,
        "fname": fname,
        "who": get_user_name(update),
    })
    await schedule_batch(update, ctx)


async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if len(text) < 30:
        return
    chat_id = update.effective_chat.id
    pending.setdefault(chat_id, []).append({
        "kind": "text",
        "text": text,
        "who": get_user_name(update),
    })
    await schedule_batch(update, ctx)


async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.effective_user.first_name or "!"
    await update.message.reply_text(
        f"👋 Привет, {name}!\n\n"
        "Отправь мне скриншоты или PDF резюме — я автоматически извлеку данные "
        "и запишу в Google Таблицу.\n\n"
        "📎 Поддерживаемые форматы: фото, jpg, png, pdf\n"
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
