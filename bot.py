"""
Lid -> Excel bot
----------------
Lid xabarini (🧾 Nomi: ... 📝 Ma'lumotlar: ...) botga tashlaysiz,
bot uni ajratib oladi va bitta Excel faylga qator qilib qo'shadi.

Env o'zgaruvchilar (Render > Environment):
  BOT_TOKEN   - BotFather bergan token (majburiy)
  ADMIN_IDS   - botdan foydalana oladiganlar ID si, vergul bilan: 123456789,987654321
                (bo'sh qolsa, hamma foydalana oladi; ID ni /start bilan bilib olasiz)
  SELF_URL    - Render'dagi bot linki (masalan https://lid-bot.onrender.com)
                bo'lsa, bot uxlab qolmasligi uchun o'ziga har 9 daqiqada ping yuboradi.
                (Render o'zi RENDER_EXTERNAL_URL ni ham beradi - SELF_URL bo'lmasa o'shani oladi)
  SEND_DELAY  - oxirgi lid tushgandan keyin necha soniyada Excel yuborilsin (standart 4)
"""

import asyncio
import io
import logging
import os
import re
import threading
import time
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger("lidbot")

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
ADMIN_IDS = {int(x) for x in re.findall(r"\d+", os.environ.get("ADMIN_IDS", ""))}
SELF_URL = os.environ.get("SELF_URL") or os.environ.get("RENDER_EXTERNAL_URL", "")
SEND_DELAY = float(os.environ.get("SEND_DELAY", "4"))
LOCAL_FILE = "lidlar.xlsx"
DATE_FMT = "%d-%m-%Y %H:%M:%S"

# ---------------------------------------------------------------- Excel tuzilishi
# (kalit, sarlavha, kenglik)
INFO_COLS = [
    ("daraja", "Daraja", 12),
    ("kurs", "Kurs", 16),
    ("vaqt", "Vaqt", 14),
    ("ism", "Ism-familiya", 26),
    ("tel", "Telefon", 17),
]
LEFT_COLS = [("n", "№", 6), ("nomi", "Forma nomi", 28)]
RIGHT_COLS = [("manba", "Manba", 13), ("sana", "Sana", 19)]
ALL_COLS = LEFT_COLS + INFO_COLS + RIGHT_COLS

rows: list[dict] = []  # barcha lidlar (xotirada)

# ---------------------------------------------------------------- Xabarni o'qish
QUESTION_RE = re.compile(r"^\s*\d+\s*[.)]\s*(.+?)\s*:\s*(.*)$")


def field_for(question: str) -> str | None:
    q = question.lower()
    if "daraja" in q or "level" in q:
        return "daraja"
    if "kurs" in q or "course" in q:
        return "kurs"
    if "vaqt" in q or "time" in q:
        return "vaqt"
    if "phone" in q or "telefon" in q or "raqam" in q or "nomer" in q:
        return "tel"
    if "name" in q or "ism" in q:
        return "ism"
    return None


def clean_phone(p: str) -> str:
    digits = re.sub(r"\D", "", p)
    if len(digits) == 9:  # 900341519 -> +998900341519
        digits = "998" + digits
    return "+" + digits if digits else p.strip()


def parse_block(lines: list[str]) -> dict | None:
    lead: dict = {}
    for line in lines:
        m = QUESTION_RE.match(line)
        if m:
            key = field_for(m.group(1))
            if key and m.group(2).strip():
                lead[key] = m.group(2).strip()
            continue
        for key, label in (("nomi", "Nomi:"), ("manba", "Manba:"), ("sana", "Sana:")):
            if label in line:
                lead[key] = line.split(label, 1)[1].strip()
    if not (lead.get("tel") or lead.get("ism")):
        return None
    if lead.get("tel"):
        lead["tel"] = clean_phone(lead["tel"])
    if lead.get("sana"):
        try:
            lead["sana"] = datetime.strptime(lead["sana"], DATE_FMT)
        except ValueError:
            pass
    return lead


def parse_message(text: str) -> list[dict]:
    """Bitta xabarda bir nechta lid bo'lsa ham ajratadi ('Nomi:' qatoridan boshlab)."""
    blocks, cur = [], []
    for line in text.splitlines():
        if "Nomi:" in line and cur:
            blocks.append(cur)
            cur = []
        cur.append(line)
    if cur:
        blocks.append(cur)
    return [l for l in (parse_block(b) for b in blocks) if l]


def lead_key(lead: dict) -> tuple:
    sana = lead.get("sana")
    if isinstance(sana, datetime):
        sana = sana.strftime(DATE_FMT)
    return (re.sub(r"\D", "", str(lead.get("tel", ""))), str(sana or ""), str(lead.get("ism", "")).lower())


def add_leads(new: list[dict]) -> tuple[int, int]:
    known = {lead_key(r) for r in rows}
    added = dup = 0
    for lead in new:
        k = lead_key(lead)
        if k in known:
            dup += 1
            continue
        known.add(k)
        rows.append(lead)
        added += 1
    return added, dup


# ---------------------------------------------------------------- Excel yasash
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FILL = PatternFill("solid", fgColor="2E75B6")
ZEBRA = PatternFill("solid", fgColor="EAF1FB")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=11)
THIN = Side(style="thin", color="9BB7D4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def build_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Lidlar"

    info_start = len(LEFT_COLS) + 1
    info_end = info_start + len(INFO_COLS) - 1

    # 1-qator: "Ma'lumotlar" umumiy sarlavha, qolganlari 2 qatorga birlashtirilgan
    for i, (_, title, width) in enumerate(ALL_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
        in_info = info_start <= i <= info_end
        if not in_info:
            ws.cell(row=1, column=i, value=title)
            ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
        else:
            ws.cell(row=2, column=i, value=title)
    ws.cell(row=1, column=info_start, value="Ma'lumotlar")
    ws.merge_cells(start_row=1, start_column=info_start, end_row=1, end_column=info_end)

    for r in (1, 2):
        for c in range(1, len(ALL_COLS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = SUB_FILL if (r == 2 and info_start <= c <= info_end) else HEAD_FILL
            cell.font = WHITE_BOLD
            cell.alignment = CENTER
            cell.border = BORDER
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 20

    # Ma'lumot qatorlari
    for idx, lead in enumerate(rows, start=1):
        r = idx + 2
        for c, (key, _, _) in enumerate(ALL_COLS, start=1):
            val = idx if key == "n" else lead.get(key, "")
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = BORDER
            cell.alignment = CENTER if key in ("n", "daraja", "vaqt", "manba", "sana", "tel") else LEFT
            if key == "tel":
                cell.number_format = "@"
            if key == "sana" and isinstance(val, datetime):
                cell.number_format = "dd.mm.yyyy hh:mm"
            if idx % 2 == 0:
                cell.fill = ZEBRA

    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    last = len(rows) + 2
    ws.auto_filter.ref = f"A2:{get_column_letter(len(ALL_COLS))}{max(last, 2)}"

    buf = io.BytesIO()
    wb.save(buf)
    data = buf.getvalue()
    try:
        with open(LOCAL_FILE, "wb") as f:
            f.write(data)
    except OSError:
        pass
    return data


def load_workbook_bytes(data: bytes) -> list[dict]:
    """Bot yasagan Excel'ni qayta o'qiydi (tiklash uchun)."""
    ws = load_workbook(io.BytesIO(data)).active
    loaded = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or all(v in (None, "") for v in row):
            continue
        lead = {}
        for (key, _, _), val in zip(ALL_COLS, row):
            if key != "n" and val not in (None, ""):
                lead[key] = val if isinstance(val, datetime) else str(val)
        if lead:
            loaded.append(lead)
    return loaded


# ---------------------------------------------------------------- Telegram
pending: dict[int, asyncio.Task] = {}
last_pinned: dict[int, int] = {}


def allowed(update: Update) -> bool:
    return not ADMIN_IDS or (update.effective_user and update.effective_user.id in ADMIN_IDS)


async def send_excel(chat_id: int, bot) -> None:
    if not rows:
        await bot.send_message(chat_id, "Hozircha lid yo'q.")
        return
    data = build_workbook()
    name = f"lidlar_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
    msg = await bot.send_document(
        chat_id, document=data, filename=name, caption=f"📊 Jami: {len(rows)} ta lid"
    )
    # Eng so'nggi faylni pin qilamiz - server qayta ishga tushsa, shundan tiklanadi
    try:
        if chat_id in last_pinned:
            await bot.unpin_chat_message(chat_id, message_id=last_pinned[chat_id])
        await bot.pin_chat_message(chat_id, msg.message_id, disable_notification=True)
        last_pinned[chat_id] = msg.message_id
    except Exception as e:  # pin bo'lmasa ham ish to'xtamaydi
        log.warning("Pin xatosi: %s", e)


async def _delayed_send(chat_id: int, bot) -> None:
    try:
        await asyncio.sleep(SEND_DELAY)
    except asyncio.CancelledError:
        return
    await send_excel(chat_id, bot)


def schedule_send(chat_id: int, bot) -> None:
    """Ketma-ket bir nechta lid tashlansa, Excel faqat bir marta - oxirida yuboriladi."""
    t = pending.get(chat_id)
    if t and not t.done():
        t.cancel()
    pending[chat_id] = asyncio.create_task(_delayed_send(chat_id, bot))


async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id
    if not allowed(update):
        await update.message.reply_text(f"Sizga ruxsat yo'q. ID: {uid}")
        return
    await update.message.reply_text(
        "Assalamu alaykum! Lid xabarlarini shu yerga tashlang (forward yoki nusxa) — "
        "men ularni Excel jadvalga qo'shib boraman.\n\n"
        "/excel — jadvalni hozir olish\n"
        "/soni — nechta lid borligi\n"
        "/tozalash — hammasini o'chirish\n\n"
        "Eski Excel faylni yuborsangiz, undagi lidlarni ham qo'shib oladi.\n\n"
        f"Sizning ID: {uid}"
    )


async def cmd_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if allowed(update):
        await send_excel(update.effective_chat.id, ctx.bot)


async def cmd_count(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if allowed(update):
        await update.message.reply_text(f"Jami: {len(rows)} ta lid")


async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return
    if ctx.args and ctx.args[0].lower() == "ha":
        rows.clear()
        await update.message.reply_text("🗑 Hamma lidlar o'chirildi.")
    else:
        await update.message.reply_text(
            f"Rostdan ham {len(rows)} ta lidni o'chiraymi? Tasdiqlash uchun yozing:\n/tozalash ha"
        )


async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return
    text = update.message.text or update.message.caption or ""
    leads = parse_message(text)
    if not leads:
        await update.message.reply_text("Bu xabardan lid topolmadim (ism yoki telefon yo'q).")
        return
    added, dup = add_leads(leads)
    parts = []
    if added:
        names = ", ".join(str(l.get("ism", "?")) for l in leads[-added:])
        parts.append(f"✅ Qo'shildi: {added} ta ({names}) — jami {len(rows)}")
    if dup:
        parts.append(f"↩️ Takror, o'tkazib yuborildi: {dup} ta")
    await update.message.reply_text("\n".join(parts))
    if added:
        schedule_send(update.effective_chat.id, ctx.bot)


async def on_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return
    doc = update.message.document
    if not (doc.file_name or "").lower().endswith(".xlsx"):
        return
    f = await doc.get_file()
    data = bytes(await f.download_as_bytearray())
    try:
        added, dup = add_leads(load_workbook_bytes(data))
    except Exception as e:
        await update.message.reply_text(f"Faylni o'qib bo'lmadi: {e}")
        return
    await update.message.reply_text(f"📥 Fayldan qo'shildi: {added} ta, takror: {dup} ta. Jami: {len(rows)}")
    if added:
        schedule_send(update.effective_chat.id, ctx.bot)


async def restore(app: Application) -> None:
    """Ishga tushganda: avval lokal fayl, bo'lmasa adminning chatidagi pin qilingan Exceldan tiklaydi."""
    if os.path.exists(LOCAL_FILE):
        with open(LOCAL_FILE, "rb") as f:
            add_leads(load_workbook_bytes(f.read()))
        log.info("Lokal fayldan tiklandi: %d", len(rows))
        return
    for admin in ADMIN_IDS:
        try:
            chat = await app.bot.get_chat(admin)
            pm = chat.pinned_message
            if pm and pm.document and (pm.document.file_name or "").endswith(".xlsx"):
                f = await pm.document.get_file()
                add_leads(load_workbook_bytes(bytes(await f.download_as_bytearray())))
                last_pinned[admin] = pm.message_id
                log.info("Pin qilingan fayldan tiklandi (%s): %d", admin, len(rows))
        except Exception as e:
            log.warning("Tiklashda xato (%s): %s", admin, e)


# ---------------------------------------------------------------- Render uchun: web + self-ping
class Health(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"OK")

    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, *a):
        pass


def run_web():
    port = int(os.environ.get("PORT", "10000"))
    HTTPServer(("0.0.0.0", port), Health).serve_forever()


def self_ping():
    while True:
        time.sleep(9 * 60)
        try:
            urllib.request.urlopen(SELF_URL, timeout=20).read()
        except Exception as e:
            log.warning("Ping xatosi: %s", e)


def main():
    if not BOT_TOKEN:
        raise SystemExit("BOT_TOKEN qo'yilmagan")
    threading.Thread(target=run_web, daemon=True).start()
    if SELF_URL:
        threading.Thread(target=self_ping, daemon=True).start()

    app = Application.builder().token(BOT_TOKEN).post_init(restore).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("excel", cmd_excel))
    app.add_handler(CommandHandler("soni", cmd_count))
    app.add_handler(CommandHandler("tozalash", cmd_clear))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
