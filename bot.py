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

from guruhlar import GURUHLAR, KURS_TURI, TAKLIF_SONI, USTOZLAR, VAQT_ORALIQ, XONA_SIGIMI, XONALAR
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger("lidbot")
logging.getLogger("httpx").setLevel(logging.WARNING)  # token loglarda chiqmasin

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
    ("ism", "Ism-familiya", 24),
    ("tel", "Telefon", 17),
]
LEFT_COLS = [("n", "№", 6)]
RIGHT_COLS = [("taklif", "Mos guruhlar (taklif)", 66), ("manba", "Manba", 13), ("sana", "Sana", 19)]
ALL_COLS = LEFT_COLS + INFO_COLS + RIGHT_COLS

rows: list[dict] = []  # barcha lidlar (xotirada)

# ---------------------------------------------------------------- Xabarni o'qish
QUESTION_RE = re.compile(r"^\s*\d+\s*[.)]\s*(.+?)\s*:\s*(.*)$")


IGNORED = ("form name", "ad name", "campaign", "adset", "ad set", "forma", "e'lon")


def field_for(question: str) -> str | None:
    q = question.lower().strip()
    if any(w in q for w in IGNORED):
        return None
    if "daraja" in q or "level" in q:
        return "daraja"
    if "kurs" in q or "course" in q:
        return "kurs"
    if "vaqt" in q or "time" in q:
        return "vaqt"
    if "phone" in q or "telefon" in q or "raqam" in q or "nomer" in q:
        return "tel"
    if "full name" in q or "ism" in q or "fio" in q or "f.i.sh" in q or q == "name":
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
            if key and m.group(2).strip() and key not in lead:
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


# ---------------------------------------------------------------- Guruhga moslash
KUN_NOMI = {"toq": "Du/Chor/Ju", "juft": "Se/Pay/Shan"}


def to_min(hhmm: str) -> int:
    h, m = str(hhmm).strip().split(":")
    return int(h) * 60 + int(m)


def fmt_min(m: int) -> str:
    return f"{m // 60:02d}:{m % 60:02d}"


def vaqt_of(start: int) -> str | None:
    for name, (a, b) in VAQT_ORALIQ.items():
        if to_min(a) <= start < to_min(b):
            return name
    return None


def g_interval(g: dict) -> tuple[int, int]:
    s = to_min(g["soat"])
    return s, s + int(g.get("davom", 120))


def slot_interval(x: dict) -> tuple[int, int]:
    a, b = x["vaqt"].split("-")
    return to_min(a), to_min(b)


def overlap(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return a[0] < b[1] and b[0] < a[1]


def slot_conflict(x: dict) -> dict | None:
    """Shu vaqtda xonada dars o'tayotgan guruh (bo'lsa)."""
    for g in GURUHLAR:
        if g["xona"] == x["xona"] and g["kunlar"] == x["kunlar"] and overlap(g_interval(g), slot_interval(x)):
            return g
    return None


def teacher_busy(ustoz: str, kunlar: str, iv: tuple[int, int]) -> bool:
    return any(g["ustoz"] == ustoz and g["kunlar"] == kunlar and overlap(g_interval(g), iv) for g in GURUHLAR)


def norm_daraja(s: str) -> str | None:
    s = str(s or "").lower().replace(" ", "")
    m = re.search(r"[abc][0-2]", s)
    if m:
        return m.group(0)
    if "boshlang" in s or "nol" in s or "0dan" in s:
        return "a1"
    if "o'rta" in s or "orta" in s or "o‘rta" in s:
        return "b1"
    return None


def norm_kurs(s: str) -> str:
    s = str(s or "").lower()
    for key, tur in KURS_TURI.items():
        if key in s:
            return tur
    return "umumiy"


def norm_vaqt(s: str) -> str | None:
    s = str(s or "").lower()
    if "ertal" in s:
        return "ertalab"
    if "kundu" in s or "tush" in s or "peshin" in s or "obed" in s:
        return "kunduzi"
    if "kech" in s or "oqshom" in s:
        return "kechki"
    return None


def match_groups(lead: dict) -> list[tuple[dict, list[str]]]:
    """Lidga eng mos mavjud guruhlar: [(guruh, [mos kelmagan narsalar]), ...]"""
    lvl, tur, vq = norm_daraja(lead.get("daraja")), norm_kurs(lead.get("kurs")), norm_vaqt(lead.get("vaqt"))
    scored = []
    for g in GURUHLAR:
        ok_kurs = g["tur"] == tur
        ok_lvl = lvl is None or lvl in g["daraja"]
        ok_vaqt = vq is None or vaqt_of(to_min(g["soat"])) == vq
        score = 4 * ok_kurs + 2 * ok_lvl + 2 * ok_vaqt
        if score < 4:  # kurs ham, daraja+vaqt ham mos emas - taklif qilinmaydi
            continue
        miss = [name for ok, name in ((ok_kurs, "kurs"), (ok_lvl, "daraja"), (ok_vaqt, "vaqt")) if not ok]
        scored.append((score, g.get("talaba", 0), g, miss))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [(g, miss) for _, _, g, miss in scored[:TAKLIF_SONI]]


def free_slots() -> list[tuple[dict, dict | None]]:
    """Barcha bo'sh xona vaqtlari va (agar bo'lsa) ular bilan to'qnashgan guruh."""
    return [(x, slot_conflict(x)) for x in XONALAR]


def new_group_options(lead: dict) -> list[tuple[dict, str]]:
    """Yangi guruh ochish uchun: [(bo'sh xona vaqti, ustoz), ...]"""
    tur, vq = norm_kurs(lead.get("kurs")), norm_vaqt(lead.get("vaqt"))
    opts, seen = [], set()
    for x, conflict in free_slots():
        if conflict:
            continue
        iv = slot_interval(x)
        slot_vaqt = vaqt_of(iv[0])
        if vq and slot_vaqt != vq:
            continue
        candidates = [x["ustoz"]] if x.get("ustoz") else list(USTOZLAR)
        for t in candidates:
            info = USTOZLAR.get(t)
            if not info or tur not in info["turlar"] or slot_vaqt not in info["vaqtlar"]:
                continue
            if teacher_busy(t, x["kunlar"], iv):
                continue
            key = (t, x["kunlar"], x["vaqt"])  # bir ustozni bir vaqtga bir marta taklif qilamiz
            if key in seen:
                continue
            seen.add(key)
            prio = 0 if x.get("ustoz") else 1  # ustozga ajratilgan vaqt birinchi
            opts.append((prio, -XONA_SIGIMI.get(x["xona"], 0), x, t))
            break
    opts.sort(key=lambda o: (o[0], o[1]))
    return [(x, t) for _, _, x, t in opts[:TAKLIF_SONI]]


def group_label(g: dict) -> str:
    s, e = g_interval(g)
    return (f"{g['ustoz']} · {g['nomi']} — {g['izoh']}, {KUN_NOMI.get(g['kunlar'], g['kunlar'])} "
            f"{fmt_min(s)}–{fmt_min(e)}, {g['xona']}, {g.get('talaba', '?')} talaba")


def slot_label(x: dict, ustoz: str) -> str:
    sig = XONA_SIGIMI.get(x["xona"])
    sig_t = f" ({sig} kishilik)" if sig else ""
    return f"{x['xona']}{sig_t} · {x['kunlar'].capitalize()} ({KUN_NOMI.get(x['kunlar'])}) {x['vaqt']} · {ustoz} Ustoz"


def _suggest(lead: dict) -> tuple[list, list]:
    found = match_groups(lead)
    perfect = any(not miss for _, miss in found)
    new = [] if perfect else new_group_options(lead)
    return found, new


def taklif_text(lead: dict) -> str:
    """Excel ustuni uchun."""
    found, new = _suggest(lead)
    lines = []
    for i, (g, miss) in enumerate(found, 1):
        note = f" ({', '.join(miss)} mos emas)" if miss else ""
        lines.append(f"{i}. {group_label(g)}{note}")
    for x, t in new:
        lines.append(f"Yangi guruh: {slot_label(x, t)}")
    return "\n".join(lines) or "Mos guruh yo'q"


def taklif_message(lead: dict) -> str:
    """Bot javobi uchun."""
    head = f"👤 {lead.get('ism', '?')} ({lead.get('daraja', '?')} · {lead.get('kurs', '?')} · {lead.get('vaqt', '?')})"
    found, new = _suggest(lead)
    lines = [head]
    if found:
        lines.append("Mos guruhlar:")
        for i, (g, miss) in enumerate(found, 1):
            mark = "✅" if not miss else "⚠️"
            note = f" ({', '.join(miss)} mos emas)" if miss else ""
            lines.append(f"{mark} {i}. {group_label(g)}{note}")
    if new:
        lines.append("🆕 Yangi guruh ochish mumkin:")
        for x, t in new:
            lines.append(f"• {slot_label(x, t)}")
    if not found and not new:
        lines.append("❗ Mos guruh ham, bo'sh xona-ustoz ham topilmadi.")
    return "\n".join(lines)


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
            if key == "n":
                val = idx
            elif key == "taklif":
                val = taklif_text(lead)
            else:
                val = lead.get(key, "")
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
    """Bot yasagan Excel'ni qayta o'qiydi (tiklash uchun) - ustunlarni sarlavha nomidan taniydi."""
    ws = load_workbook(io.BytesIO(data)).active
    by_title = {title: key for key, title, _ in ALL_COLS if key not in ("n", "taklif")}
    colmap = {}
    for c in range(1, ws.max_column + 1):
        title = ws.cell(row=2, column=c).value or ws.cell(row=1, column=c).value
        if title in by_title:
            colmap[c] = by_title[title]
    loaded = []
    for r in range(3, ws.max_row + 1):
        lead = {}
        for c, key in colmap.items():
            val = ws.cell(row=r, column=c).value
            if val not in (None, ""):
                lead[key] = val if isinstance(val, datetime) else str(val)
        if lead.get("tel") or lead.get("ism"):
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
        "/guruhlar — guruhlar ro'yxati\n"
        "/xonalar — bo'sh xonalar\n"
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


async def cmd_groups(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return
    lines, last = [], None
    for g in sorted(GURUHLAR, key=lambda g: (g["ustoz"], g["soat"])):
        if g["ustoz"] != last:
            lines.append(f"\n👨‍🏫 {g['ustoz']} Ustoz")
            last = g["ustoz"]
        lines.append("• " + group_label(g).split(" · ", 1)[1])
    await update.message.reply_text("📚 Guruhlar:" + "\n".join(lines))


async def cmd_rooms(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return
    lines, last = [], None
    for x, conflict in free_slots():
        if x["xona"] != last:
            sig = XONA_SIGIMI.get(x["xona"])
            lines.append(f"\n🚪 {x['xona']}" + (f" ({sig} kishilik)" if sig else ""))
            last = x["xona"]
        if conflict:
            st = f"❌ band: {conflict['ustoz']} · {conflict['nomi']} guruhi"
        elif x.get("ustoz"):
            st = f"🔸 {x['ustoz']} uchun" + (f" ({x['izoh']})" if x.get("izoh") else "")
        else:
            st = "✅ bo'sh"
        lines.append(f"• {x['kunlar'].capitalize()} {x['vaqt']} — {st}")
    await update.message.reply_text("🗓 Bo'sh xonalar:" + "\n".join(lines))


async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return
    if ctx.args and ctx.args[0].lower() == "ha":
        rows.clear()
        try:
            os.remove(LOCAL_FILE)
        except OSError:
            pass
        try:
            await ctx.bot.unpin_all_chat_messages(update.effective_chat.id)
        except Exception:
            pass
        last_pinned.pop(update.effective_chat.id, None)
        await update.message.reply_text("🗑 Hamma lidlar o'chirildi. Endi lidlarni qaytadan tashlashingiz mumkin.")
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
        parts.append(f"✅ Qo'shildi: {added} ta — jami {len(rows)}")
        for lead in rows[-added:]:
            parts.append(taklif_message(lead))
    if dup:
        parts.append(f"↩️ Takror, o'tkazib yuborildi: {dup} ta")
    await update.message.reply_text("\n\n".join(parts))
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

    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .connect_timeout(30)
        .read_timeout(30)
        .write_timeout(30)
        .get_updates_connect_timeout(30)
        .get_updates_read_timeout(30)
        .post_init(restore)
        .build()
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("excel", cmd_excel))
    app.add_handler(CommandHandler("soni", cmd_count))
    app.add_handler(CommandHandler("tozalash", cmd_clear))
    app.add_handler(CommandHandler("guruhlar", cmd_groups))
    app.add_handler(CommandHandler("xonalar", cmd_rooms))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
