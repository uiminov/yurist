# -*- coding: utf-8 -*-
"""
TELEGRAM BOT — sud arizalarini avtomatik shakllantirish.

Ishlash tartibi:
  1. Foydalanuvchi /start bosadi — bot Excel ro'yxatni so'raydi.
  2. Foydalanuvchi Excel (.xlsx) faylni yuboradi.
  3. Bot har bir qarzdor uchun PDF arizasini tayyorlab qaytaradi
     (1 ta bo'lsa — .pdf, ko'p bo'lsa — .zip arxiv).

Ishga tushirish:  BOT.bat  (yoki:  python bot.py)
Kutubxonalar:     pyTelegramBotAPI, python-docx, openpyxl, pywin32 (PDF uchun)
"""

import io
import os
import re
import sys
import html
import json
import time
import glob
import shutil
import zipfile
import logging
import tempfile
import datetime
import threading
import traceback

# stderr тоже: в него пишет logging, и кириллица ломала вывод в консоли Windows
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    import telebot
    from telebot import types
except ImportError:
    print("Kutubxona o'rnatilmagan. Buyruq satrida bajaring:")
    print("    pip install pyTelegramBotAPI python-docx openpyxl")
    input("\nEnter...")
    sys.exit(1)

import generate_sud as gs          # generate_sud.py shu papkada bo'lishi shart
import openpyxl


# ═══════════════════════════════════════════════════════════════════
#  SOZLAMALAR
# ═══════════════════════════════════════════════════════════════════

# Bot tokeni (@BotFather). Xohlasangiz TELEGRAM_TOKEN muhit o'zgaruvchisi orqali ham.
TOKEN = os.environ.get("TELEGRAM_TOKEN", "8687284274:AAEvsB4Ktyx9q5yDM9iE2V3-YX6h5zuYjFk")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "bot_sozlamalari.json")
LOG_FILE = os.path.join(BASE_DIR, "bot.log")

# Ruxsat etilgan foydalanuvchilar (Telegram ID). Bo'sh ro'yxat = hamma uchun ochiq.
# O'z ID'ingizni bilish uchun botga /id yozing va bu yerga qo'shing:  [123456789]
RUXSAT_ETILGANLAR = []

# Standart xarajatlar
DEFAULT_SUD = 100_000            # Excelda "суд харажати" bo'sh bo'lsa
DEFAULT_POCHTA = 22_000          # har doim; Exceldagi "почта харажати" hisobga olinmaydi

# Telegram cheklovlari
MAX_ZIP_MB = 45                  # bitta arxiv hajmi
MAX_FILES_PER_ZIP = 150          # bitta arxivdagi hujjatlar soni
MAX_ROWS = 3000                  # bitta ro'yxatdagi qatorlar chegarasi
MAX_DOWNLOAD_MB = 20             # Bot API orqali yuklab olish chegarasi
MAX_MESSAGE_LEN = 4096           # bitta xabardagi belgilar chegarasi
SEND_TIMEOUT = 180               # katta arxivni yuborish uchun kutish (soniya)

# ═══════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"),
              logging.StreamHandler()],
)
log = logging.getLogger("sudbot")

bot = telebot.TeleBot(TOKEN, parse_mode="HTML")


# ── Sozlamalar (har bir chat uchun alohida) ────────────────────────

def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return {}
    try:
        with io.open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        log.warning("Sozlamalar fayli buzilgan (%s) — yangidan boshlanadi", e)
        try:
            os.replace(SETTINGS_FILE, SETTINGS_FILE + ".buzilgan")
        except Exception:
            pass
        return {}


def save_settings(data):
    try:
        tmp = SETTINGS_FILE + ".tmp"
        with io.open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, SETTINGS_FILE)     # atomik: uzilib qolsa ham fayl buzilmaydi
    except Exception as e:
        log.warning("Sozlamalarni saqlab bo'lmadi: %s", e)


SETTINGS = load_settings()
# Обработчики работают в разных потоках — общий словарь под замком
_SETTINGS_LOCK = threading.Lock()


def cfg(chat_id, key, default=None):
    with _SETTINGS_LOCK:
        return SETTINGS.get(str(chat_id), {}).get(key, default)


def set_cfg(chat_id, key, value):
    with _SETTINGS_LOCK:
        SETTINGS.setdefault(str(chat_id), {})[key] = value
        save_settings(SETTINGS)


# ── Shablonlar ─────────────────────────────────────────────────────

def templates():
    """Papkadagi barcha .docx shablonlar (tayyor hujjatlar papkasidan tashqari)."""
    out = []
    for f in sorted(glob.glob(os.path.join(BASE_DIR, "*.docx"))):
        if os.path.basename(f).startswith("~$"):
            continue
        out.append(f)
    return out


def current_template(chat_id):
    name = cfg(chat_id, "shablon")
    if name:
        p = os.path.join(BASE_DIR, name)
        if os.path.exists(p):
            return p
    t = templates()
    return t[0] if t else None


def template_kind(path):
    return gs.detect_template(gs.doc_text(gs.Document(path)))


def template_kind_safe(path):
    """Шаблон может быть повреждён или открыт в Word — обработчик падать не должен."""
    try:
        return template_kind(path)
    except Exception as e:
        log.warning("Shablonni o'qib bo'lmadi (%s): %s", path, e)
        return None


def kind_label(kind):
    if kind is None:
        return "aniqlanmadi — faylni o'qib bo'lmadi"
    return "yangi (qavsli belgilar)" if kind == "new" else "eski (BOSH HARFLI belgilar)"


def has_markers(text):
    """У шаблона должны быть метки. Иначе прислали не шаблон, а готовую аризу."""
    low = text.lower()
    return any(w.lower() in low for w in gs.LEFTOVERS_NEW + gs.LEFTOVERS_OLD)


# ── Ruxsat ─────────────────────────────────────────────────────────

def allowed(message):
    if not RUXSAT_ETILGANLAR:
        return True
    return message.from_user.id in RUXSAT_ETILGANLAR


def deny(message):
    bot.reply_to(message,
                 "⛔️ Kechirasiz, sizda ushbu botdan foydalanish huquqi yo'q.\n"
                 f"Sizning ID: <code>{message.from_user.id}</code>")


# ── Xato matnlarini o'zbekchaga o'girish ───────────────────────────

_UZ = [
    ("В Excel не найдены обязательные колонки:", "Excelda majburiy ustunlar topilmadi:"),
    ("Заголовки в файле:", "Fayldagi sarlavhalar:"),
    ("«жами карздорликк»", "«жами карздорликк» ustuni"),
    ("а остаток+суд харажати", "hisob-kitob (qoldiq + sud xarajati)"),
    ("взято", "hujjatga olindi"),
    ("не заполнено", "to'ldirilmagan"),
    ("строка", "qator"),
    ("Лист:", "Varaq:"),
    ("не заменено", "almashtirilmadi"),
    ("ОШИБКА", "XATO"),
    ("fio", "ФИО"),
    ("ostatok", "Остаток долга"),
]


def uz(text):
    t = str(text)
    # «строка 16 (...)» -> «16-qator (...)»: takror ogohlantirishlar bilan bir xil ko'rinish
    t = re.sub(r"строка\s+(\d+)", r"\1-qator", t)
    for a, b in _UZ:
        t = t.replace(a, b)
    return t


def esc(value):
    """parse_mode=HTML: символы < > & в ФИО/имени файла/тексте ошибки ломают
    отправку целиком (400: can't parse entities). Экранируем всё подставляемое."""
    return html.escape(str(value), quote=False)


# ── Matnlar ────────────────────────────────────────────────────────

SALOM = (
    "🏛 <b>Sud arizalari — avtomatik shakllantirish</b>\n\n"
    "Men qarzdorlar ro'yxatidan har biri uchun alohida sud arizasini "
    "(PDF hujjat) tayyorlab beraman.\n\n"
    "📋 <b>Ishni boshlaymiz:</b>\n"
    "Qarzdorlar ro'yxati yozilgan <b>Excel faylni (.xlsx) yuboring.</b>\n\n"
    "Namuna kerakmi — /namuna buyrug'ini bosing.\n"
    "Barcha imkoniyatlar — /yordam"
)

YORDAM = (
    "ℹ️ <b>Buyruqlar</b>\n\n"
    "/start — boshlash\n"
    "/namuna — to'g'ri ustunlar bilan namuna Excel fayl\n"
    "/shablon — ariza shablonini tanlash\n"
    "/tekshirish — oxirgi ro'yxatni hujjatsiz qayta ko'rib chiqish\n"
    "/oxirgi — oxirgi natijani qayta yuborish\n"
    "/sozlamalar — joriy sozlamalarni ko'rish\n"
    "/sud 100000 — sud (undiruv) xarajatini o'zgartirish\n"
    "/pochta 22000 — pochta xarajatini o'zgartirish\n"
    "/id — Telegram ID raqamingiz\n"
    "/yordam — shu yordam\n\n"
    "📎 <b>Fayl yuborish</b>\n"
    "• <b>.xlsx</b> — qarzdorlar ro'yxati → arizalar tayyorlanadi\n"
    "• <b>.docx</b> — yangi shablon → papkaga qo'shiladi\n\n"
    "📦 <b>Bitta yuborishda hammasi keladi</b>\n"
    "• arizalar — .zip arxiv (bitta bo'lsa — .pdf)\n"
    "• <b>Reyestr</b> — barcha arizalar va summalar jadvali (.xlsx)\n"
    "• hisobot — qarzdorlar soni, jami summalar, ogohlantirishlar\n"
    "Faylni ikki marta yuborish shart emas.\n\n"
    "🔍 <b>Bot har bir ro'yxatni o'zi tekshiradi</b>\n"
    "• takror qarzdorlar — bir xil ЖШШИР yoki ID заказа\n"
    "• to'ldirilmagan kataklar (ID, sana, summa)\n"
    "• «жами карздорликк» ustuni hisob-kitobga mos kelishi\n\n"
    "🧮 <b>Bot o'zi hisoblaydi</b>\n"
    "• tug'ilgan sana — JSHSHIR (PINFL) dan\n"
    "• jami qarzdorlik = qoldiq + sud xarajati\n"
    "• undirish summasi = qoldiq + sud xarajati + pochta xarajati\n\n"
    "📑 <b>Excel ustunlari</b> (tartibi va katta-kichik harf muhim emas)\n"
    "ID заказа, ФИО, Остаток долга, Сумма заказа, дата заказа, "
    "дата истеченя срок, ПИНФЛ, Паспорт серия, Телефон, Адресс, "
    "суд харажати\n"
    "Majburiy: <b>ФИО</b> va <b>Остаток долга</b>.\n\n"
    "✉️ <b>Pochta xarajati</b> Exceldan olinmaydi — har bir arizada bir xil summa.\n"
    "Joriy summa — /sozlamalar, o'zgartirish — /pochta 22000"
)


MENU_NAMUNA = "📋 Namuna Excel"
MENU_SHABLON = "📄 Shablon"
MENU_TEKSHIR = "🔍 Tekshirish"
MENU_OXIRGI = "📦 Oxirgi natija"
MENU_SOZLAMA = "⚙️ Sozlamalar"
MENU_YORDAM = "ℹ️ Yordam"


def main_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row(MENU_NAMUNA, MENU_SHABLON)
    kb.row(MENU_TEKSHIR, MENU_OXIRGI)
    kb.row(MENU_SOZLAMA, MENU_YORDAM)
    return kb


# ── Buyruqlar ──────────────────────────────────────────────────────

@bot.message_handler(commands=["start"])
def cmd_start(m):
    if not allowed(m):
        return deny(m)
    bot.send_message(m.chat.id, SALOM, reply_markup=main_menu())


@bot.message_handler(commands=["yordam", "help"])
def cmd_help(m):
    if not allowed(m):
        return deny(m)
    bot.send_message(m.chat.id, YORDAM)


@bot.message_handler(commands=["id"])
def cmd_id(m):
    bot.reply_to(m, f"Sizning Telegram ID: <code>{m.from_user.id}</code>")


@bot.message_handler(commands=["namuna"])
def cmd_namuna(m):
    if not allowed(m):
        return deny(m)
    path = build_sample_excel()
    try:
        with open(path, "rb") as f:
            bot.send_document(m.chat.id, f, visible_file_name="namuna.xlsx",
                              caption="📋 Ustunlarni shu tartibda to'ldiring "
                                      "va faylni menga yuboring.")
    finally:
        shutil.rmtree(os.path.dirname(path), ignore_errors=True)


@bot.message_handler(commands=["shablon"])
def cmd_shablon(m):
    if not allowed(m):
        return deny(m)
    tpls = templates()
    if not tpls:
        return bot.send_message(m.chat.id, "❌ Papkada birorta .docx shablon yo'q. "
                                           "Shablonni menga fayl qilib yuboring.")
    kb = types.InlineKeyboardMarkup()
    cur = current_template(m.chat.id)
    for i, t in enumerate(tpls):
        name = os.path.basename(t)
        try:
            same = bool(cur) and os.path.samefile(t, cur)
        except OSError:
            same = False
        kb.add(types.InlineKeyboardButton(("✅ " if same else "") + name,
                                          callback_data=f"tpl:{i}"))
    bot.send_message(m.chat.id, "📄 <b>Ariza shablonini tanlang:</b>", reply_markup=kb)


@bot.callback_query_handler(func=lambda c: c.data.startswith("tpl:"))
def cb_shablon(c):
    tpls = templates()
    i = int(c.data.split(":")[1])
    if i >= len(tpls):
        return bot.answer_callback_query(c.id, "Shablon topilmadi")
    name = os.path.basename(tpls[i])
    set_cfg(c.message.chat.id, "shablon", name)
    kind = template_kind_safe(tpls[i])
    bot.answer_callback_query(c.id, "Tanlandi")
    bot.edit_message_text(f"✅ Shablon tanlandi:\n<b>{esc(name)}</b>\n"
                          f"Turi: {kind_label(kind)}",
                          c.message.chat.id, c.message.message_id)


@bot.message_handler(commands=["sozlamalar"])
def cmd_sozlamalar(m):
    if not allowed(m):
        return deny(m)
    t = current_template(m.chat.id)
    tname = esc(os.path.basename(t)) if t else "— yo'q —"
    kind = kind_label(template_kind_safe(t)) if t else "—"
    bot.send_message(m.chat.id,
                     "⚙️ <b>Joriy sozlamalar</b>\n\n"
                     f"📄 Shablon: <b>{tname}</b>\n"
                     f"   Turi: {kind}\n"
                     f"⚖️ Sud xarajati: <b>{gs.money(cfg(m.chat.id, 'sud', DEFAULT_SUD))}</b> so'm\n"
                     f"✉️ Pochta xarajati: <b>{gs.money(cfg(m.chat.id, 'pochta', DEFAULT_POCHTA))}</b> so'm\n\n"
                     "O'zgartirish: /sud 100000  yoki  /pochta 22000")


@bot.message_handler(commands=["sud", "pochta"])
def cmd_xarajat(m):
    if not allowed(m):
        return deny(m)
    key = "sud" if m.text.split()[0].lstrip("/").startswith("sud") else "pochta"
    parts = m.text.split(maxsplit=1)
    if len(parts) < 2:
        return bot.reply_to(m, f"Summani yozing. Masalan: <code>/{key} 100000</code>")
    val = gs.to_num(parts[1])
    if val is None or val < 0:
        return bot.reply_to(m, "❌ Summa noto'g'ri. Masalan: <code>/%s 100000</code>" % key)
    set_cfg(m.chat.id, key, val)
    nom = "Sud xarajati" if key == "sud" else "Pochta xarajati"
    bot.reply_to(m, f"✅ {nom}: <b>{gs.money(val)}</b> so'm")


@bot.message_handler(commands=["tekshirish", "tekshir"])
def cmd_tekshirish(m):
    """Ro'yxat baribir avtomatik tekshiriladi. Bu tugma — oxirgi ro'yxatni
    hujjat tayyorlamasdan qayta ko'rish uchun."""
    if not allowed(m):
        return deny(m)
    last = cfg(m.chat.id, "oxirgi_royxat") or {}
    if not last.get("id"):
        return bot.send_message(m.chat.id,
                                "🔍 <b>Tekshirish</b>\n\n"
                                "Har bir ro'yxat avtomatik tekshiriladi — arizalar bilan "
                                "birga ogohlantirishlar ham keladi, alohida buyruq shart emas.\n\n"
                                "Bu tugma <b>oxirgi yuborilgan ro'yxatni</b> hujjat "
                                "tayyorlamasdan qayta tekshiradi. Hozircha ro'yxat yo'q — "
                                "qarzdorlar ro'yxatini (.xlsx) yuboring.")
    data = download_file(m.chat.id, last["id"], last.get("hajm") or 0)
    if data is None:
        return
    check_list(m.chat.id, data, last.get("nom") or "ro'yxat.xlsx")


@bot.message_handler(commands=["oxirgi", "natija"])
def cmd_oxirgi(m):
    if not allowed(m):
        return deny(m)
    last = cfg(m.chat.id, "oxirgi_natija") or {}
    files = last.get("fayllar") or []
    if not files:
        return bot.send_message(m.chat.id,
                                "📦 Hozircha natija yo'q.\n"
                                "Qarzdorlar ro'yxatini (.xlsx) yuboring — tayyor arizalar "
                                "shu yerda saqlanadi va istagan vaqtda qayta olasiz.")
    bot.send_message(m.chat.id, f"📦 <b>Oxirgi natija</b> — {esc(last.get('vaqt', '—'))}")
    yoq = []
    for it in files:
        try:
            bot.send_document(m.chat.id, it["id"], timeout=SEND_TIMEOUT,
                              visible_file_name=it.get("nom"), caption=it.get("izoh"))
        except Exception as e:
            # file_id Telegram serverida saqlanadi, lekin abadiy emas
            log.warning("Oxirgi natijani yuborib bo'lmadi: %s", e)
            yoq.append(esc(it.get("nom") or "fayl"))
    if last.get("hisobot"):
        bot.send_message(m.chat.id, last["hisobot"])
    if yoq:
        bot.send_message(m.chat.id,
                         "⚠️ Bu fayllar endi mavjud emas: " + ", ".join(yoq) +
                         "\nRo'yxatni qayta yuboring — arizalar yangidan tayyorlanadi.")


def remember_file(items, msg, nom, izoh):
    """Yuborilgan faylning file_id si — qayta yuborish uchun (diskda saqlanmaydi)."""
    try:
        items.append({"id": msg.document.file_id, "nom": nom, "izoh": izoh})
    except Exception:
        pass


def remember_result(chat_id, items, hisobot):
    if not items:
        return
    set_cfg(chat_id, "oxirgi_natija", {
        "vaqt": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
        "fayllar": items[:20],
        "hisobot": hisobot[:3000],
    })


@bot.message_handler(func=lambda m: m.text in (MENU_NAMUNA, MENU_SHABLON, MENU_TEKSHIR,
                                               MENU_OXIRGI, MENU_SOZLAMA, MENU_YORDAM))
def menu_buttons(m):
    if not allowed(m):
        return deny(m)
    return {MENU_NAMUNA: cmd_namuna,
            MENU_SHABLON: cmd_shablon,
            MENU_TEKSHIR: cmd_tekshirish,
            MENU_OXIRGI: cmd_oxirgi,
            MENU_SOZLAMA: cmd_sozlamalar}.get(m.text, cmd_help)(m)


# ── Namuna Excel ───────────────────────────────────────────────────

SAMPLE_HEADERS = ["ID заказа", "ФИО", "Остаток долга", "Сумма заказа", "дата заказа",
                  "дата истеченя срок", "ПИНФЛ", "Паспорт серия", "Телефон",
                  "Адресс", "суд харажати"]
SAMPLE_ROW = ["809821", "NASIMOV RUSTAMJON BAXRIDDIN O‘G‘LI", 731581.28, 1848000,
              "14.10.2024", "14.10.2025", "30109995780023", "AB4886665",
              "998934312777", "Termiz sh., Bog'zor MFY, Sa'diy ko'chasi, 1/1",
              100000]


def build_sample_excel():
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист1"
    ws.append(SAMPLE_HEADERS)
    ws.append(SAMPLE_ROW)
    fill = PatternFill("solid", fgColor="DDEBF7")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = fill
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    widths = [12, 38, 15, 14, 13, 16, 18, 15, 15, 40, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i) if i <= 26 else "A"].width = w
    ws.freeze_panes = "A2"
    path = os.path.join(tempfile.mkdtemp(), "namuna.xlsx")
    wb.save(path)
    return path


# ── Reyestr — barcha arizalar va summalar jadvali ──────────────────

# (ustun nomi, kengligi, kalit).  "#" — raqamli ustun, jamlanadi.
REYESTR_COLS = [
    ("№",                6,  ""),
    ("ID заказа",        13, "id"),
    ("ФИО",              36, "fio"),
    ("ЖШШИР",            17, "pinfl"),
    ("Туғилган сана",    14, "tugilgan"),
    ("Паспорт",          12, "passport"),
    ("Телефон",          15, "tel"),
    ("Манзил",           34, "adres"),
    ("Шартнома санаси",  16, "sana"),
    ("Муддат тугаши",    15, "srok"),
    ("Шартнома суммаси", 17, "#summa"),
    ("Асосий қарз",      15, "#ostatok"),
    ("Суд харажати",     14, "#sud"),
    ("Почта харажати",   15, "#pochta"),
    ("Жами қарздорлик",  17, "#jami"),
    ("Ундириш суммаси",  17, "#jami2"),
    ("Ариза файли",      44, "_fayl"),
]


def build_reyestr(rows, path, ogohlar=()):
    """Jadval: har bir qarzdor — bitta qator, oxirida JAMI.
    Summalar matn emas, haqiqiy son — Excelda darrov jamlash mumkin."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reyestr"
    ws.append([c[0] for c in REYESTR_COLS])

    jamilar = {}
    for i, d in enumerate(rows, 1):
        line = [i]
        for _, _, key in REYESTR_COLS[1:]:
            if key.startswith("#"):
                n = gs.to_num(d.get(key[1:])) or 0.0
                jamilar[key] = jamilar.get(key, 0.0) + n
                line.append(round(n, 2))
            else:
                line.append(str(d.get(key) or ""))
        ws.append(line)
    oxirgi = ws.max_row                       # so'nggi qarzdor qatori

    # JAMI qatori: hisoblangan son yoziladi — formula qayta hisoblanishini kutmaydi
    total = ["JAMI"]
    for _, _, key in REYESTR_COLS[1:]:
        total.append(round(jamilar.get(key, 0.0), 2) if key.startswith("#") else "")
    ws.append(total)
    jami_qator = ws.max_row

    head_fill = PatternFill("solid", fgColor="DDEBF7")
    jami_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="B8C4D0")
    ramka = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    for idx, (_, w, key) in enumerate(REYESTR_COLS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
        raqam = key.startswith("#")
        for r in range(2, jami_qator + 1):
            c = ws.cell(r, idx)
            c.border = ramka
            if raqam:
                c.number_format = "#,##0.00"
                c.alignment = Alignment(horizontal="right")
            elif idx == 1:
                c.alignment = Alignment(horizontal="center")
            else:
                c.alignment = Alignment(vertical="top", wrap_text=True)

    for c in ws[jami_qator]:
        c.font = Font(bold=True)
        c.fill = jami_fill

    ws.freeze_panes = "C2"
    if oxirgi >= 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(REYESTR_COLS))}{oxirgi}"

    if ogohlar:
        ws2 = wb.create_sheet("Ogohlantirishlar")
        ws2.append(["№", "Ogohlantirish"])
        for i, t in enumerate(ogohlar, 1):
            ws2.append([i, uz(t)])
        for c in ws2[1]:
            c.font = Font(bold=True)
            c.fill = head_fill
        ws2.column_dimensions["A"].width = 6
        ws2.column_dimensions["B"].width = 110
        for r in range(2, ws2.max_row + 1):
            ws2.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(path)
    return path


def takrorlar(rows):
    """Ro'yxatga ikki marta tushgan qarzdor: ЖШШИР yoki ID заказа takrorlansa."""
    ogoh = []
    pinfl_seen, id_seen = {}, {}
    for d in rows:
        n = d.get("_qator") or "?"
        fio = d.get("fio") or ""
        pinfl = (d.get("pinfl") or "").strip()
        zid = (d.get("id") or "").strip()
        if pinfl and pinfl in pinfl_seen:
            ogoh.append(f"{n}-qator, {fio}: ЖШШИР {pinfl_seen[pinfl]}-qator bilan "
                        "bir xil — takror qarzdor bo'lishi mumkin")
        elif zid and zid in id_seen:
            ogoh.append(f"{n}-qator, {fio}: ID заказа {id_seen[zid]}-qator bilan "
                        "bir xil — takror ariza bo'lishi mumkin")
        if pinfl:
            pinfl_seen.setdefault(pinfl, n)
        if zid:
            id_seen.setdefault(zid, n)
    return ogoh


def xarajat_farqi(rows, chat_id):
    """Excel ustunidagi sud xarajati sozlamadan farq qilsa — bitta umumiy ogohlantirish.
    Pochta tekshirilmaydi: u Exceldan olinmaydi, har doim sozlamadagi summa."""
    ogoh = []
    for key, nom, default in (("sud", "суд харажати", DEFAULT_SUD),):
        kerak = cfg(chat_id, key, default)
        farq = {}
        for d in rows:
            n = gs.to_num(d.get(key))
            if n is not None and abs(n - kerak) > 0.01:
                farq[n] = farq.get(n, 0) + 1
        if farq:
            qiymat = ", ".join(f"{gs.money(v)} ({c} ta qatorda)" for v, c in sorted(farq.items()))
            ogoh.append(f"«{nom}»: Excelda {qiymat}, sozlamada {gs.money(kerak)} — "
                        "hujjatga Exceldagi qiymat olindi. Excelni to'g'rilang "
                        "yoki ustunni o'chiring")
    return ogoh


def jami_summalar(rows):
    """(jami qarzdorlik, pochta bilan undirish summasi)."""
    qarz = sum(gs.to_num(d.get("jami")) or 0 for d in rows)
    undiruv = sum(gs.to_num(d.get("jami2")) or 0 for d in rows)
    return qarz, undiruv


# ── Jarayon ko'rsatkichi ───────────────────────────────────────────

# Telegram bir xabarni tez-tez tahrirlashga yo'l qo'ymaydi — 3 soniyada bir marta
PROGRESS_EVERY = 3.0


def progress_bar(done, total, width=14):
    """'████████░░░░░░  57%'"""
    ulush = (done / total) if total else 0.0
    ulush = 0.0 if ulush < 0 else (1.0 if ulush > 1 else ulush)
    to_la = int(round(ulush * width))
    return "█" * to_la + "░" * (width - to_la) + f"  {int(round(ulush * 100))}%"


def qolgan_vaqt(boshlandi, done, total):
    """Taxminiy qolgan vaqt. Bir nechta hujjatdan keyingina ma'noli bo'ladi."""
    o_tgan = time.time() - boshlandi
    if done < 3 or done >= total or o_tgan < 1:
        return ""
    qoldi = o_tgan / done * (total - done)
    if qoldi < 60:
        return f"⏳ Taxminan {int(qoldi) + 1} soniya qoldi"
    return f"⏳ Taxminan {int(qoldi // 60) + 1} daqiqa qoldi"


class Progress:
    """Bitta status xabarini yangilab turadi: progress bar + hisoblagich."""

    def __init__(self, status, sarlavha, birlik="ta ariza"):
        self.status = status
        self.sarlavha = sarlavha
        self.birlik = birlik
        self.boshlandi = time.time()
        self.oxirgi = 0.0

    def show(self, done, total, izoh="", majburiy=False):
        if not majburiy and time.time() - self.oxirgi < PROGRESS_EVERY:
            return
        satr = [self.sarlavha,
                f"<code>{progress_bar(done, total)}</code>",
                f"📄 <b>{done}</b> / {total} {self.birlik}"]
        if izoh:
            satr.append("👤 " + esc(izoh))
        vaqt = qolgan_vaqt(self.boshlandi, done, total)
        if vaqt:
            satr.append(vaqt)
        edit(self.status, "\n".join(satr))
        self.oxirgi = time.time()


# ── Fayllarni qabul qilish ─────────────────────────────────────────

@bot.message_handler(content_types=["document"])
def on_document(m):
    if not allowed(m):
        return deny(m)
    name = (m.document.file_name or "").strip()
    low = name.lower()

    if low.endswith(".docx"):
        return save_template(m, name)
    if low.endswith((".xlsx", ".xlsm")):
        # ro'yxat eslab qolinadi — "🔍 Tekshirish" tugmasi shuni qayta o'qiydi
        set_cfg(m.chat.id, "oxirgi_royxat", {"id": m.document.file_id, "nom": name,
                                             "hajm": m.document.file_size or 0})
        return process_list(m, name)

    bot.reply_to(m, "❌ Faqat <b>.xlsx</b> (qarzdorlar ro'yxati) yoki "
                    "<b>.docx</b> (ariza shabloni) fayl yuboring.")


def download_file(chat_id, file_id, size=0):
    """Байты файла или None, если скачать не удалось (ответ уже отправлен)."""
    if size > MAX_DOWNLOAD_MB * 1024 * 1024:
        bot.send_message(chat_id, f"❌ Fayl juda katta ({size / 1048576:.1f} MB). "
                                  f"Telegram bot orqali {MAX_DOWNLOAD_MB} MB gacha "
                                  "olish mumkin.\nRo'yxatni bir necha faylga bo'lib yuboring.")
        return None
    try:
        info = bot.get_file(file_id)
        return bot.download_file(info.file_path)
    except Exception as e:
        log.warning("Faylni yuklab bo'lmadi: %s", e)
        bot.send_message(chat_id, f"❌ Faylni yuklab bo'lmadi: {esc(e)}")
        return None


def download(m):
    return download_file(m.chat.id, m.document.file_id,
                         getattr(m.document, "file_size", None) or 0)


def save_template(m, name):
    data = download(m)
    if data is None:
        return
    fname = gs.safe_filename(name) or "shablon.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"
    path = os.path.join(BASE_DIR, fname)
    with open(path, "wb") as f:
        f.write(data)

    try:
        text = gs.doc_text(gs.Document(path))
    except Exception as e:
        os.remove(path)
        return bot.reply_to(m, f"❌ Faylni Word hujjati sifatida o'qib bo'lmadi: {esc(e)}")

    # Без этой проверки готовая ариза становилась «шаблоном» и всё ломалось молча
    if not has_markers(text):
        os.remove(path)
        return bot.reply_to(m, "❌ Bu faylda ariza belgilari topilmadi.\n"
                               "Shablonda <code>(фиш)</code>, <code>(заказ ид)</code> yoki "
                               "<code>ЗАКАЗ ИД</code> kabi belgilar bo'lishi kerak.\n"
                               "Tayyor ariza o'rniga shablon yuborganingizni tekshiring.")

    set_cfg(m.chat.id, "shablon", os.path.basename(path))
    bot.reply_to(m, f"✅ Shablon saqlandi va tanlandi:\n"
                    f"<b>{esc(os.path.basename(path))}</b>\n"
                    f"Turi: {kind_label(gs.detect_template(text))}\n\n"
                    "Endi qarzdorlar ro'yxatini (.xlsx) yuboring.")


def read_list(status, xls_path, chat_id):
    """Excelni o'qish. Xato bo'lsa — tushunarli xabar yuborib, (None, None)."""
    try:
        rows, warns = gs.read_rows(xls_path,
                                   sud_default=cfg(chat_id, "sud", DEFAULT_SUD),
                                   pochta_default=cfg(chat_id, "pochta", DEFAULT_POCHTA))
    except (gs.DataError, SystemExit) as e:
        edit(status, "❌ <b>Excel xatosi</b>\n\n" + esc(uz(e)) +
                     "\n\nTo'g'ri ustunlar namunasi — /namuna")
        return None, None
    except Exception as e:
        log.exception("Excelni o'qishda xato")
        edit(status, "❌ <b>Faylni Excel sifatida o'qib bo'lmadi</b>\n\n" + esc(e) +
                     "\n\nFayl haqiqatan <b>.xlsx</b> ekanini tekshiring. Namuna — /namuna")
        return None, None

    if not rows:
        edit(status, "❌ Excelda ma'lumot topilmadi. "
                     "Sarlavha qatori va kamida bitta qarzdor bo'lishi kerak.")
        return None, None
    if len(rows) > MAX_ROWS:
        edit(status, f"❌ Juda ko'p qator ({len(rows)}). "
                     f"Chegara — {MAX_ROWS} ta. Ro'yxatni bo'lib yuboring.")
        return None, None
    return rows, warns


def check_list(chat_id, data, name):
    """Faqat tekshirish — hujjat tayyorlanmaydi. Arizalar ro'yxat yuborilganda
    allaqachon tayyorlangan, bu yerda uni qaytadan ko'rib chiqamiz."""
    status = bot.send_message(chat_id, "🔍 Ro'yxat tekshirilmoqda...")
    work = tempfile.mkdtemp(prefix="sud_tekshir_")
    try:
        xls = os.path.join(work, "royxat.xlsx")
        with open(xls, "wb") as f:
            f.write(data)

        rows, warns = read_list(status, xls, chat_id)
        if rows is None:
            return

        problems = xarajat_farqi(rows, chat_id) + list(warns) + takrorlar(rows)
        jami_qarz, jami_undiruv = jami_summalar(rows)
        tpl = current_template(chat_id)

        rep = ["🔍 <b>Tekshiruv natijasi</b>",
               f"📎 Fayl: {esc(name)}",
               f"👥 Qarzdorlar: <b>{len(rows)}</b> ta",
               f"💰 Jami qarzdorlik: <b>{gs.money(jami_qarz)}</b> so'm",
               f"⚖️ Jami undirish (pochta bilan): <b>{gs.money(jami_undiruv)}</b> so'm",
               f"✉️ Pochta xarajati: <b>{gs.money(cfg(chat_id, 'pochta', DEFAULT_POCHTA))}</b> so'm (har bir arizada)",
               f"📄 Shablon: {esc(os.path.basename(tpl)) if tpl else '— yo‘q —'}"]
        if problems:
            rep.append("\n⚠️ <b>Diqqat, tekshiring:</b>")
            rep += ["• " + esc(uz(p)) for p in problems[:15]]
            if len(problems) > 15:
                rep.append(f"• ... va yana {len(problems) - 15} ta ogohlantirish")
        else:
            rep.append("\n✅ Xatolik topilmadi.")
        rep.append("\nTayyor arizalarni qayta olish — 📦 <b>Oxirgi natija</b>.\n"
                   "Excelni to'g'rilagan bo'lsangiz, yangilangan faylni yuboring.")
        edit(status, "\n".join(rep))

        # Reyestr: tekshirishda ham summalarni jadvalda ko'rish qulay
        stamp = datetime.datetime.now().strftime("%d.%m.%Y")
        rname = f"Reyestr {stamp}.xlsx"
        try:
            build_reyestr(rows, os.path.join(work, rname), problems)
            with open(os.path.join(work, rname), "rb") as f:
                bot.send_document(chat_id, f, visible_file_name=rname, timeout=SEND_TIMEOUT,
                                  caption="📊 Reyestr — tekshirilgan ro'yxat jadvali")
        except Exception as e:
            log.exception("Tekshirish reyestri")
            bot.send_message(chat_id, f"⚠️ Reyestr tayyorlanmadi: {esc(e)}")

        log.info("chat=%s tekshirish qator=%s ogohlantirish=%s",
                 chat_id, len(rows), len(problems))
    except Exception as e:
        log.exception("check_list")
        edit(status, f"❌ Kutilmagan xato: {esc(e)}")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def process_list(m, name):
    tpl = current_template(m.chat.id)
    if not tpl:
        return bot.reply_to(m, "❌ Shablon topilmadi. Avval ariza shablonini (.docx) yuboring.")

    data = download(m)
    if data is None:
        return

    status = bot.reply_to(m, "⏳ Ro'yxat qabul qilindi, o'qiyapman...")
    work = tempfile.mkdtemp(prefix="sud_")
    try:
        xls = os.path.join(work, "royxat.xlsx")
        with open(xls, "wb") as f:
            f.write(data)

        rows, warns = read_list(status, xls, m.chat.id)
        if rows is None:
            return

        try:
            tpl_text = gs.doc_text(gs.Document(tpl))
        except Exception as e:
            log.exception("Shablonni o'qishda xato")
            return edit(status, "❌ Shablonni o'qib bo'lmadi: " + esc(e) +
                                "\nBoshqa shablon tanlang — /shablon")
        kind = gs.detect_template(tpl_text)
        tpl_fio = gs.detect_template_fio(tpl_text) if kind == "new" else ""

        out_dir = os.path.join(work, "arizalar")
        os.makedirs(out_dir, exist_ok=True)

        files, problems = [], xarajat_farqi(rows, m.chat.id) + list(warns) + takrorlar(rows)
        jami_qarz = jami_undiruv = 0.0

        # PDF: Word bir marta ishga tushadi va butun ro'yxat uchun ishlaydi
        conv = None
        if gs.OUTPUT_FORMAT == "pdf":
            edit(status, "📄 PDF uchun Microsoft Word ishga tushirilmoqda...")
            try:
                conv = gs.WordPdf().open()
            except gs.PdfError as e:
                # arizalar yo'qolmasin — Word ko'rinishida yuboramiz va ogohlantiramiz
                log.warning("PDF qilib bo'lmadi: %s", e)
                problems.insert(0, f"PDF qilib bo'lmadi ({e}) — arizalar Word (.docx) "
                                   "ko'rinishida yuborildi")
        tmp_docx = os.path.join(work, "ariza.docx")

        # ro'yxat o'qilishi bilan tekshiriladi — natijasi darrov ko'rinib turadi
        sarlavha = "⚙️ <b>Arizalar tayyorlanmoqda</b>" + (" (PDF)" if conv else "")
        if problems:
            sarlavha += f"\n🔍 Tekshiruv: {len(problems)} ta ogohlantirish (hisobotda)"
        pr = Progress(status, sarlavha)
        pr.show(0, len(rows), majburiy=True)

        try:
            for i, d in enumerate(rows, 1):
                sana = d["sana"] if d["sana"] != "__.__.____" else \
                    datetime.date.today().strftime("%d.%m.%Y")
                fname = gs.safe_filename(gs.FILE_NAME_TEMPLATE.format(
                    fio=d["fio"], sana=sana, id=d["id"] or "b-n"))
                if not fname.lower().endswith(".docx"):
                    fname += ".docx"
                nom = os.path.splitext(fname)[0]
                try:
                    if conv:
                        left = gs.make_document(tpl, d, tmp_docx, kind, tpl_fio)
                        path = unique_path(out_dir, nom + ".pdf")
                        try:
                            conv.convert(tmp_docx, path)
                        except Exception as e:
                            # PDF chiqmadi — hujjat yo'qolmasin, Word faylini yuboramiz
                            log.exception("PDF xatosi: %s", d["fio"])
                            path = unique_path(out_dir, nom + ".docx")
                            shutil.copyfile(tmp_docx, path)
                            problems.append(f"{d['fio']}: PDF qilinmadi, Word fayl "
                                            f"yuborildi — {e}")
                    else:
                        path = unique_path(out_dir, fname)
                        left = gs.make_document(tpl, d, path, kind, tpl_fio)
                    files.append(path)
                    d["_fayl"] = os.path.basename(path)      # reyestr uchun
                    if left:
                        problems.append(f"{d['fio']}: almashtirilmadi — {', '.join(left)}")
                    jami_qarz += gs.to_num(d["jami"]) or 0
                    jami_undiruv += gs.to_num(d["jami2"]) or 0
                except Exception as e:
                    problems.append(f"{d['fio']}: XATO — {e}")
                    d["_fayl"] = "XATO — ariza tayyorlanmadi"
                    log.exception("Hujjat yaratishda xato: %s", d["fio"])

                pr.show(i, len(rows), d["fio"])
        finally:
            if conv:
                conv.close()

        if not files:
            return edit(status, "❌ Birorta hujjat yaratilmadi.\n" +
                        "\n".join(esc(uz(p)) for p in problems[:10]))

        pr.show(len(rows), len(rows), majburiy=True)
        sent, send_errors, items = send_results(m.chat.id, files, work, status)
        problems += send_errors

        # Reyestr — barcha arizalar va summalar jadvali
        stamp = datetime.datetime.now().strftime("%d.%m.%Y")
        rname = f"Reyestr {stamp}.xlsx"
        rizoh = "📊 Reyestr — barcha arizalar va summalar jadvali"
        try:
            edit(status, "📊 <b>Reyestr tayyorlanmoqda...</b>")
            build_reyestr(rows, os.path.join(work, rname), problems)
            with open(os.path.join(work, rname), "rb") as f:
                rmsg = bot.send_document(m.chat.id, f, visible_file_name=rname,
                                         timeout=SEND_TIMEOUT, caption=rizoh)
            remember_file(items, rmsg, rname, rizoh)
        except Exception as e:
            log.exception("Reyestr tayyorlashda xato")
            problems.append(f"Reyestr yuborilmadi — {e}")

        # yakuniy hisobot
        rep = ["✅ <b>Tayyor</b>" if not send_errors else "⚠️ <b>Qisman tayyor</b>",
               f"👥 Qarzdorlar: <b>{len(files)}</b> ta",
               f"📤 Yuborildi: <b>{sent}</b> ta",
               f"💰 Jami qarzdorlik: <b>{gs.money(jami_qarz)}</b> so'm",
               f"⚖️ Jami undirish (pochta bilan): <b>{gs.money(jami_undiruv)}</b> so'm",
               f"✉️ Pochta xarajati: <b>{gs.money(cfg(m.chat.id, 'pochta', DEFAULT_POCHTA))}</b> so'm (har bir arizada)",
               f"📄 Shablon: {esc(os.path.basename(tpl))}"]
        if problems:
            rep.append("\n⚠️ <b>Diqqat, tekshiring:</b>")
            rep += ["• " + esc(uz(p)) for p in problems[:15]]
            if len(problems) > 15:
                rep.append(f"• ... va yana {len(problems) - 15} ta ogohlantirish")
        edit(status, "\n".join(rep))
        remember_result(m.chat.id, items, "\n".join(rep))
        log.info("chat=%s hujjat=%s yuborildi=%s shablon=%s",
                 m.chat.id, len(files), sent, os.path.basename(tpl))

    except Exception as e:
        log.exception("process_list")
        edit(status, f"❌ Kutilmagan xato: {esc(e)}")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def unique_path(folder, fname):
    path = os.path.join(folder, fname)
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(fname)
    n = 2
    while os.path.exists(os.path.join(folder, f"{stem} ({n}){ext}")):
        n += 1
    return os.path.join(folder, f"{stem} ({n}){ext}")


def send_results(chat_id, files, work, status=None):
    """Возвращает (сколько отправлено, список ошибок, список отправленных файлов).
    Сбой одной отправки не должен ронять всю обработку и терять документы."""
    errors, items = [], []

    def holat(text):
        if status is not None:
            edit(status, text)

    if len(files) == 1:
        nom = os.path.basename(files[0])
        holat("📤 <b>Hujjat yuborilmoqda...</b>")
        try:
            with open(files[0], "rb") as f:
                msg = bot.send_document(chat_id, f, timeout=SEND_TIMEOUT,
                                        visible_file_name=nom)
            remember_file(items, msg, nom, None)
            return 1, errors, items
        except Exception as e:
            log.exception("Hujjatni yuborishda xato")
            return 0, [f"hujjat yuborilmadi — {e}"], items

    limit = MAX_ZIP_MB * 1024 * 1024
    chunks, cur, size = [], [], 0
    for p in files:
        sz = os.path.getsize(p)
        if cur and (len(cur) >= MAX_FILES_PER_ZIP or size + sz > limit):
            chunks.append(cur)
            cur, size = [], 0
        cur.append(p)
        size += sz
    if cur:
        chunks.append(cur)

    stamp = datetime.datetime.now().strftime("%d.%m.%Y")
    sent = 0
    for i, chunk in enumerate(chunks, 1):
        zname = f"Sud arizalari {stamp}.zip" if len(chunks) == 1 else \
                f"Sud arizalari {stamp} ({i}-qism).zip"
        zpath = os.path.join(work, zname)
        izoh = f"📦 {len(chunk)} ta ariza"
        if len(chunks) > 1:
            hisob = (f"<code>{progress_bar(i - 1, len(chunks))}</code>\n"
                     f"📦 <b>{i}</b> / {len(chunks)}-qism · {len(chunk)} ta ariza")
        else:
            hisob = f"📄 <b>{len(chunk)}</b> ta ariza"
        try:
            holat("📦 <b>Arxivga joylanmoqda...</b>\n" + hisob)
            with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
                for p in chunk:
                    z.write(p, os.path.basename(p))
            holat(f"📤 <b>Yuborilmoqda</b> — "
                  f"{os.path.getsize(zpath) / 1048576:.1f} MB\n" + hisob)
            with open(zpath, "rb") as f:
                msg = bot.send_document(chat_id, f, visible_file_name=zname,
                                        timeout=SEND_TIMEOUT, caption=izoh)
            remember_file(items, msg, zname, izoh)
            sent += len(chunk)
        except Exception as e:
            log.exception("Arxivni yuborishda xato: %s", zname)
            errors.append(f"{zname} yuborilmadi — {e}")
        finally:
            if os.path.exists(zpath):
                os.remove(zpath)
    return sent, errors, items


def edit(status_msg, text):
    # Telegram обрезает не сам — при длине > 4096 он возвращает ошибку,
    # и отчёт не доходил до пользователя вообще
    if len(text) > MAX_MESSAGE_LEN:
        text = text[:MAX_MESSAGE_LEN - 24] + "\n… (qisqartirildi)"
    try:
        bot.edit_message_text(text, status_msg.chat.id, status_msg.message_id)
        return
    except Exception as e:
        if "message is not modified" in str(e):
            return
        log.warning("edit_message_text: %s", e)
    try:
        bot.send_message(status_msg.chat.id, text)
    except Exception as e:
        log.error("Xabar yuborilmadi: %s | matn: %.300s", e, text)


@bot.message_handler(func=lambda m: True, content_types=["text"])
def fallback(m):
    if not allowed(m):
        return deny(m)
    bot.reply_to(m, "📋 Menga qarzdorlar ro'yxati yozilgan <b>Excel (.xlsx)</b> faylni yuboring.\n"
                    "Namuna — /namuna,  buyruqlar — /yordam")


# ── Ishga tushirish ────────────────────────────────────────────────

def main():
    if not TOKEN or ":" not in TOKEN:
        print("❌ Bot tokeni ko'rsatilmagan (TOKEN).")
        return
    try:
        me = bot.get_me()
    except Exception as e:
        print("❌ Telegramga ulanib bo'lmadi. Internet yoki tokenni tekshiring.")
        print("   ", e)
        return

    tpls = [os.path.basename(t) for t in templates()]
    print("=" * 60)
    print(f"  BOT ISHGA TUSHDI:  @{me.username}")
    print(f"  Shablonlar ({len(tpls)}): " + (", ".join(tpls) if tpls else "YO'Q!"))
    print("  Ruxsat: " + ("hamma uchun ochiq" if not RUXSAT_ETILGANLAR
                           else f"{len(RUXSAT_ETILGANLAR)} ta foydalanuvchi"))
    print("  To'xtatish: Ctrl+C yoki oynani yopish")
    print("=" * 60)

    try:
        bot.infinity_polling(timeout=30, long_polling_timeout=30, skip_pending=True)
    except KeyboardInterrupt:
        print("\nBot to'xtatildi.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
    if os.name == "nt":
        input("\nChiqish uchun Enter...")
