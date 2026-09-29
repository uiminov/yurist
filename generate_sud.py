# -*- coding: utf-8 -*-
"""
Автоматическое формирование судебных заявлений по списку должников.

Что делает:
  1. Читает Excel со списком должников (по НАЗВАНИЯМ колонок, порядок не важен).
  2. Определяет, какой шаблон Word перед ним, и подставляет данные клиента.
  3. Сохраняет отдельный .docx на каждого должника в папку "Готовые заявления".

Поддерживаются два шаблона:
  • "Ариза Олча палатага.docx"  — новый (метки в скобках: (фиш), (адрес), (заказ ид)...)
  • "Судга ариза.docx"          — старый (метки заглавными: ЗАКАЗ ИД, ОСТАТОК ДОЛГ...)

Запуск:  дважды кликнуть ЗАПУСК.bat   (или:  python generate_sud.py)
Требования: Python 3.9+, python-docx, openpyxl (ЗАПУСК.bat ставит их сам).
"""

import os
import re
import sys
import glob
import shutil
import tempfile
import datetime

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from docx import Document
    from docx.text.run import Run
    from docx.text.hyperlink import Hyperlink
    import openpyxl
except ImportError:
    print("Не установлены библиотеки. Выполните в командной строке:")
    print("    pip install python-docx openpyxl")
    input("\nНажмите Enter...")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════
#  НАСТРОЙКИ  — при необходимости меняйте только этот блок
# ═══════════════════════════════════════════════════════════════════

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Имена файлов. Оставьте "" — скрипт найдёт сам, если файл в папке один.
EXCEL_FILE = "Книга10.xlsx"
TEMPLATE_FILE = "Ариза Олча палатага.docx"

OUTPUT_DIR = os.path.join(BASE_DIR, "Готовые заявления")

# Суд/ундирув харажати, если в Excel нет колонки "суд харажати"
XIZMAT_HAQI = 100_000
# Почтовые расходы — одна сумма для всех заявлений
POCHTA_HAQI = 22_000
# False — колонка "почта харажати" в Excel игнорируется, всегда POCHTA_HAQI
#         (в старых списках там 20 600, и она тихо попадала в документы).
# True  — брать сумму из Excel, POCHTA_HAQI — только для пустых ячеек.
POCHTA_IZ_EXCEL = False

# Шаблон имени файла. Доступно: {fio} {sana} {id}
FILE_NAME_TEMPLATE = "{fio} {sana} суд карори.docx"

# Формат готовых заявлений:
#   "pdf"  — через установленный Microsoft Word (нужен pywin32), выглядит один в один с Word
#   "docx" — как раньше, файлы Word
OUTPUT_FORMAT = "pdf"

# Убирать жёлтую подсветку с подставленного текста
REMOVE_HIGHLIGHT = True

# Разделитель разрядов между тысячами
THOUSANDS_SEP = " "      # " " = неразрывный пробел, "" = без разделителя
# Копейки: "auto" — только если есть, "always" — всегда, "never" — округлять
MONEY_DECIMALS = "auto"
DECIMAL_SEP = "."

# ФИО-образец в шаблоне. "" — скрипт найдёт его сам по метке (фиш).
TEMPLATE_FIO = ""

# Названия колонок Excel: ключ -> варианты написания (регистр/пробелы не важны)
COLUMNS = {
    "id":         ["id заказа", "заказ id", "заказ ид", "id"],
    "fio":        ["фио", "ф.и.ш", "ф.и.о", "мижоз", "клиент"],
    "ostatok":    ["остаток долга", "остаток основного долга", "остаток", "қолдиқ"],
    "summa":      ["сумма заказа", "заказ суммаси", "сумма договора"],
    "sana":       ["дата заказа", "дата договора", "заказ санаси"],
    "srok":       ["дата истеченя срок", "дата истечения срока", "срок истечения",
                   "срок договора", "муддати тугаш санаси"],
    "pinfl":      ["пинфл", "жшшир", "жшшир (пинфл)", "пнфл"],
    "passport":   ["паспорт серия", "паспорт", "паспорт серияси", "п.сериа"],
    "tel":        ["телефон", "тел", "тел.", "телефон раками"],
    "adres":      ["адресс", "адрес", "манзил"],
    "sud":        ["суд харажати", "ундирув харажатлари", "суд харажатлари"],
    "pochta":     ["почта харажати", "почта харажатлари", "почта"],
    "jami_col":   ["жами карздорликк", "жами карздорлик", "жами қарздорлик", "итого долг"],
    # только для старого шаблона:
    "days":       ["просрочка (дней)", "просрочка дней", "просрочка кун", "кун"],
    "prosrochka": ["просрочка сумма", "сумма просрочки", "муддати ўтган"],
}

# ═══════════════════════════════════════════════════════════════════
#  ДАЛЬШЕ МЕНЯТЬ НЕ НУЖНО
# ═══════════════════════════════════════════════════════════════════

NBSP = " "


class DataError(Exception):
    """Понятная ошибка данных (нет колонки, нет файла и т.п.).
    Раньше бросался SystemExit — его не ловит `except Exception`."""


def ph(value, blank="____"):
    """Пустое значение -> видимый прочерк.
    Иначе из фразы молча пропадала сумма и документ выглядел исправным."""
    s = "" if value is None else str(value).strip()
    return s if s else blank


def to_num(value):
    """'1 234,50' / 1234.5 -> float. Не число -> None."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    s = str(value).strip().replace(NBSP, "").replace(" ", "")
    if not s:
        return None
    s = s.replace(",", ".")
    if s.count(".") > 1:                       # 1.234.567 -> разделитель разрядов
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def money(value):
    """731581.28 -> '731 581.28' ; 100000 -> '100 000'"""
    n = to_num(value)
    if n is None:
        return str(value or "").strip()
    n = round(n, 2)
    show_dec = MONEY_DECIMALS == "always" or (
        MONEY_DECIMALS == "auto" and abs(n - round(n)) > 1e-9)
    if show_dec:
        whole, dec = f"{n:,.2f}".split(".")
        return whole.replace(",", THOUSANDS_SEP) + DECIMAL_SEP + dec
    return f"{int(round(n)):,}".replace(",", THOUSANDS_SEP)


def clean_str(value):
    """Убирает хвост '.0' у чисел, пришедших из Excel."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def fmt_date(value, empty="__.__.____"):
    if value is None or str(value).strip() == "":
        return empty
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d.%m.%Y")
    s = str(value).strip()
    for f in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
              "%m/%d/%Y", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return datetime.datetime.strptime(s, f).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return s


def birth_from_pinfl(pinfl):
    """ЖШШИР 43110894320043 -> '31.10.1989' (цифры 2-7 = ДДММГГ, 1-я = век)."""
    d = re.sub(r"\D", "", clean_str(pinfl))
    if len(d) < 7:
        return "__.__.____"
    century = {"1": 1800, "2": 1800, "3": 1900, "4": 1900, "5": 2000, "6": 2000}
    base = century.get(d[0])
    if base is None:
        return "__.__.____"
    dd, mm, yy = d[1:3], d[3:5], d[5:7]
    try:
        datetime.date(base + int(yy), int(mm), int(dd))
    except ValueError:
        return "__.__.____"
    return f"{dd}.{mm}.{base + int(yy)}"


def safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', " ", str(name))
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name[:150]


# ── Замена текста в Word с сохранением форматирования ───────────────
# Word рвёт фразу на несколько "run"-ов, поэтому ищем по склеенному
# тексту абзаца и аккуратно раскладываем замену обратно по run-ам.

def para_runs(paragraph):
    """Все run-ы абзаца по порядку, включая те, что лежат внутри гиперссылок.
    paragraph.runs их не отдаёт — метка в ссылке оставалась незаменённой."""
    out = []
    for item in paragraph.iter_inner_content():
        if isinstance(item, Run):
            out.append(item)
        elif isinstance(item, Hyperlink):
            out.extend(item.runs)
    return out


def replace_in_paragraph(paragraph, pattern, repl, flags=0):
    runs = para_runs(paragraph)
    if not runs:
        return 0
    rx = re.compile(pattern, flags)
    count, search_from = 0, 0
    for _ in range(200):                       # защита от зацикливания
        text = "".join(r.text for r in runs)
        if search_from > len(text):
            break
        m = rx.search(text, search_from)
        if not m:
            break
        start, end = m.start(), m.end()
        if end == start:                       # пустое совпадение — сдвигаемся
            search_from = start + 1
            continue

        pos, s_run, s_off, e_run, e_off = 0, None, 0, None, 0
        for i, r in enumerate(runs):
            nxt = pos + len(r.text)
            if s_run is None and start < nxt:
                s_run, s_off = i, start - pos
            if e_run is None and end <= nxt:
                e_run, e_off = i, end - pos
                break
            pos = nxt
        if s_run is None or e_run is None:
            break

        if s_run == e_run:
            t = runs[s_run].text
            runs[s_run].text = t[:s_off] + repl + t[e_off:]
            touched = [s_run]
        else:
            runs[s_run].text = runs[s_run].text[:s_off] + repl
            for i in range(s_run + 1, e_run):
                runs[i].text = ""
            runs[e_run].text = runs[e_run].text[e_off:]
            touched = list(range(s_run, e_run + 1))

        if REMOVE_HIGHLIGHT:
            for i in touched:
                try:
                    runs[i].font.highlight_color = None
                except Exception:
                    pass
        count += 1
        # Дальше ищем ЗА подставленным текстом. Раньше поиск начинался с нуля,
        # и правило с lookahead находило свою же подстановку — 30 холостых кругов,
        # а второе совпадение в том же абзаце до замены не доходило.
        search_from = start + len(repl)
    return count


def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                for t2 in cell.tables:
                    for r2 in t2.rows:
                        for c2 in r2.cells:
                            for p in c2.paragraphs:
                                yield p
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            if part is None:
                continue
            for p in part.paragraphs:
                yield p
            for t in part.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p


def doc_text(doc):
    return "\n".join(p.text for p in iter_paragraphs(doc))


# ── Правила замены ─────────────────────────────────────────────────

I = re.IGNORECASE
NUM = r"\d[\d\s ]*"            # число (пробелы внутри допускаются)
NUMD = r"\d[\d\s .,]*"         # число с копейками


def rules_new(d, tpl_fio):
    """Шаблон "Ариза Олча палатага.docx" — метки в круглых скобках."""
    def g(key):
        return ph(d.get(key))

    r = [
        # 1 325 214( остаток долг + 100 000 + почта )   ← первым, он длиннее
        (rf"{NUMD}\(\s*остаток\s*долг\s*\+\s*100[\s\xa0]*000\s*\+\s*почта\s*\)\s*",
         g("jami2") + " ", I),
        # 1 305 214( остаток долг + 100 000 )
        (rf"{NUMD}\(\s*остаток\s*долг\s*\+\s*100[\s\xa0]*000\s*\)\s*",
         g("jami") + " ", I),
        # 1 155 214(Остаток основного долга)
        (rf"{NUMD}\(\s*Остаток\s*основного\s*долга\s*\)", g("ostatok"), I),
        # 12 936 000Сумма заказа   (без скобок!)
        (rf"{NUM}\s*Сумма\s*заказа", g("summa"), I),
        # 23.08.2024(дата заказа)
        (r"\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4}\s*\(\s*дата\s*заказа\s*\)", g("sana"), I),
        # 01.09.2025(Срок истечение договора)
        (r"\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4}\s*\(\s*Срок\s*истечени?е?\s*договора\s*\)",
         g("srok"), I),
        # 774376(заказ ид)
        (rf"{NUM}\(\s*заказ\s*ид\s*\)", g("id"), I),
        # ...дом 17 (адрес)   — вся строка до метки
        (r"[^\n]*\(\s*адрес\s*\)", g("adres"), I),
        # 31.10.1989 й.т.   — дата рождения (берётся из ПИНФЛ)
        (r"\d{1,2}\.\d{1,2}\.\d{4}(?=\s*й\.\s*т\.)", g("tugilgan"), I),
        # AC2197627(п.сериа)
        (r"[A-ZА-Яa-zа-я0-9]+\s*\(\s*п\.?\s*сери[аи]\s*\)", g("passport"), I),
        # 43110894320043 (пнфл)
        (r"\d+\s*\(\s*(?:пнфл|пинфл|жшшир)\s*\)", g("pinfl"), I),
        # 998905880214(тел)
        (r"[\d+][\d+()\-\s]*\(\s*тел\s*\)", g("tel"), I),
        # 100 000 сўм ундирув харажатлари
        (rf"{NUMD}(?=\s+сўм\s+ундирув\s+харажатлари)", g("sud"), I),
        # 20 600.00 сўм почта харажатини
        (rf"{NUMD}(?=\s+сўм\s+почта\s+харажатини)", g("pochta"), I),
    ]
    if tpl_fio:
        # ФИО-образец из шаблона, с меткой (фиш) или без неё
        pat = r"\s+".join(re.escape(w) for w in tpl_fio.split())
        r.append((pat + r"(?:\s*\(\s*фиш\s*\))?", g("fio"), I))
    return r


def rules_old(d):
    """Шаблон "Судга ариза.docx" — метки ЗАГЛАВНЫМИ буквами."""
    def g(key):
        return ph(d.get(key))

    return [
        (rf"{NUM}[-–—\s]*\(\s*ОСТАТОК\s*ДОЛГ\s*\+\s*100[\s\xa0]*000\s*\)\s*",
         g("jami") + " ", 0),
        (rf"{NUM}[-–—]\s*СУММА\s*ЗАКАЗА", g("summa"), 0),
        (rf"{NUM}[-–—]\s*ОСТАТОК\s*ДОЛГ", g("ostatok"), 0),
        (rf"{NUM}[-–—]\s*ПРОСРОЧКА\s*ДНЕЙ", g("days"), 0),
        (rf"{NUM}[-–—]\s*ПРО[СC]+РОЧКА\s*СУММА", g("prosrochka"), 0),
        (rf"{NUM}[-–—]?\s*ЗАКАЗ\s*ИД", g("id"), 0),
        (r"\b\d{1,2}[./]\d{1,2}[./]\d{4}\b", g("sana"), 0),
        (r"Ф\.?\s?И\.?\s?Ш\.?", g("fio"), 0),
    ]


LEFTOVERS_NEW = ["(фиш)", "(адрес)", "(тел)", "(пнфл)", "п.сериа", "(дата заказа)",
                 "(заказ ид)", "Сумма заказа", "Остаток основного долга",
                 "остаток долг", "Срок истечение"]
LEFTOVERS_OLD = ["ЗАКАЗ ИД", "СУММА ЗАКАЗА", "ОСТАТОК ДОЛГ", "ПРОСРОЧКА ДНЕЙ",
                 "ПРОССРОЧКА СУММА", "ПРОСРОЧКА СУММА", "Ф.И.Ш"]


def detect_template(text):
    if re.search(r"\(\s*заказ\s*ид\s*\)", text, I) or re.search(r"\(\s*фиш\s*\)", text, I):
        return "new"
    if "ЗАКАЗ ИД" in text or "ОСТАТОК ДОЛГ" in text:
        return "old"
    return "new"


def detect_template_fio(text):
    if TEMPLATE_FIO:
        return TEMPLATE_FIO
    m = re.search(r"([A-Z][A-Z‘’'`\-]*(?:\s+[A-Z][A-Z‘’'`\-]*){1,4})\s*\(\s*фиш\s*\)",
                  text, I)
    return m.group(1).strip() if m else ""


def make_document(template_path, data, out_path, kind, tpl_fio):
    doc = Document(template_path)
    rules = rules_new(data, tpl_fio) if kind == "new" else rules_old(data)
    for p in iter_paragraphs(doc):
        for pattern, value, flags in rules:
            replace_in_paragraph(p, pattern, value, flags)
    doc.save(out_path)

    text = doc_text(Document(out_path))
    marks = LEFTOVERS_NEW if kind == "new" else LEFTOVERS_OLD
    return [w for w in marks if w.lower() in text.lower()]


# ── PDF через Microsoft Word ───────────────────────────────────────

class PdfError(Exception):
    """PDF не сделать: нет Word, нет pywin32 или Word не запустился."""


class WordPdf:
    """.docx -> .pdf средствами установленного Word: шапка, логотип и шрифты
    остаются ровно как в документе. Word запускается один раз на всю пачку
    (старт ~2 с, дальше ~0.4 с на документ).

        conv = WordPdf(); conv.open()
        try:     conv.convert("a.docx", "a.pdf")
        finally: conv.close()
    """

    def __init__(self):
        self.word = None
        self._com = None
        self._tmp = None

    def open(self):
        try:
            import pythoncom
        except ImportError:
            raise PdfError("не установлен pywin32 (pip install pywin32)")
        self._com = pythoncom
        pythoncom.CoInitialize()          # COM инициализируется в каждом потоке заново
        try:
            self._start()
        except Exception:
            pythoncom.CoUninitialize()
            self._com = None
            raise
        self._tmp = tempfile.mkdtemp(prefix="wordpdf_")
        return self

    def _start(self):
        import win32com.client
        try:
            # DispatchEx — отдельный процесс Word: открытые у пользователя документы не трогаем
            self.word = win32com.client.DispatchEx("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0   # никаких диалогов — иначе Word молча зависнет
        except Exception as e:
            self.word = None
            raise PdfError(f"Microsoft Word не запустился: {e}")

    def _export(self, src, dst):
        doc = self.word.Documents.Open(src, ConfirmConversions=False, ReadOnly=True,
                                       AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17,   # 17 = PDF
                                    OpenAfterExport=False, OptimizeFor=0)  # 0 = для печати
        finally:
            doc.Close(0)                  # 0 = не сохранять изменения

    def convert(self, docx_path, pdf_path):
        # Word получает короткие латинские имена: кириллица, «O‘G‘LI» и длинные
        # пути в COM иногда дают ошибку, а переименовать готовый файл — надёжно.
        src = os.path.join(self._tmp, "in.docx")
        dst = os.path.join(self._tmp, "out.pdf")
        shutil.copyfile(docx_path, src)
        try:
            self._export(src, dst)
        except Exception:
            # Word мог упасть посреди пачки — перезапускаем один раз и пробуем снова
            self._quit()
            self._start()
            self._export(src, dst)
        shutil.move(dst, pdf_path)

    def _quit(self):
        if self.word is not None:
            try:
                self.word.Quit(0)
            except Exception:
                pass
        self.word = None

    def close(self):
        self._quit()
        if self._tmp:
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None
        if self._com is not None:
            self._com.CoUninitialize()
            self._com = None


# ── Чтение Excel ───────────────────────────────────────────────────

def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower().replace("ё", "е")


def map_headers(header_row):
    heads = [norm(c) for c in header_row]
    found, used = {}, set()
    # 1) точное совпадение
    for key, variants in COLUMNS.items():
        for i, h in enumerate(heads):
            if h and i not in used and any(h == norm(v) for v in variants):
                found[key] = i
                used.add(i)
                break
    # 2) частичное — начинается с варианта или наоборот
    for key, variants in COLUMNS.items():
        if key in found:
            continue
        for v in sorted(variants, key=len, reverse=True):
            nv = norm(v)
            for i, h in enumerate(heads):
                if not h or i in used:
                    continue
                # Обратное совпадение — только для заголовков от 3 букв, иначе
                # колонка "п" перехватит "пинфл", "т" — "тел" и т.д.
                if h.startswith(nv) or (len(h) >= 3 and nv.startswith(h)):
                    found[key] = i
                    used.add(i)
                    break
            if key in found:
                break
    return found


def find_header_row(rows, limit=10):
    """Заголовки не всегда в первой строке — сверху бывает название таблицы.
    Возвращает (индекс строки заголовков, карта колонок)."""
    best_i, best_cols = None, {}
    for i, row in enumerate(rows[:limit]):
        cols = map_headers(row)
        if "fio" in cols and "ostatok" in cols:
            return i, cols
        if len(cols) > len(best_cols):
            best_i, best_cols = i, cols
    return best_i, best_cols


def read_rows(excel_path, sud_default=None, pochta_default=None):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    try:
        # Данные не всегда на первом листе и не всегда в первой строке.
        ws = rows = None
        head_i, cols = None, {}
        for sheet in wb.worksheets:
            r = list(sheet.iter_rows(values_only=True))
            if not r:
                continue
            i, c = find_header_row(r)
            if "fio" in c and "ostatok" in c:
                ws, rows, head_i, cols = sheet, r, i, c
                break
            if ws is None:
                ws, rows, head_i, cols = sheet, r, i, c
    finally:
        wb.close()

    if not rows:
        return [], []

    missing = [k for k in ("fio", "ostatok") if k not in cols]
    if missing:
        seen = rows[head_i] if head_i is not None else rows[0]
        raise DataError(
            "В Excel не найдены обязательные колонки: " + ", ".join(missing) +
            "\nЛист: " + str(ws.title) +
            "\nЗаголовки в файле: " + ", ".join(str(c) for c in seen if c))

    def val(row, key):
        i = cols.get(key)
        return row[i] if i is not None and i < len(row) else None

    out, warns = [], []
    for n, row in enumerate(rows[head_i + 1:], start=head_i + 2):
        if not any(c not in (None, "") for c in row):
            continue
        if not str(val(row, "fio") or "").strip():
            continue

        ostatok = to_num(val(row, "ostatok")) or 0.0
        _sud_def = XIZMAT_HAQI if sud_default is None else sud_default
        _poc_def = POCHTA_HAQI if pochta_default is None else pochta_default
        sud = to_num(val(row, "sud"))
        sud = _sud_def if sud is None else sud
        pochta = to_num(val(row, "pochta")) if POCHTA_IZ_EXCEL else None
        pochta = _poc_def if pochta is None else pochta
        jami = round(ostatok + sud, 2)

        jami_x = to_num(val(row, "jami_col"))
        if jami_x is not None and abs(jami_x - jami) > 0.01:
            warns.append(f"строка {n}: «жами карздорликк» = {money(jami_x)}, "
                         f"а остаток+суд харажати = {money(jami)} — взято {money(jami)}")

        pinfl = clean_str(val(row, "pinfl"))
        days = to_num(val(row, "days"))
        rec = {
            # номер строки Excel — нужен в предупреждениях («16-qator ...»)
            "_qator":    n,
            "fio":       clean_str(val(row, "fio")),
            "id":        clean_str(val(row, "id")),
            "sana":      fmt_date(val(row, "sana")),
            "srok":      fmt_date(val(row, "srok")),
            "summa":     money(val(row, "summa")),
            "ostatok":   money(ostatok),
            "sud":       money(sud),
            "pochta":    money(pochta),
            "jami":      money(jami),
            "jami2":     money(round(jami + pochta, 2)),
            "pinfl":     pinfl,
            "tugilgan":  birth_from_pinfl(pinfl),
            "passport":  clean_str(val(row, "passport")),
            "tel":       clean_str(val(row, "tel")),
            "adres":     clean_str(val(row, "adres")),
            # пусто, а не 0: иначе в документе появится «0 кунлик» как факт
            "days":      "" if days is None else clean_str(int(days)),
            "prosrochka": money(val(row, "prosrochka")),
        }

        # Пустое поле превращается в прочерк — об этом лучше предупредить
        blanks = [name for name, key in (("ID заказа", "id"),
                                         ("Сумма заказа", "summa"),
                                         ("дата заказа", "sana"))
                  if not rec[key].strip() or rec[key] == "__.__.____"]
        if blanks:
            warns.append(f"строка {n} ({rec['fio']}): не заполнено — " + ", ".join(blanks))

        out.append(rec)
    return out, warns


def choose_best_xlsx(cands):
    best_path, best_score = None, -1
    for p in cands:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            try:
                ws = wb.worksheets[0]
                rows = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
            finally:
                wb.close()
            cols = find_header_row(rows)[1]
            score = 0
            if "fio" in cols:
                score += 4
            if "ostatok" in cols:
                score += 4
            if "sana" in cols:
                score += 1
            if "summa" in cols:
                score += 1
            if "sud" in cols:
                score += 1
            if "pochta" in cols:
                score += 1
            if score > best_score:
                best_path, best_score = p, score
        except Exception:
            continue
    return best_path


def choose_best_docx(cands):
    best_path, best_score = None, -1
    for p in cands:
        try:
            text = doc_text(Document(p))
        except Exception:
            continue
        score = 0
        if re.search(r"\(\s*фиш\s*\)", text, I):
            score += 5
        if re.search(r"\(\s*адрес\s*\)", text, I):
            score += 2
        if re.search(r"\(\s*заказ\s*ид\s*\)", text, I):
            score += 2
        if "ЗАКАЗ ИД" in text or "ОСТАТОК ДОЛГ" in text:
            score += 1
        if score > best_score:
            best_path, best_score = p, score
    return best_path


def find_file(preferred, extension):
    if preferred:
        p = os.path.join(BASE_DIR, preferred)
        if os.path.exists(p):
            return p
    cands = [f for f in glob.glob(os.path.join(BASE_DIR, "*" + extension))
             if not os.path.basename(f).startswith("~$")]
    if len(cands) == 1:
        return cands[0]
    if not cands:
        raise DataError(f"В папке нет файла {extension}: {BASE_DIR}")

    if extension == ".xlsx":
        best = choose_best_xlsx(cands)
        if best:
            return best
    elif extension == ".docx":
        best = choose_best_docx(cands)
        if best:
            return best

    raise DataError(
        f"В папке несколько файлов {extension}. Укажите нужный в настройках скрипта:\n  " +
        "\n  ".join(os.path.basename(c) for c in cands))


def main():
    excel_path = find_file(EXCEL_FILE, ".xlsx")
    template_path = find_file(TEMPLATE_FILE, ".docx")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    tpl_text = doc_text(Document(template_path))
    kind = detect_template(tpl_text)
    tpl_fio = detect_template_fio(tpl_text) if kind == "new" else ""

    print("Excel   :", os.path.basename(excel_path))
    print("Шаблон  :", os.path.basename(template_path),
          f"({'новый — метки в скобках' if kind == 'new' else 'старый — метки заглавными'})")
    if kind == "new" and not tpl_fio:
        print("  [!] В шаблоне не найдена метка (фиш) — ФИО может не подставиться.")
    print("Выгрузка:", OUTPUT_DIR)
    print("-" * 66)

    rows, problems = read_rows(excel_path)
    if not rows:
        print("В Excel нет данных.")
        return

    conv = None
    if OUTPUT_FORMAT == "pdf":
        print("Запускаю Microsoft Word для PDF...")
        try:
            conv = WordPdf().open()
        except PdfError as e:
            print(f"  [!] PDF не получится: {e}\n      Документы будут в Word (.docx).")
            conv = None

    ok = 0
    try:
        for i, d in enumerate(rows, 1):
            sana_file = d["sana"] if d["sana"] != "__.__.____" else \
                datetime.date.today().strftime("%d.%m.%Y")
            fname = safe_filename(FILE_NAME_TEMPLATE.format(
                fio=d["fio"], sana=sana_file, id=d["id"] or "б-н"))
            if not fname.lower().endswith(".docx"):
                fname += ".docx"
            out_path = os.path.join(OUTPUT_DIR, fname)

            try:
                left = make_document(template_path, d, out_path, kind, tpl_fio)
            except Exception as e:
                problems.append(f"{d['fio']}: ОШИБКА — {e}")
                print(f"[{i}/{len(rows)}]  x  {d['fio']}  —  {e}")
                continue

            if conv:
                pdf_path = os.path.splitext(out_path)[0] + ".pdf"
                try:
                    conv.convert(out_path, pdf_path)
                    os.remove(out_path)
                    fname = os.path.basename(pdf_path)
                except Exception as e:
                    # .docx остаётся в папке — документ не теряется
                    problems.append(f"{d['fio']}: PDF не сделан, оставлен .docx — {e}")

            ok += 1
            print(f"[{i}/{len(rows)}]  {'+' if not left else '!'}  {fname}")
            print(f"          заказ {d['id']} от {d['sana']} | остаток {d['ostatok']}"
                  f" | +суд {d['sud']} = {d['jami']} | +почта {d['pochta']} = {d['jami2']}")
            if left:
                problems.append(f"{d['fio']}: не заменено — {', '.join(left)}")
    finally:
        if conv:
            conv.close()

    print("-" * 66)
    print(f"Готово: {ok} из {len(rows)} документов -> {OUTPUT_DIR}")
    if problems:
        print("\nВНИМАНИЕ, проверьте вручную:")
        for p in problems:
            print("  -", p)


if __name__ == "__main__":
    try:
        main()
    except (DataError, SystemExit) as e:
        print(e)
    except Exception:
        import traceback
        traceback.print_exc()
    if os.name == "nt":
        input("\nНажмите Enter для выхода...")
