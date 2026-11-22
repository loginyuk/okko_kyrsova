"""
Крок 1. Витягує з сирих даних усі звіти компаній ОККО.

Звідки:
  • Держстат (data.gov.ua): папки й zip-архіви, де лежать звіти ВСІХ підприємств
    України — по одному XML на звіт, назва файлу починається з ЄДРПОУ;
  • Опендатабот: xlsx-файли зі звітністю за 2023 рік (форми 1–2 ОККО-ДРАЙВ);
  • Опендатабот, сайт: таблиці «Фінансові показники» (переписані зі скріншотів у
    raw_data/opendatabot/фінпоказники_<компанія>_<ЄДРПОУ>.csv) — підсумки агрегатора:
    дохід, чистий прибуток, активи, працівники тощо.

Куди (data/1_extracted/):
  originals/    оригінали знайдених звітів (XML без цифрового підпису; xlsx як є)
  reports.csv   перелік звітів: компанія, форма, звітний рік, дата подання, звідки взято
  fields.csv    усі поля всіх звітів: id звіту, поле (напр. A1300), значення
  opendatabot_summary.csv   таблиці «Фінансові показники» з сайту в довгому форматі

Поля названо так, як у XML Держстату: A1300 — рядок 1300, графа 3; B1300 — графа 4;
A3000_3 / A3000_4 — те саме для Ф3; A4000_7 — Ф4, графа 7. Дані Опендатаботу
перекладаються в ці ж назви, тож далі всі джерела обробляються однаково.
З таблиць сайту в рядки форм ідуть лише показники, що мають точний відповідник:
дохід → р.2000, чистий прибуток → р.2350 (або збиток → р.2355), активи → р.1300 і р.1900.

Запуск:  python3 step1_extract.py
"""
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from common import (AGGREGATOR, COMPANIES, DERZHSTAT_SOURCES, EXTRACTED_DIR, FORM_NAMES,
                    OPENDATABOT_DIR, read_csv, write_csv)

ORIGINALS = EXTRACTED_DIR / "originals"
REPORT_COLUMNS = ["id", "компанія", "ЄДРПОУ", "форма", "назва форми", "звітний рік",
                  "дата подання", "№ документа", "джерело", "звідки взято", "оригінал"]

reports = []   # рядки для reports.csv
fields = []    # рядки для fields.csv
summary = []   # рядки для opendatabot_summary.csv


def save_original(company, name, data):
    path = ORIGINALS / company / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return str(path.relative_to(EXTRACTED_DIR))


def add_report(meta, values, original_bytes=None, original_ext=""):
    """Зберігає один знайдений звіт: оригінал у originals/ (якщо переданий), метадані й поля — у списки."""
    company = COMPANIES[meta["ЄДРПОУ"]]
    filed = meta.get("дата подання", "")
    report_id = "_".join(x for x in [
        company, meta["звітний рік"], "ф" + meta["форма"],
        filed.replace("-", "") or re.sub(r"\W+", "-", meta["джерело"].lower()).strip("-"),
        "№" + meta["№ документа"] if meta.get("№ документа") else "",
    ] if x)

    if original_bytes is not None:
        meta["оригінал"] = save_original(company, report_id + original_ext, original_bytes)

    reports.append({**meta, "id": report_id, "компанія": company,
                    "назва форми": FORM_NAMES.get(meta["форма"], "")})
    fields.extend({"id звіту": report_id, "поле": k, "значення": v} for k, v in values.items())
    print(f"  ✓ {report_id}")


# ── Держстат: XML ─────────────────────────────────────────────────────────────

def is_our_file(name):
    """Файли Держстату названо «<ЄДРПОУ>_…xml», тож чужі звіти відсіюємо за назвою."""
    base = Path(name).name
    return base.lower().endswith(".xml") and base[:8] in COMPANIES


def read_derzhstat_xml(data, where):
    # Звіти, підписані КЕП (M.E.Doc), мають бінарний підпис до і після XML — вирізаємо сам XML.
    start, end = data.find(b"<?xml"), data.rfind(b"</DECLAR>")
    if start < 0 or end < 0:
        print(f"  ! у {where} немає XML-звіту, пропускаю")
        return
    xml_bytes = data[start:end + len(b"</DECLAR>")]
    root = ET.fromstring(xml_bytes)

    head = {e.tag: (e.text or "").strip() for e in root.find("DECLARHEAD")}
    body = {e.tag: (e.text or "").strip() for e in root.find("DECLARBODY")}
    if head.get("TIN") not in COMPANIES:
        return
    d = head.get("D_FILL", "")            # ддммрррр → рррр-мм-дд
    meta = {
        "ЄДРПОУ": head["TIN"],
        "форма": head.get("C_DOC_SUB", ""),
        "звітний рік": head.get("PERIOD_YEAR", ""),
        "дата подання": f"{d[4:]}-{d[2:4]}-{d[:2]}" if len(d) == 8 else "",
        "№ документа": head.get("C_DOC_CNT", ""),
        "джерело": "Держстат",
        "звідки взято": where,
    }
    add_report(meta, body, xml_bytes, ".xml")


def scan_derzhstat(path):
    if path.is_dir():
        for f in sorted(path.rglob("*")):
            if f.suffix.lower() == ".zip" or is_our_file(f.name):
                scan_derzhstat(f)
    elif path.suffix.lower() == ".zip":
        print(f"архів: {path.name}")
        with zipfile.ZipFile(path) as z:
            for member in z.namelist():          # сотні тисяч файлів, але відкриваємо лише свої
                if is_our_file(member):
                    read_derzhstat_xml(z.read(member), f"{path.name} → {member}")
    elif is_our_file(path.name):
        read_derzhstat_xml(path.read_bytes(), str(path))


# ── Опендатабот: xlsx ─────────────────────────────────────────────────────────

def read_xlsx(path):
    """Мінімальний читач xlsx: повертає назву першого аркуша і рядки як {колонка: текст}.
    (openpyxl не відкриває файли Опендатаботу через нестандартні стилі, тому читаємо XML усередині.)"""
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as z:
        sheet_name = ET.fromstring(z.read("xl/workbook.xml")).find(f".//{ns}sheet").get("name")
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
                shared.append("".join(t.text or "" for t in si.iter(ns + "t")))
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

    rows = []
    for row in sheet.iter(ns + "row"):
        cells = {}
        for c in row.iter(ns + "c"):
            column = re.match(r"[A-Z]+", c.get("r")).group()
            v = c.find(ns + "v")
            if c.get("t") == "s":
                text = shared[int(v.text)] if v is not None else ""
            elif c.get("t") == "inlineStr":
                text = "".join(t.text or "" for t in c.iter(ns + "t"))
            else:
                text = v.text if v is not None else ""
            cells[column] = (text or "").strip()
        rows.append(cells)
    return sheet_name, rows


def clean_number(text):
    """«-», «(-)» і порожньо → '' (як порожнє поле в XML); «(123)» → '123'; прибирає пробіли."""
    text = (text or "").replace(" ", "").replace("\xa0", "")
    if text in ("", "-", "(-)"):
        return ""
    return text.strip("()")


def xml_field(form, code, column):
    """Назва поля як у XML Держстату: Ф1/Ф2 — A/B (графи 3/4), Ф3 — A…_3 / A…_4."""
    if form in ("001", "002"):
        return ("A" if column == 3 else "B") + code
    return f"A{code}_{column}"


def read_opendatabot(path):
    sheet_name, rows = read_xlsx(path)   # аркуш названо кодом бланку, напр. S0100114 → форма 001
    form = sheet_name[3:6]
    header_text = " ".join(" ".join(r.values()) for r in rows[:20])
    tin = next(r["D"].split()[0] for r in rows if r.get("C") == "за ЄДРПОУ")
    year = re.search(r"(?:Рік|31\.12\.)\s*(20\d\d)", header_text).group(1)
    if tin not in COMPANIES:
        return

    values = {}
    for r in rows:                       # рядок форми: A — стаття, B — код, C — графа 3, D — графа 4
        code = r.get("B", "")
        if re.fullmatch(r"\d{4}", code):
            values[xml_field(form, code, 3)] = clean_number(r.get("C"))
            values[xml_field(form, code, 4)] = clean_number(r.get("D"))

    meta = {"ЄДРПОУ": tin, "форма": form, "звітний рік": year, "джерело": "Опендатабот",
            "звідки взято": str(path.relative_to(OPENDATABOT_DIR.parent))}
    add_report(meta, values, path.read_bytes(), ".xlsx")


# ── Опендатабот, сайт: таблиці «Фінансові показники» (зі скріншотів) ─────────

def thousands(hryvnias):
    """На сайті суми в гривнях, у формах звітності — в тис. грн."""
    x = float(hryvnias) / 1000
    return str(int(x)) if x.is_integer() else str(x)


def read_opendatabot_site(path):
    tin = re.search(r"\d{8}", path.name).group()     # ЄДРПОУ записано в назві файлу
    company = COMPANIES[tin]
    table = {row["показник"]: row for row in read_csv(path)}

    # 1) уся таблиця як є — для окремого аркуша в Excel
    for row in table.values():
        for period, value in row.items():
            if period not in ("показник", "одиниця") and value:
                summary.append({"компанія": company, "показник": row["показник"],
                                "одиниця": row["одиниця"], "період": period, "значення": value})

    # 2) показники з точним відповідником у формах — як «звіт» за кожен рік
    original = save_original(company, path.name, path.read_bytes())
    screenshot = path.with_suffix(".png")            # скріншот, з якого переписано таблицю
    if screenshot.exists():
        save_original(company, screenshot.name, screenshot.read_bytes())
    for year in (p for p in table["Дохід"] if p.isdigit()):
        meta = {"ЄДРПОУ": tin, "звітний рік": year, "джерело": AGGREGATOR,
                "звідки взято": str(path.relative_to(OPENDATABOT_DIR.parent)), "оригінал": original}
        income, profit, assets = (table[k][year] for k in ("Дохід", "Чистий прибуток", "Активи"))
        if income or profit:
            values = {}
            if income:
                values["A2000"] = thousands(income)
            if profit:                                   # прибуток → р.2350, збиток → р.2355
                loss = float(profit) < 0
                values["A2350"] = "" if loss else thousands(profit)
                values["A2355"] = thousands(-float(profit)) if loss else ""
            add_report({**meta, "форма": "002"}, values)
        if assets:                                       # актив балансу = пасив балансу
            add_report({**meta, "форма": "001"}, {"B1300": thousands(assets), "B1900": thousands(assets)})


# ──────────────────────────────────────────────────────────────────────────────

def main():
    shutil.rmtree(EXTRACTED_DIR, ignore_errors=True)   # кожен запуск — з чистого аркуша
    for source in DERZHSTAT_SOURCES:
        scan_derzhstat(source)
    print(f"Опендатабот: {OPENDATABOT_DIR.name}")
    for path in sorted(OPENDATABOT_DIR.glob("*.xlsx")):
        read_opendatabot(path)
    for path in sorted(OPENDATABOT_DIR.glob("фінпоказники_*.csv")):
        read_opendatabot_site(path)

    reports.sort(key=lambda r: (r["компанія"], r["звітний рік"], r["форма"],
                                r.get("дата подання", ""), int(r.get("№ документа") or 0)))
    write_csv(EXTRACTED_DIR / "reports.csv", reports, REPORT_COLUMNS)
    write_csv(EXTRACTED_DIR / "fields.csv", fields, ["id звіту", "поле", "значення"])
    write_csv(EXTRACTED_DIR / "opendatabot_summary.csv", summary,
              ["компанія", "показник", "одиниця", "період", "значення"])
    print(f"\nКрок 1 готово: {len(reports)} звітів → {EXTRACTED_DIR}")


if __name__ == "__main__":
    main()
