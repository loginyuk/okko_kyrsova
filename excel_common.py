"""
Спільне для кроків 3 і 4: читання data/2_values.csv, стилі, формули показників і типові аркуші
(форма звітності по роках, показники, таблиця Опендатаботу, джерела).
"""
import math
import re
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from common import EXTRACTED_DIR, FORM_NAMES, FULL_NAME, REFERENCE_DIR, ROLE, SHEET_PREFIX, VALUES_CSV, YEARS, read_csv

SHEET_SUFFIX = {"001": "Баланс", "002": "Фінрез", "003": "Рух_коштів", "040": "Капітал"}
TITLES = {
    "001": "Форма 1. Баланс (звіт про фінансовий стан) на 31 грудня",
    "002": "Форма 2. Звіт про фінансові результати за рік",
    "003": "Форма 3. Звіт про рух грошових коштів (прямий метод) за рік",
    "040": "Форма 4. Звіт про власний капітал",
}

# Розрахункові рядки під формою (формули). Перевірки мають давати 0.
EXTRA_ROWS = {
    "001": [("Перевірка: актив − пасив (має бути 0)", "Ф1[1300] - Ф1[1900]"),
            ("Усього зобов’язання (розд. II + III)", "Ф1[1595] + Ф1[1695]")],
    "002": [("Перевірка: дохід − собівартість − валовий результат (має бути 0)",
             "Ф2[2000] - Ф2[2050] - (Ф2[2090] - Ф2[2095])")],
    "003": [("Перевірка: залишок на кінець − (початок + рух + курсові) (має бути 0)",
             "Ф3[3415] - (Ф3[3405] + Ф3[3400] + Ф3[3410])")],
}

# ── Стилі ─────────────────────────────────────────────────────────────────────
BLUE = Font(color="0000FF")                       # число зі звітності
GREY = Font(color="808080")                       # н/д
BOLD = Font(bold=True)
ITALIC = Font(italic=True)
TITLE = Font(bold=True, size=12)
SUBTITLE = Font(color="595959", size=9)
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
YELLOW = PatternFill("solid", fgColor="FFF2CC")   # не з оригінальної форми: сайт Опендатаботу або ручні дані
NOT_FROM_FORMS = ("агрегатор", "ручний")           # такі значення фарбуємо жовтим
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
RIGHT = Alignment(horizontal="right")
NUMBER_FORMATS = {
    "тис. грн": "#,##0;\\(#,##0\\);\\-",
    "%": "0.0%;\\(0.0%\\);\\-",
    "0.00": "0.00;\\(0.00\\);\\-",
    "0.000": "0.000;\\(0.000\\);\\-",
}


# ── Дані ──────────────────────────────────────────────────────────────────────

def load_values():
    """Повертає (обрані, усі): {(компанія, форма, код, графа): {рік: рядок CSV}} і список усіх рядків."""
    all_rows = read_csv(VALUES_CSV)
    chosen = defaultdict(dict)
    for r in all_rows:
        if r["обрано"]:
            chosen[(r["компанія"], r["форма"], r["код"], r["графа"])][int(r["рік"])] = r
    return chosen, all_rows


def number(text):
    x = float(text)
    return int(x) if x.is_integer() else x


def load_reference():
    names = {(r["форма"], r["код"]): r["назва"] for r in read_csv(REFERENCE_DIR / "row_names.csv")}
    notes = {(r["компанія"], r["форма"], r["код"]): r["примітка"] for r in read_csv(REFERENCE_DIR / "notes.csv")}
    return names, notes


def load_opendatabot_summary():
    """Таблиці з сайту Опендатаботу: {компанія: {(показник, одиниця): {період: число}}} і список періодів."""
    tables, periods = defaultdict(dict), []
    for r in read_csv(EXTRACTED_DIR / "opendatabot_summary.csv"):
        tables[r["компанія"]].setdefault((r["показник"], r["одиниця"]), {})[r["період"]] = number(r["значення"])
        if r["період"] not in periods:
            periods.append(r["період"])
    return tables, periods


# ── Формули: «Ф2[2090] / Ф2[2000]» → посилання на клітинки ─────────────────────
# Ф<номер форми>[<код рядка>] — значення за рік стовпця; з «(t-1)» — за попередній рік.
REF = re.compile(r"Ф(\d)\[(\d{4})\](\(t-1\))?")
FORM_BY_DIGIT = {"1": "001", "2": "002", "3": "003"}
row_of = {}   # (компанія, форма, код) → номер рядка на аркуші форми; заповнюється під час запису


def new_workbook(first_sheet_title):
    """Нова книга. Пам'ять про рядки форм очищаємо, щоб формули не послалися на рядки з іншої книги."""
    row_of.clear()
    wb = Workbook()
    wb.active.title = first_sheet_title
    return wb


def sheet_name(company, form):
    return f"{SHEET_PREFIX[company]}_{SHEET_SUFFIX[form]}"


def year_column(year):
    return get_column_letter(3 + YEARS.index(year))     # перший рік — стовпець C


def excel_formula(template, company, year):
    def cell_ref(m):
        form, code, previous = FORM_BY_DIGIT[m.group(1)], m.group(2), m.group(3)
        y = year - 1 if previous else year
        if y not in YEARS:
            return "#NA"                                 # попереднього року в таблиці немає
        return f"'{sheet_name(company, form)}'!{year_column(y)}{row_of[(company, form, code)]}"

    expr = REF.sub(cell_ref, template).replace(" ", "")
    if "#NA" in expr:
        return "н/д"
    return f'=IFERROR({expr},"н/д")'


def readable_formula(template):
    """«Ф2[2090] / Ф2[2000](t-1)» → «р.2090 / р.2000(t−1)» для стовпця «Формула»."""
    return REF.sub(lambda m: f"р.{m.group(2)}" + ("(t−1)" if m.group(3) else ""), template)


def codes_used_in_formulas(templates):
    """Рядки, на які посилаються формули, — їх показуємо завжди, навіть якщо там нулі."""
    used = defaultdict(set)
    for t in templates:
        for digit, code, _ in REF.findall(t):
            used[FORM_BY_DIGIT[digit]].add(code)
    return used


# ── Допоміжне для оформлення ──────────────────────────────────────────────────

def write_header(ws, row, titles):
    for col, title in enumerate(titles, start=1):
        c = ws.cell(row, col, title)
        c.font, c.fill, c.border = HEADER_FONT, HEADER_FILL, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put_value(cell, record, number_format=NUMBER_FORMATS["тис. грн"]):
    """Число зі звітності — синім; сайт/ручні дані — на жовтому тлі; немає даних — сіре «н/д»."""
    if record is None:
        cell.value, cell.font, cell.alignment = "н/д", GREY, RIGHT
    else:
        cell.value, cell.font = number(record["значення"]), BLUE
        if record["тип"] in NOT_FROM_FORMS:
            cell.fill = YELLOW
    cell.number_format = number_format
    cell.border = BORDER


def set_widths(ws, widths):
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def write_title(ws, title, subtitle=""):
    ws["A1"], ws["A1"].font = title, TITLE
    if subtitle:
        ws["A2"], ws["A2"].font = subtitle, SUBTITLE


def write_text_rows(ws, row, items, last_col):
    """Рядки «тема | текст»: текст займає стовпці B…last_col і переноситься."""
    for topic, text in items:
        ws.cell(row, 1, topic).font = BOLD
        ws.cell(row, 1).alignment = WRAP
        ws.cell(row, 2, text).alignment = WRAP
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
        ws.row_dimensions[row].height = 15 * max(1, math.ceil(len(text) / 95))
        row += 1
    return row


def source_label(record):
    if record["тип"] == "ручний":
        return record["джерело"]
    if record["тип"] == "агрегатор":
        return f"{record['джерело']}, підсумки (не з форми)"
    label = f"{record['джерело']}, звіт за {record['звіт за рік']} р."
    return label + " (порівняльна колонка)" if record["тип"] == "порівняльний" else label


def year_sources(chosen, company, form):
    """{рік: найчастіше джерело серед значень цього року}."""
    result = {}
    for year in YEARS:
        labels = Counter(source_label(by_year[year]) for (c, f, _, _), by_year in chosen.items()
                         if c == company and f == form and year in by_year)
        result[year] = labels.most_common(1)[0][0] if labels else None
    return result


def availability(records):
    if not records:
        return "✗ немає"
    if all(r["тип"] in NOT_FROM_FORMS for r in records):
        return "◐ лише підсумки"
    main = Counter((r["джерело"], r["звіт за рік"]) for r in records if r["тип"] not in NOT_FROM_FORMS)
    source, report_year = main.most_common(1)[0][0]
    return f"✓ {source} {report_year}"


# ── Типові аркуші ─────────────────────────────────────────────────────────────

def write_form_sheet(wb, company, form, chosen, names, notes, always_show, codes=None, extra_rows=True):
    """Аркуш форми 1–3: рядки — коди, стовпці — роки. Повертає (аркуш, перший вільний рядок).

    codes=None — показати всі рядки з ненульовими числами + потрібні формулам; або явний список кодів.
    """
    ws = wb.create_sheet(sheet_name(company, form))
    sources = year_sources(chosen, company, form)
    write_title(ws, f"{FULL_NAME[company]} — {TITLES[form]}, тис. грн",
                "Джерело: " + "; ".join(f"{y} — {s or 'немає даних'}" for y, s in sources.items()))
    write_header(ws, 4, ["Код рядка", "Стаття", *map(str, YEARS), "Примітка"])

    if codes is None:
        codes = {code for (c, f, code, _), by_year in chosen.items()
                 if c == company and f == form and any(float(r["значення"]) for r in by_year.values())}
        codes = sorted(codes | always_show[form], key=int)

    row = 5
    for code in codes:
        row_of[(company, form, code)] = row
        ws.cell(row, 1, code).border = BORDER
        ws.cell(row, 2, names.get((form, code), f"рядок {code}")).border = BORDER
        for i, year in enumerate(YEARS):
            put_value(ws.cell(row, 3 + i), chosen.get((company, form, code, ""), {}).get(year))
        note = ws.cell(row, 3 + len(YEARS), notes.get((company, form, code)))
        note.border, note.alignment = BORDER, WRAP
        row += 1

    row += 1
    for title, template in (EXTRA_ROWS.get(form, []) if extra_rows else []):
        ws.cell(row, 2, title).font = ITALIC
        for i, year in enumerate(YEARS):
            c = ws.cell(row, 3 + i, excel_formula(template, company, year))
            c.number_format = NUMBER_FORMATS["тис. грн"]
        row += 1

    set_widths(ws, [10, 62] + [14] * len(YEARS) + [46])
    ws.freeze_panes = "C5"
    return ws, row


def write_indicators_sheet(wb, indicators, companies):
    """Показники — формули Excel, що посилаються на аркуші форм (їх треба створити раніше)."""
    ws = wb.create_sheet("Показники")
    write_title(ws, "Аналітичні показники (формули з аркушів форм). Показники балансу — на кінець року.",
                "«н/д» — для розрахунку бракує даних. Список показників і формули — у reference/indicators.csv.")
    row = 4
    for company in companies:
        ws.cell(row, 1, f"{FULL_NAME[company]} ({ROLE[company]})").font = BOLD
        write_header(ws, row + 1, ["Група", "Показник", *map(str, YEARS), "Формула", "Коментар"])
        row += 2
        for ind in indicators:
            ws.cell(row, 1, ind["група"])
            ws.cell(row, 2, ind["показник"])
            for i, year in enumerate(YEARS):
                c = ws.cell(row, 3 + i, excel_formula(ind["формула"], company, year))
                c.number_format = NUMBER_FORMATS[ind["формат"]]
                if c.value == "н/д":
                    c.font, c.alignment = GREY, RIGHT
            ws.cell(row, 3 + len(YEARS), readable_formula(ind["формула"]))
            ws.cell(row, 4 + len(YEARS), ind["коментар"])
            for col in range(1, 5 + len(YEARS)):
                ws.cell(row, col).border = BORDER
            row += 1
        row += 1
    set_widths(ws, [20, 55] + [12] * len(YEARS) + [42, 50])


def write_availability(ws, row, chosen, companies, forms):
    """Матриця «форма × рік → джерело даних». Повертає перший вільний рядок."""
    ws.cell(row, 1, "Наявність даних за роками").font = BOLD
    ws.cell(row + 1, 1, "У клітинці — джерело і рік звіту, з якого взято дані "
                        "(2021 і 2024 — порівняльні колонки звітів за 2022 і 2025 рр.).").font = SUBTITLE
    row += 2
    for company in companies:
        write_header(ws, row, [FULL_NAME[company], *map(str, YEARS)])
        row += 1
        for form in forms:
            ws.cell(row, 1, FORM_NAMES[form])
            for i, year in enumerate(YEARS):
                records = [by_year[year] for (c, f, _, _), by_year in chosen.items()
                           if c == company and f == form and year in by_year]
                c = ws.cell(row, 2 + i, availability(records))
                c.alignment = Alignment(horizontal="center")
                if records and all(r["тип"] in NOT_FROM_FORMS for r in records):
                    c.fill = YELLOW
            for col in range(1, 2 + len(YEARS)):
                ws.cell(row, col).border = BORDER
            row += 1
        row += 1
    return row


def write_opendatabot_sheet(wb, companies, subtitle):
    """Таблиці «Фінансові показники» з сайту Опендатаботу — як на скріншотах."""
    ws = wb.create_sheet("Опендатабот")
    write_title(ws, "Опендатабот: «Фінансові показники» з сайту (підсумки агрегатора, не оригінальні форми)",
                subtitle)
    formats = {"грн": "#,##0", "осіб": "#,##0", "%": '0.00"%"'}
    tables, periods = load_opendatabot_summary()

    row = 4
    for company in companies:
        ws.cell(row, 1, f"{FULL_NAME[company]} ({ROLE[company]})").font = BOLD
        write_header(ws, row + 1, ["Показник", "Одиниця", *periods])
        row += 2
        for (indicator, unit), values in tables[company].items():
            ws.cell(row, 1, indicator).border = BORDER
            ws.cell(row, 2, unit).border = BORDER
            for i, period in enumerate(periods):
                c = ws.cell(row, 3 + i, values.get(period, "—"))
                c.border, c.number_format = BORDER, formats.get(unit, "General")
                if c.value == "—":
                    c.font, c.alignment = GREY, RIGHT
            row += 1
        row += 1
    set_widths(ws, [36, 9] + [17] * len(periods))


def write_sources_sheet(wb, items, report_ids):
    """items — пари (що, джерело); report_ids — використані звіти."""
    ws = wb.create_sheet("Джерела")
    write_header(ws, 1, ["Що", "Джерело"])
    row = 2
    for what, source in items:
        ws.cell(row, 1, what).alignment = WRAP
        ws.cell(row, 2, source).alignment = WRAP
        row += 1

    row += 1
    ws.cell(row, 1, "Використані звіти (останні подання; оригінали — у data/1_extracted/originals/)").font = BOLD
    row += 1
    for report_id in sorted(report_ids):
        ws.cell(row, 1, report_id)
        row += 1
    set_widths(ws, [60, 100])
