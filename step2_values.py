"""
Крок 2. Зводить числа з усіх звітів в одну «довгу» таблицю data/2_values.csv:
один рядок CSV = одне число (компанія, форма, рік, код рядка, графа).

Кожен звіт містить два роки: звітний і попередній (порівняльна колонка).
Тому одне й те саме значення може прийти з кількох місць — «кандидатів»:
  1) звітна колонка звіту за цей рік                       (найнадійніше)
  2) порівняльна колонка звіту за наступний рік            (напр., 2021 — зі звіту за 2022)
  3) підсумки з сайту Опендатаботу (агрегатор)            (лише якщо форм за цей рік немає)
  4) ручні дані з reference/manual_values.csv
Якщо компанія подавала звіт кілька разів (уточнення), береться лише найпізніше подання.
У таблиці лишаються всі кандидати, а обраний позначено «так» у колонці «обрано».

Запуск:  python3 step2_values.py
"""
import re
from collections import defaultdict

from common import AGGREGATOR, EXTRACTED_DIR, FORMS_IN_REPORT, REFERENCE_DIR, VALUES_CSV, YEARS, read_csv, write_csv

COLUMNS = ["компанія", "форма", "рік", "код", "графа", "значення", "тип",
           "джерело", "звіт за рік", "id звіту", "обрано"]
PRIORITY = {"звітний": 1, "порівняльний": 2, "агрегатор": 3, "ручний": 4}


def latest_reports():
    """Для кожної (компанія, форма, звітний рік, джерело) — лише останнє подання."""
    latest = {}
    for r in read_csv(EXTRACTED_DIR / "reports.csv"):
        if r["форма"] not in FORMS_IN_REPORT:
            continue
        key = (r["компанія"], r["форма"], r["звітний рік"], r["джерело"])
        order = (r["дата подання"], int(r["№ документа"] or 0))
        if key not in latest or order > latest[key][0]:
            latest[key] = (order, r)
    return {r["id"]: r for _, r in latest.values()}


def parse_field(form, field, report_year):
    """Перекладає назву поля XML у (код рядка, рік, графа, тип).

    Графи паперової форми: у Ф1 графа 3 — «на початок року» (тобто кінець попереднього),
    графа 4 — «на кінець року»; у Ф2 і Ф3 графа 3 — звітний рік, графа 4 — попередній;
    у Ф4 графи 3…10 — складові власного капіталу, усі за звітний рік.
    """
    m = re.fullmatch(r"([AB])(\d{4})(?:_(\d+))?", field)
    if not m:
        return None                                  # текстові поля (назва фірми тощо)
    letter, code, suffix = m.groups()
    column = int(suffix) if suffix else (3 if letter == "A" else 4)

    if form == "040":
        return code, report_year, str(column), "звітний"
    current_year_column = 4 if form == "001" else 3
    if column == current_year_column:
        return code, report_year, "", "звітний"
    return code, report_year - 1, "", "порівняльний"


def to_number(text):
    """Порожнє поле у формі означає 0. Решта — ціле або дробове число."""
    text = (text or "").strip()
    if text == "":
        return "0"
    return text


def main():
    reports = latest_reports()
    candidates = []

    for f in read_csv(EXTRACTED_DIR / "fields.csv"):
        report = reports.get(f["id звіту"])
        if not report:
            continue                                  # старе подання або непотрібна форма
        parsed = parse_field(report["форма"], f["поле"], int(report["звітний рік"]))
        if not parsed:
            continue
        code, year, column, kind = parsed
        if report["джерело"] == AGGREGATOR:
            kind = "агрегатор"
        candidates.append({"компанія": report["компанія"], "форма": report["форма"], "рік": year,
                           "код": code, "графа": column, "значення": to_number(f["значення"]),
                           "тип": kind, "джерело": report["джерело"],
                           "звіт за рік": report["звітний рік"], "id звіту": report["id"]})

    for m in read_csv(REFERENCE_DIR / "manual_values.csv"):
        candidates.append({**m, "рік": int(m["рік"]), "графа": m.get("графа", ""), "тип": "ручний"})

    # обираємо найкращого кандидата для кожного значення
    groups = defaultdict(list)
    for c in candidates:
        if c["рік"] in YEARS:
            groups[(c["компанія"], c["форма"], c["рік"], c["код"], c["графа"])].append(c)
    rows = []
    for key in sorted(groups):
        group = sorted(groups[key], key=lambda c: PRIORITY[c["тип"]])
        for i, c in enumerate(group):
            c["обрано"] = "так" if i == 0 else ""
        rows.extend(group)

    write_csv(VALUES_CSV, rows, COLUMNS)
    chosen = sum(1 for r in rows if r["обрано"])
    print(f"Звітів (останні подання): {len(reports)}")
    print(f"Крок 2 готово: {chosen} значень (усього кандидатів {len(rows)}) → {VALUES_CSV}")


if __name__ == "__main__":
    main()
