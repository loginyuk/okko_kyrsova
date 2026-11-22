"""
Крок 3. Повна таблиця по обох компаніях: output/ОККО_фінзвітність_2021-2025.xlsx

Аркуші:
  Огляд                 про компанії, звідки дані за кожен рік, пояснення
  ОД_Баланс … ГНГ_Капітал  форми 1–4 по роках (сині цифри — зі звітності, жовте тло — підсумки агрегатора)
  Показники             коефіцієнти — це формули Excel, що посилаються на аркуші форм
  Опендатабот           таблиці «Фінансові показники» з сайту (переписані зі скріншотів)
  Розбіжності           де два джерела дають різні числа для того самого року
  Джерела               звідки взято дані

Спільні функції (стилі, формули, аркуші форм) — у excel_common.py; тексти — у reference/.
Скорочену таблицю для курсової будує крок 4 (step4_kursova.py) — він від цього файлу не залежить.

Запуск:  .venv/bin/python step3_excel.py
"""
from collections import defaultdict

from common import COMPANIES, FULL_NAME, OUTPUT_XLSX, REFERENCE_DIR, YEARS, read_csv
from excel_common import (BORDER, BOLD, EXTRA_ROWS, NUMBER_FORMATS, SHEET_SUFFIX, TITLES, codes_used_in_formulas,
                          load_reference, load_values, new_workbook, number, put_value, set_widths, sheet_name,
                          source_label, write_availability, write_form_sheet, write_header, write_indicators_sheet,
                          write_opendatabot_sheet, write_sources_sheet, write_text_rows, write_title)

F4_COLUMNS = {"3": "Зареєстрований капітал", "4": "Капітал у дооцінках", "5": "Додатковий капітал",
              "6": "Резервний капітал", "7": "Нерозподілений прибуток (непокритий збиток)",
              "8": "Неоплачений капітал", "9": "Вилучений капітал", "10": "Всього"}


def write_capital_sheet(wb, company, chosen, names):
    """Ф4 подається лише за звітний рік і має багато граф — окремий блок на кожен рік."""
    ws = wb.create_sheet(sheet_name(company, "040"))
    write_title(ws, f"{FULL_NAME[company]} — {TITLES['040']}, тис. грн",
                "Форма 4 містить дані лише за звітний рік, тому тут є тільки роки, за які подано звіт.")
    row = 4
    for year in YEARS:
        codes = sorted({code for (c, f, code, _), by_year in chosen.items()
                        if c == company and f == "040" and year in by_year
                        and float(by_year[year]["значення"])}, key=int)
        if not codes:
            continue
        ws.cell(row, 1, f"Звіт за {year} р.").font = BOLD
        write_header(ws, row + 1, ["Код рядка", "Стаття", *F4_COLUMNS.values()])
        row += 2
        for code in codes:
            ws.cell(row, 1, code).border = BORDER
            ws.cell(row, 2, names.get(("040", code), f"рядок {code}")).border = BORDER
            for i, column in enumerate(F4_COLUMNS):
                put_value(ws.cell(row, 3 + i), chosen.get((company, "040", code, column), {}).get(year))
            row += 1
        row += 1
    set_widths(ws, [10, 50] + [15] * len(F4_COLUMNS))


def write_differences_sheet(wb, all_rows, names):
    """Значення, для яких різні джерела дають різні числа."""
    ws = wb.create_sheet("Розбіжності")
    write_title(ws, "Розбіжності між джерелами для того самого року",
                "Наприклад, дані за 2022 р. є і у звіті за 2022 р., і в порівняльній колонці звіту за 2023 р. "
                "Або підсумки з сайту Опендатаботу відрізняються від форм. У таблицях використано «обране» "
                "значення (пріоритет: звітна колонка → порівняльна колонка → сайт Опендатаботу → ручні дані).")
    write_header(ws, 4, ["Компанія", "Форма", "Рік", "Код", "Стаття", "Графа",
                         "Обране значення", "Джерело обраного", "Інше значення", "Джерело іншого", "Різниця"])
    groups = defaultdict(list)
    for r in all_rows:
        groups[(r["компанія"], r["форма"], r["рік"], r["код"], r["графа"])].append(r)
    row = 5
    for key, group in sorted(groups.items()):
        best = next(r for r in group if r["обрано"])
        for other in group:
            if other is best or number(other["значення"]) == number(best["значення"]):
                continue
            company, form, year, code, column = key
            values = [company, form, int(year), code, names.get((form, code), ""), column,
                      number(best["значення"]), source_label(best),
                      number(other["значення"]), source_label(other),
                      number(best["значення"]) - number(other["значення"])]
            for col, v in enumerate(values, start=1):
                c = ws.cell(row, col, v)
                c.border = BORDER
                if col in (7, 9, 11):
                    c.number_format = NUMBER_FORMATS["тис. грн"]
            row += 1
    if row == 5:
        ws.cell(5, 1, "Розбіжностей немає.")
    set_widths(ws, [13, 7, 6, 7, 45, 7, 14, 42, 14, 42, 12])


def main():
    chosen, all_rows = load_values()
    names, notes = load_reference()
    indicators = read_csv(REFERENCE_DIR / "indicators.csv")
    templates = [i["формула"] for i in indicators] + [t for rows in EXTRA_ROWS.values() for _, t in rows]
    always_show = codes_used_in_formulas(templates)
    companies = list(COMPANIES.values())

    wb = new_workbook("Огляд")
    for company in companies:
        for form in ("001", "002", "003"):
            write_form_sheet(wb, company, form, chosen, names, notes, always_show)
        write_capital_sheet(wb, company, chosen, names)
    write_indicators_sheet(wb, indicators, companies)   # після аркушів форм: формулам потрібні номери рядків
    write_opendatabot_sheet(wb, companies,
                            "Переписано зі скріншотів raw_data/opendatabot/фінпоказники_*.png. Суми — у гривнях "
                            "(не тис.). Дохід, чистий прибуток і активи використано в аркушах форм лише за роки без "
                            "форм звітності (жовті клітинки); де є і форма, і сайт — порівняння на аркуші «Розбіжності».")
    write_differences_sheet(wb, all_rows, names)
    write_sources_sheet(wb, [(s["що"], s["джерело"]) for s in read_csv(REFERENCE_DIR / "sources.csv")],
                        {r["id звіту"] for r in all_rows if r["обрано"] and r["id звіту"]})

    overview = wb["Огляд"]
    write_title(overview, "Фінансова звітність групи ОККО, 2021–2025 (тис. грн)")
    row = write_availability(overview, 3, chosen, companies, SHEET_SUFFIX)
    write_text_rows(overview, row, [(t["тема"], t["текст"]) for t in read_csv(REFERENCE_DIR / "overview.csv")],
                    last_col=1 + len(YEARS))
    set_widths(overview, [36] + [21] * len(YEARS))

    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)
    print(f"Крок 3 готово: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
