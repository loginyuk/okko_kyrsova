"""
Крок 4. Таблиця для курсової «Інноваційні бізнес-моделі як чинник стратегічного розвитку
підприємства (на прикладі ТОВ «ОККО-ДРАЙВ»)»: output/OKKO_kursova_2021-2025.xlsx

Об'єкт аналізу — лише ТОВ «ОККО-ДРАЙВ»; АТ «Концерн Галнафтогаз» — тільки довідка.
Будується з тих самих даних, що й крок 3 (data/2_values.csv), але від кроку 3 не залежить:
повну таблицю можна прибрати, і цей крок працюватиме далі.

Аркуші:
  Огляд              наявність даних ОККО-ДРАЙВ, опис компанії, особливості, позначення
  ОД_Баланс          Ф1 повністю + примітка про розбіжність із сайтом Опендатабота
  ОД_Фінрез          Ф2 повністю
  ОД_Рух_коштів      Ф3 — лише ключові рядки (KEY_CASH_FLOW_CODES)
  Показники          показники з позначкою «так» у колонці «у курсовій» (reference/indicators.csv)
  Опендатабот        таблиця з сайту — лише ОККО-ДРАЙВ
  Довідка_ГНГ        кілька цифр про Галнафтогаз (числа, не формули)
  Джерела            тексти — reference/kursova_texts.csv

Запуск:  .venv/bin/python step4_kursova.py
"""
import re

from common import KURSOVA_XLSX, REFERENCE_DIR, YEARS, read_csv
from excel_common import (BLUE, BORDER, EXTRA_ROWS, GREY, ITALIC, NUMBER_FORMATS, RIGHT, YELLOW, codes_used_in_formulas,
                          load_opendatabot_summary, load_reference, load_values, new_workbook, number, put_value,
                          set_widths, write_availability, write_form_sheet, write_header, write_indicators_sheet,
                          write_opendatabot_sheet, write_sources_sheet, write_text_rows, write_title)

COMPANY = "ОККО-ДРАЙВ"
FORMS = ["001", "002", "003"]
KEY_CASH_FLOW_CODES = ["3000", "3195", "3260", "3355"]   # надходження від реалізації, ОГП, капінвестиції, дивіденди

# Назви, якими сайт Опендатабота називає відповідні рядки форм (для приміток про розбіжності)
SITE_NAMES = {"1300": "Активи", "2000": "Чистий дохід", "2350": "Чистий прибуток"}

# Довідка про Галнафтогаз: (назва рядка, звідки брати, джерело для підпису)
GNG = "Галнафтогаз"
GNG_YEARS = [2024, 2025]
GNG_ROWS = [
    ("Чистий дохід ГНГ (ф. 2, р. 2000), тис. грн", ("форма", "002", "2000"),
     "Держстат, ф. 2 за 2025 р. (2024 — порівняльна колонка)"),
    ("Надходження від операційної оренди (ф. 3, р. 3040), тис. грн", ("форма", "003", "3040"),
     "Держстат, ф. 3 за 2025 р. (2024 — порівняльна колонка)"),
    ("Капітальні інвестиції ГНГ (ф. 3, р. 3260), тис. грн", ("форма", "003", "3260"),
     "Держстат, ф. 3 за 2025 р. (2024 — порівняльна колонка)"),
    ("Кількість працівників ГНГ, осіб", ("сайт", "Кількість працівників"),
     "Опендатабот, «Фінансові показники» (raw_data/opendatabot/фінпоказники_ГНГ_31729918.png)"),
    ("АЗС у власності ГНГ на кінець 2025 р., шт.", ("число", {2025: 341}),
     "Звіт про управління АТ «Концерн Галнафтогаз» за 2025 р. (звітність за МСФЗ, portal.frs.gov.ua)"),
]


def site_discrepancy_notes(all_rows, form):
    """Тексти приміток: де сайт Опендатабота дає інше число, ніж форма звітності."""
    notes = []
    for best in all_rows:
        if not (best["компанія"] == COMPANY and best["форма"] == form and best["обрано"]
                and best["код"] in SITE_NAMES):
            continue
        for other in all_rows:
            if (other["тип"] == "агрегатор" and other["компанія"] == COMPANY and other["форма"] == form
                    and other["рік"] == best["рік"] and other["код"] == best["код"]
                    and number(other["значення"]) != number(best["значення"])):
                form_no = int(form)
                notes.append(f"{SITE_NAMES[best['код']]} {best['рік']} р. за формою {form_no} — "
                             f"{number(best['значення']):,} тис. грн; на сайті Опендатабота — "
                             f"{number(other['значення']):,} тис. грн. У роботі використано дані форми {form_no}."
                             .replace(",", " "))
    return notes


def write_gng_sheet(wb, chosen, note):
    ws = wb.create_sheet("Довідка_ГНГ")
    write_title(ws, "Довідка: АТ «Концерн Галнафтогаз» — власник активів групи (не об'єкт аналізу)",
                "Значення перенесено як числа. Сині — зі звітності, жовта заливка — з сайту Опендатабота.")
    write_header(ws, 4, ["Показник", *map(str, GNG_YEARS)])
    site_tables, _ = load_opendatabot_summary()

    row = 5
    for title, (kind, *where), _ in GNG_ROWS:
        ws.cell(row, 1, title).border = BORDER
        for i, year in enumerate(GNG_YEARS):
            c = ws.cell(row, 2 + i)
            if kind == "форма":
                form, code = where
                put_value(c, chosen.get((GNG, form, code, ""), {}).get(year))
                continue
            if kind == "сайт":
                values = next(v for (name, _), v in site_tables[GNG].items() if name == where[0])
                c.value, c.fill = values.get(str(year)), YELLOW
            else:                                   # число, внесене вручну з тексту звіту
                c.value = where[0].get(year)
            c.font, c.border, c.number_format = BLUE, BORDER, NUMBER_FORMATS["тис. грн"]
            if c.value is None:
                c.value, c.font, c.alignment = "—", GREY, RIGHT
        row += 1

    row += 1
    ws.cell(row, 1, "Джерела рядків:").font = ITALIC
    row += 1
    for title, _, source in GNG_ROWS:
        short_title = re.sub(r", (тис\. грн|осіб|шт\.)$", "", title)     # без одиниці виміру
        ws.cell(row, 1, f"• {short_title} — {source}")
        row += 1
    row += 1
    ws.cell(row, 1, f"Примітка: {note}").font = ITALIC
    set_widths(ws, [62, 16, 16])


def main():
    chosen, all_rows = load_values()
    names, notes = load_reference()
    indicators = [i for i in read_csv(REFERENCE_DIR / "indicators.csv") if i["у курсовій"] == "так"]
    texts = read_csv(REFERENCE_DIR / "kursova_texts.csv")
    templates = [i["формула"] for i in indicators] + [t for f in ("001", "002") for _, t in EXTRA_ROWS[f]]
    always_show = codes_used_in_formulas(templates)

    wb = new_workbook("Огляд")
    ws, row = write_form_sheet(wb, COMPANY, "001", chosen, names, notes, always_show)
    for text in site_discrepancy_notes(all_rows, "001"):
        ws.cell(row + 1, 1, f"Примітка: {text}").font = ITALIC
        row += 1
    write_form_sheet(wb, COMPANY, "002", chosen, names, notes, always_show)
    write_form_sheet(wb, COMPANY, "003", chosen, names, notes, always_show,
                     codes=KEY_CASH_FLOW_CODES, extra_rows=False)
    write_indicators_sheet(wb, indicators, [COMPANY])   # після аркушів форм: формулам потрібні номери рядків
    write_opendatabot_sheet(wb, [COMPANY],
                            "Переписано зі скріншота raw_data/opendatabot/фінпоказники_ОККО-ДРАЙВ_20122395.png. "
                            "Суми — у гривнях (не тис.). Для аналізу використано форми звітності; сайт — довідково.")
    gng_note = next(t["текст"] for t in texts if t["аркуш"] == "Довідка_ГНГ")
    write_gng_sheet(wb, chosen, gng_note)
    write_sources_sheet(wb, [(t["тема"], t["текст"]) for t in texts if t["аркуш"] == "Джерела"],
                        {r["id звіту"] for r in all_rows
                         if r["обрано"] and r["компанія"] == COMPANY and r["форма"] in FORMS and r["id звіту"]})

    overview = wb["Огляд"]
    write_title(overview, "ТОВ «ОККО-ДРАЙВ»: фінансова звітність 2021–2025 для курсової роботи (тис. грн)",
                "«Інноваційні бізнес-моделі як чинник стратегічного розвитку підприємства "
                "(на прикладі ТОВ «ОККО-ДРАЙВ»)»")
    row = write_availability(overview, 4, chosen, [COMPANY], FORMS)
    write_text_rows(overview, row, [(t["тема"], t["текст"]) for t in texts if t["аркуш"] == "Огляд"],
                    last_col=1 + len(YEARS))
    set_widths(overview, [36] + [21] * len(YEARS))

    KURSOVA_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(KURSOVA_XLSX)
    print(f"Крок 4 готово: {KURSOVA_XLSX}")


if __name__ == "__main__":
    main()
