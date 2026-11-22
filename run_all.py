"""Запускає весь пайплайн: raw_data → data/1_extracted → data/2_values.csv → output/ (повна таблиця + курсова)"""
import step1_extract
import step2_values
import step3_excel
import step4_kursova

if __name__ == "__main__":
    print("── Крок 1: витягуємо звіти ОККО з сирих даних")
    step1_extract.main()
    print("\n── Крок 2: зводимо всі числа в одну таблицю")
    step2_values.main()
    print("\n── Крок 3: будуємо повну таблицю (обидві компанії)")
    step3_excel.main()
    print("\n── Крок 4: будуємо таблицю для курсової (ОККО-ДРАЙВ)")
    step4_kursova.main()
