import os
import time
from datetime import datetime
import xlwings as xw

# ==============================================================================
# БЛОК 1: НАСТРОЙКИ И ИСХОДНЫЙ РАСЧЕТ СЦЕНАРИЕВ В ТЕРМИНАЛ
# ==============================================================================
FILE_PATH = r"C:\00.Python\Работа с Excel\LOM_KGOK.xlsx"
INPUT_SHEET = "Macro&Data"
OUTPUT_SHEET = "Valuation"

ROW_USD_RATE = "L21:AA21"   
ROW_INFLATION = "L12:AA12"  
CELL_PROJECT_NPV = "F3"

scenarios_block1 = {
    "Базовый (План)": {
        "usd": [86.0, 78.0, 85.0, 93.0, 101.0, 107.0, 111.0, 115.0, 120.0, 124.0, 129.0, 134.0, 140.0, 145.0, 151.0, 158.0],
        "inf": [1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04, 1.04]
    },
    "Кризис (Крепкий рубль + Инфляция)": {
        "usd": [82.0, 72.0, 82.0, 90.0, 100.0, 101.0, 103.0, 105.0, 107.0, 110.0, 114.0, 120.0, 125.0, 130.0, 135.0, 140.0],
        "inf": [1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06, 1.06]
    },
    "Девальвационный (Рост курса)": {
        "usd": [96.0, 88.0, 95.0, 103.0, 111.0, 117.0, 121.0, 125.0, 130.0, 144.0, 159.0, 164.0, 170.0, 185.0, 191.0, 198.0],
        "inf": [1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05]
    }
}

if not os.path.exists(FILE_PATH):
    raise FileNotFoundError(f"Файл не найден: {FILE_PATH}")

print("Запуск Excel для Блока 1...")
app = xw.App(visible=True)

try:
    wb = xw.Book(FILE_PATH)
    in_sh = wb.sheets[INPUT_SHEET]
    out_sh = wb.sheets[OUTPUT_SHEET]
    
    print("\n=== БЛОК 1: Расчет погодовых макросценариев ===")
    
    for name, data in scenarios_block1.items():
        in_sh.range(ROW_USD_RATE).value = data["usd"]
        in_sh.range(ROW_INFLATION).value = data["inf"]
        
        time.sleep(0.7)
        current_npv = out_sh.range(CELL_PROJECT_NPV).value
        
        print(f"\n[+] Сценарий: {name}")
        print(f"    Курс USD по годам: {data['usd']}")
        print(f"    Итоговый NPV: ${current_npv:,.0f}")

finally:
    print("Закрытие Excel для Блока 1...")
    if 'wb' in locals():
        wb.close()
    if 'app' in locals():
        app.quit()


# ==============================================================================
# БЛОК 2: РАСЧЕТ И АВТОМАТИЧЕСКОЕ СОХРАНЕНИЕ ОТЧЕТА В ОТДЕЛЬНЫЙ EXCEL-ФАЙЛ
# ==============================================================================
#%%
print("\n[ЗАПУСК] Выполняется Блок 2 (Расчет + Запись отчета)...")

FILE_PATH_B2 = r"C:\00.Python\Работа с Excel\LOM_KGOK.xlsx"
INPUT_SHEET_B2 = "Macro&Data"
OUTPUT_SHEET_B2 = "Valuation"
REPORT_PATH = r"C:\00.Python\Работа с Excel\Сценарии_Отчет.xlsx"

ROW_USD_RATE_B2 = "L21:AA21"   
ROW_INFLATION_B2 = "L12:AA12"  
CELL_PROJECT_NPV_B2 = "F3"

scenarios_block2 = {
    "Базовый (План)": {
        "usd": [86.0, 78.0, 85.0, 93.0, 101.0, 107.0, 111.0, 115.0, 120.0, 124.0, 129.0, 134.0, 140.0, 145.0, 151.0, 158.0],
        "inf": [1.04] * 16
    },
    "Кризис (Крепкий рубль + Инфляция)": {
        "usd": [82.0, 72.0, 82.0, 90.0, 100.0, 101.0, 103.0, 105.0, 107.0, 110.0, 114.0, 120.0, 125.0, 130.0, 135.0, 140.0],
        "inf": [1.06] * 16
    },
    "Девальвационный (Рост курса)": {
        "usd": [96.0, 88.0, 95.0, 103.0, 111.0, 117.0, 121.0, 125.0, 130.0, 144.0, 159.0, 164.0, 170.0, 185.0, 191.0, 198.0],
        "inf": [1.05] * 16
    }
}

print("Запуск нового процесса Excel для Блока 2...")
app2 = xw.App(visible=True)
report_data = []

try:
    wb2 = xw.Book(FILE_PATH_B2)
    in_sh2 = wb2.sheets[INPUT_SHEET_B2]
    out_sh2 = wb2.sheets[OUTPUT_SHEET_B2]
    
    print("\n=== БЛОК 2: Пересчет модели и сбор данных ===")
    
    for name, data in scenarios_block2.items():
        in_sh2.range(ROW_USD_RATE_B2).value = data["usd"]
        in_sh2.range(ROW_INFLATION_B2).value = data["inf"]
        
        time.sleep(0.7)
        current_npv = out_sh2.range(CELL_PROJECT_NPV_B2).value
        
        print(f"[+] Сценарий: {name} --> Итоговый NPV: ${current_npv:,.0f}")
        
        # Сохраняем результаты в массив
        report_data.append([
            name, 
            current_npv, 
            round(sum(data["usd"]) / len(data["usd"]), 2), 
            datetime.now().strftime("%Y-%m-%d %H:%M")
        ])

finally:
    print("Закрытие фонового файла LOM_KGOK...")
    if 'wb2' in locals():
        wb2.close()
    if 'app2' in locals():
        app2.quit()

# ==============================================================================
# НАДЕЖНАЯ ЗАПИСЬ ОТЧЕТА НАПРЯМУЮ ЧЕРЕЗ XLWINGS
# ==============================================================================
print("\nСоздание нового итогового файла отчета...")
app_rep = xw.App(visible=False) # Скрытно, чтобы не моргало на экране

try:
    wb_rep = xw.Book() # Создаем абсолютно новый пустой файл Excel
    sh_rep = wb_rep.sheets[0]
    sh_rep.name = "Результаты"
    
    # 1. Записываем шапку таблицы
    sh_rep.range("A1").value = ["Название сценария", "Итоговый NPV ($)", "Средний курс USD", "Дата расчета"]
    
    # 2. Мгновенно выгружаем весь массив собранных данных, начиная со второй строки
    sh_rep.range("A2").value = report_data
    
    # 3. Красивое выравнивание колонок по ширине текста
    sh_rep.autofit()
    
    # 4. Сохраняем в вашу рабочую папку
    wb_rep.save(REPORT_PATH)
    print(f"\n[УСПЕХ] Отчет УСПЕШНО сохранен по пути: {REPORT_PATH}")
    
finally:
    if 'wb_rep' in locals():
        wb_rep.close()
    if 'app_rep' in locals():
        app_rep.quit()