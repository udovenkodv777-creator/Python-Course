import os
import time
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import xlwings as xw

# === НАСТРОЙКИ ИСХОДНЫХ ДАННЫХ И КООРДИНАТ ===
FILE_NAME = r"C:\00.Python\Работа с Excel\SMS.xlsx"
SHEET_SRC = "КП_28_2"
SHEET_NPV = "ТЭП"
NPV_CELL = "D60"

COL_START = "F"
COL_END = "AB"

# СТРОКИ ДЛЯ ВВОДА ДАННЫХ (Лист: КП_28_2)
ROW_FE_TARGET = 240    # Строка для Fe общ (числа вида 10.55)
ROW_P2O5_TARGET = 243  # Строка для P2O5 (числа вида 4.14)

# ДИАПАЗОНЫ ОТКЛОНЕНИЙ: Применяем шаги напрямую как числа (БЕЗ ДЕЛЕНИЯ НА 100)
FE_SHIFTS = np.array([-4, -3, -2, -1, 0, 1, 2, 3, 4]) * 0.50   # от -2.0 до +2.0
P2O5_SHIFTS = np.array([-4, -3, -2, -1, 0, 1, 2, 3, 4]) * 0.20 # от -0.8 до +0.8


def run_sensitivity_analysis():
    if not os.path.exists(FILE_NAME):
        print(f"Критическая ошибка: Файл не найден по пути: {FILE_NAME}")
        return

    print("Запуск Excel и открытие финансовой модели...")
    app = xw.App(visible=True, add_book=False)
    
    try:
        wb = app.books.open(FILE_NAME)
        app.calculation = 'automatic'
        
        ws_src = wb.sheets[SHEET_SRC]
        ws_npv = wb.sheets[SHEET_NPV]

        # 1. Сканируем и кэшируем базовый профиль содержаний
        print(f"Сканирование базовых профилей содержаний на листе '{SHEET_SRC}'...")
        base_fe_values = ws_src.range(f"{COL_START}{ROW_FE_TARGET}:{COL_END}{ROW_FE_TARGET}").value
        base_p2o5_values = ws_src.range(f"{COL_START}{ROW_P2O5_TARGET}:{COL_END}{ROW_P2O5_TARGET}").value

        # Подготовка пустой матрицы результатов под DataFrame
        fe_labels = [f"{shift:+.2f}%" if shift != 0 else "0.00%" for shift in FE_SHIFTS]
        p2o5_labels = [f"{shift:+.2f}%" if shift != 0 else "0.00%" for shift in P2O5_SHIFTS]
        df_results = pd.DataFrame(index=p2o5_labels, columns=fe_labels, dtype=float)

        print("\nЗапуск итерационного анализа чувствительности...")
        for i, fe_shift in enumerate(FE_SHIFTS):
            for j, p2o5_shift in enumerate(P2O5_SHIFTS):

                # Прямое сложение чисел без масштабирования
                new_fe = [x + fe_shift if x is not None else None for x in base_fe_values]
                new_p2o5 = [x + p2o5_shift if x is not None else None for x in base_p2o5_values]

                # Запись в Excel
                ws_src.range(f"{COL_START}{ROW_FE_TARGET}").value = new_fe
                ws_src.range(f"{COL_START}{ROW_P2O5_TARGET}").value = new_p2o5

                # Полный глубокий пересчет книги Excel (Ctrl+Alt+F9)
                wb.app.calculate()
                if app.api:
                    app.api.CalculateFull()
                
                time.sleep(0.5)

                # Считываем итоговое значение NPV
                current_npv = ws_npv.range(NPV_CELL).value
                df_results.iloc[j, i] = current_npv if current_npv is not None else 0.0

                print(f"Смещение [Fe: {fe_shift:+.2f}%, P2O5: {p2o5_shift:+.2f}%] -> NPV: {current_npv:,.2f} млн USD")

        # 2. Восстановление исходной структуры данных перед выходом
        print(f"\nВозврат параметров листа '{SHEET_SRC}' к исходному базовому плану...")
        ws_src.range(f"{COL_START}{ROW_FE_TARGET}").value = base_fe_values
        ws_src.range(f"{COL_START}{ROW_P2O5_TARGET}").value = base_p2o5_values
        
        if app.api:
            app.api.CalculateFull()
        wb.save()

        # Экспорт матрицы расчетов в CSV
        csv_filename = "SMS_NPV_Sensitivity_Matrix.csv"
        df_results.to_csv(csv_filename, encoding="utf-8-sig")
        print(f"Таблица результатов успешно экспортирована: {csv_filename}")

        # 3. Визуализация расширенной тепловой карты (Heatmap)
        print("Отрисовка и сохранение аналитической тепловой карты...")
        plt.figure(figsize=(13, 11))
        
        ax = sns.heatmap(
            df_results,
            annot=True,
            fmt=".2f",
            cmap="RdYlGn",
            linewidths=0.5,
            cbar_kws={"label": "NPV проекта СМС (млн USD)"}
        )

        # Подсвечиваем БАЗИС (0.00% изменений) жирным контуром
        base_fe_idx = list(FE_SHIFTS).index(0.0)
        base_p2o5_idx = list(P2O5_SHIFTS).index(0.0)
        rect = plt.Rectangle((base_fe_idx, base_p2o5_idx), 1, 1, fill=False, edgecolor="black", lw=3)
        ax.add_patch(rect)
        ax.text(base_fe_idx + 0.5, base_p2o5_idx + 0.3, "БАЗИС", color="black", ha="center", va="center", fontweight="bold", fontsize=9)

        plt.title("Матрица чувствительности NPV проекта СМС (млн USD)\nк изменению содержаний полезных компонентов (Вход через лист КП_28_2)", fontsize=12, fontweight="bold", pad=15)
        plt.xlabel("Абсолютное изменение базового содержания Fe общ (%)", fontsize=10)
        plt.ylabel("Абсолютное изменение базового содержания P2O5 (%)", fontsize=10)
        
        plt.gca().invert_yaxis()
        plt.tight_layout()
        
        image_filename = "SMS_NPV_Sensitivity_Heatmap_Extended.png"
        plt.savefig(image_filename, dpi=300)
        plt.close()
        print(f"Графический отчет успешно сохранен: {image_filename}")

    except Exception as e:
        print(f"\nКритическая ошибка при выполнении скрипта: {e}")

    finally:
        if 'wb' in locals():
            wb.close()
        app.quit()
        print("\nРабота с Excel-сессией полностью завершена.")


if __name__ == "__main__":
    run_sensitivity_analysis()