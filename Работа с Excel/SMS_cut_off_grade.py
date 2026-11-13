import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.chart import LineChart, Reference

def create_cut_off_fem(filename="Cut_Off_Grade_MAR_FEM.xlsx"):
    wb = openpyxl.Workbook()
    
    ws_in = wb.active
    ws_in.title = "Inputs"
    ws_calc = wb.create_sheet(title="Unit_Economics_Calc")
    ws_mat = wb.create_sheet(title="CutOff_Matrices_Visual")

    font_family = "Calibri"
    
    f_title = Font(name=font_family, size=14, bold=True, color="1F4E78")
    f_sec_hdr = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    f_tbl_hdr = Font(name=font_family, size=10, bold=True, color="1F4E78")
    f_input = Font(name=font_family, size=10, bold=False, color="0000FF")
    f_formula = Font(name=font_family, size=10, bold=False, color="000000")
    f_link = Font(name=font_family, size=10, bold=False, color="008000")
    f_total = Font(name=font_family, size=10, bold=True, color="000000")
    f_note = Font(name=font_family, size=9, italic=True, color="595959")

    fill_sec_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_tbl_hdr = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    fill_waste = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    fill_of = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_sms = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    thin_side = Side(border_style="thin", color="BFBFBF")
    double_side = Side(border_style="double", color="000000")
    top_thin = Side(border_style="thin", color="000000")
    
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(top=top_thin, bottom=double_side, left=thin_side, right=thin_side)

    fmt_int = "#,##0"
    fmt_dec2 = "#,##0.00"
    fmt_pct = "0.00%"
    fmt_rub = '#,##0.00" ₽"'
    fmt_usd = '$#,##0.00'

    ws_in.views.sheetView[0].showGridLines = True
    ws_in["B2"] = "ФЭМ РАСЧЕТА CUT-OFF GRADE (МАР): ИСХОДНЫЕ ДАННЫЕ"
    ws_in["B2"].font = f_title

    ws_in["F2"] = "ЛЕГЕНДА ФОРМАТИРОВАНИЯ (СТАНДАРТ BIG4 / McKinsey)"
    ws_in["F2"].font = f_tbl_hdr
    
    legend_items = [
        ("F3", "Ручной ввод (Inputs)", f_input, fill_input, "Изменяемые параметры сценария"),
        ("F4", "Формула на текущем листе", f_formula, PatternFill(fill_type=None), "Внутренние расчеты"),
        ("F5", "Ссылка на другой лист", f_link, PatternFill(fill_type=None), "Импорт данных из Inputs/Calc"),
        ("F6", "Итоговые / ключевые строки", f_total, fill_total, "Результирующие показатели")
    ]
    for cell_ref, text, font_style, fill_style, desc in legend_items:
        ws_in[cell_ref] = text
        ws_in[cell_ref].font = font_style
        if fill_style.fill_type:
            ws_in[cell_ref].fill = fill_style
        ws_in[cell_ref].border = border_data
        
        desc_cell = ws_in[f"G{cell_ref[1:]}"]
        desc_cell.value = desc
        desc_cell.font = f_note

    inputs_data = [
        ("1. ОБЩИЕ ПАРАМЕТРЫ И БАЗОВАЯ ХИМИЯ БЛОКА", None, None, None, None, None, None),
        (None, "Базовый объем анализируемого блока (сырая масса)", 1.0, "т", fmt_dec2, "Расчет на 1 т извлекаемой горной массы", "VOL_BASE"),
        (None, "Содержание влаги в исходной руде", 0.035, "%", fmt_pct, "Пересчет сырой массы в сухую", "MOISTURE"),
        (None, "Базовое содержание Fe (общее / из FeO) в руде", 0.140, "%", fmt_pct, "Для точечного расчета на листе Calc", "FE_BASE"),
        (None, "Коэффициент пересчета FeO в Fe (если ввод в FeO)", 1.000, "доля", fmt_dec2, "1.00 если ввод сразу в Fe; 0.777 если ввод в FeO", "FEO_CONV"),
        (None, "Базовое содержание P2O5 в руде", 0.035, "%", fmt_pct, "Для точечного расчета на листе Calc", "P2O5_BASE"),
        (None, "Содержание ZrO2 в руде (попутный компонент)", 0.0015, "%", fmt_pct, "Учитывается при переработке на ОФ (БК)", "ZRO2_BASE"),
        (None, "Разубоживание при выемке в карьере", 0.050, "%", fmt_pct, "Снижение содержаний при подаче из карьера", "DILUTION"),

        ("2. МАКРОЭКОНОМИКА И ЦЕНЫ РЕАЛИЗАЦИИ (FCA/FOB)", None, None, None, None, None, None),
        (None, "Курс доллара США", 92.50, "руб./$", fmt_dec2, "Курс пересчета валютной выручки", "FX_RATE"),
        (None, "Цена ЖРК (железорудный концентрат)", 115.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_FE"),
        (None, "Цена АПК (апатитовый концентрат)", 165.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_P"),
        (None, "Цена БК (бадделеитовый концентрат)", 3200.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_ZR"),

        ("3. КАЧЕСТВО КОНЦЕНТРАТОВ И СКВОЗНОЕ ИЗВЛЕЧЕНИЕ (ОФ)", None, None, None, None, None, None),
        (None, "Содержание Fe в товарном ЖРК", 0.655, "%", fmt_pct, "Базисное качество ЖРК", "GRADE_FE_CONC"),
        (None, "Содержание P2O5 в товарном АПК", 0.390, "%", fmt_pct, "Базисное качество АПК", "GRADE_P_CONC"),
        (None, "Содержание ZrO2 в товарном БК", 0.985, "%", fmt_pct, "Базисное качество БК", "GRADE_ZR_CONC"),
        (None, "Сквозное извлечение Fe в ЖРК (на ОФ)", 0.780, "%", fmt_pct, "От питания ОФ в товарный концентрат", "REC_FE"),
        (None, "Сквозное извлечение P2O5 в АПК (на ОФ)", 0.740, "%", fmt_pct, "От питания ОФ в товарный концентрат", "REC_P"),
        (None, "Сквозное извлечение ZrO2 в БК (на ОФ)", 0.550, "%", fmt_pct, "От питания ОФ в товарный концентрат", "REC_ZR"),

        ("4. ПАРАМЕТРЫ РУДОРАЗБОРКИ НА ДСК-СМС (2 СТАДИИ)", None, None, None, None, None, None),
        (None, "Выход предконцентрата после 2-х стадий СМС", 0.7236, "%", fmt_pct, "По условию задачи: 72,36% идет на ОФ", "SMS_YIELD"),
        (None, "Извлечение Fe в предконцентрат СМС", 0.940, "%", fmt_pct, "Доля Fe, переходящая в предконцентрат", "SMS_REC_FE"),
        (None, "Извлечение P2O5 в предконцентрат СМС", 0.880, "%", fmt_pct, "Доля апатита, уходящая с предконцентратом на ОФ", "SMS_REC_P"),
        (None, "Извлечение ZrO2 в предконцентрат СМС", 0.850, "%", fmt_pct, "Доля бадделеита, уходящая на ОФ", "SMS_REC_ZR"),
        (None, "Опция хвостов СМС (1 = в отвал; 2 = щебень для дорог)", 2, "код", fmt_int, "1 - затраты на вывоз; 2 - вывоз на штабель + экономия на щебне", "SMS_TAIL_OPT"),
        (None, "Экономия / ценность щебня из хвостов СМС", 350.00, "руб./т", fmt_rub, "Альтернативная стоимость покупки щебня для дорог карьера", "GRAVEL_VAL"),

        ("5. ЛОГИСТИКА: ПЛЕЧИ ТРАНСПОРТИРОВКИ", None, None, None, None, None, None),
        (None, "Плечо 1: Карьер -> Отвал вскрыши (Сценарий 1)", 3.20, "км", fmt_dec2, "Вывоз пустой породы", "DIST_WASTE"),
        (None, "Плечо 2: Карьер -> ОФ напрямую (Сценарий 2)", 6.50, "км", fmt_dec2, "Прямая доставка руды на фабрику", "DIST_OF"),
        (None, "Плечо 3а: Карьер -> ДСК-СМС (Сценарий 3)", 2.10, "км", fmt_dec2, "Доставка горной массы на корпус СМС", "DIST_TO_SMS"),
        (None, "Плечо 3б: ДСК-СМС -> ОФ (предконцентрат 72,36%)", 4.80, "км", fmt_dec2, "Доставка обогащенного продукта на ОФ", "DIST_SMS_OF"),
        (None, "Плечо 3в (опция 1): ДСК-СМС -> Отвал хвостов", 1.80, "км", fmt_dec2, "Если хвосты СМС не используются как щебень", "DIST_SMS_WASTE"),
        (None, "Плечо 3г (опция 2): ДСК-СМС -> Штабель щебня", 0.90, "км", fmt_dec2, "Доставка хвостов СМС на склад дорожного щебня", "DIST_SMS_GRAVEL"),

        ("6. УДЕЛЬНЫЕ ОПЕРАЦИОННЫЕ РАСХОДЫ (УДР / OPEX)", None, None, None, None, None, None),
        (None, "УДР транспортировки горной массы / продуктов", 28.50, "руб./т*км", fmt_rub, "Карьерный автотранспорт / конвейер", "OPEX_TR"),
        (None, "УДР дробления на ОФ (стандартная схема)", 110.00, "руб./т", fmt_rub, "Крупное, среднее и мелкое дробление на ОФ", "OPEX_CRUSH_OF"),
        (None, "УДР передела ДСК-СМС (2 стадии дробления + СМС)", 145.00, "руб./т", fmt_rub, "Затраты на 1 т исходного питания ДСК-СМС", "OPEX_SMS"),
        (None, "УДР переработки ЖРК (измельчение -> концентрат)", 420.00, "руб./т руды", fmt_rub, "На 1 т сухой руды, поступившей на ОФ", "OPEX_OF_FE"),
        (None, "УДР переработки АПК (флотация -> концентрат)", 290.00, "руб./т руды", fmt_rub, "На 1 т сухой руды, поступившей на ОФ", "OPEX_OF_P"),
        (None, "УДР переработки БК (гравитация/доводка -> конц.)", 85.00, "руб./т руды", fmt_rub, "На 1 т сухой руды, поступившей на ОФ", "OPEX_OF_ZR"),
        (None, "НДПИ и БВР/Экскавация (справочно, не в Cut-off)", 310.00, "руб./т", fmt_rub, "Sunk costs: одинаковы для всех 3-х сценариев в контуре карьера", "SUNK_COST")
    ]

    row_idx = 8
    cell_map = {}

    headers = ["Параметр / Статья", "Значение", "Ед. изм.", "Методическое примечание"]
    for col_offset, h_text in enumerate(headers):
        c = ws_in.cell(row=row_idx, column=2 + col_offset, value=h_text)
        c.font = f_tbl_hdr
        c.fill = fill_tbl_hdr
        c.border = border_data
        align_dir = "center" if col_offset == 1 or col_offset == 2 else "left"
        c.alignment = Alignment(horizontal=align_dir, vertical="center")
    row_idx += 1
    for cat, param, val, unit, num_fmt, note, code in inputs_data:
        if cat:
            ws_in.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=5)
            c = ws_in.cell(row=row_idx, column=2, value=cat)
            c.font = f_sec_hdr
            c.fill = fill_sec_hdr
            c.alignment = Alignment(horizontal="left", vertical="center")
            ws_in.row_dimensions[row_idx].height = 22
        else:
            c_param = ws_in.cell(row=row_idx, column=2, value=param)
            c_param.font = f_formula
            c_param.border = border_data

            c_val = ws_in.cell(row=row_idx, column=3, value=val)
            c_val.font = f_input
            c_val.fill = fill_input
            c_val.border = border_data
            c_val.number_format = num_fmt
            c_val.alignment = Alignment(horizontal="right")

            c_unit = ws_in.cell(row=row_idx, column=4, value=unit)
            c_unit.font = f_note
            c_unit.border = border_data
            c_unit.alignment = Alignment(horizontal="center")

            c_note = ws_in.cell(row=row_idx, column=5, value=note)
            c_note.font = f_note
            c_note.border = border_data

            cell_map[code] = f"Inputs!$C${row_idx}"
        row_idx += 1

    ws_in.column_dimensions["A"].width = 3
    ws_in.column_dimensions["B"].width = 52
    ws_in.column_dimensions["C"].width = 16
    ws_in.column_dimensions["D"].width = 14
    ws_in.column_dimensions["E"].width = 55
    ws_in.column_dimensions["F"].width = 32
    ws_in.column_dimensions["G"].width = 35

    ws_calc.views.sheetView[0].showGridLines = True
    ws_calc["B2"] = "МОДЕЛЬ UNIT-ЭКОНОМИКИ 1 ТОННЫ ПОРОДЫ ПО ТРЕМ АЛГОРИТМАМ (СЦЕНАРИЯМ)"
    ws_calc["B2"].font = f_title

    ws_calc["B3"] = "Примечание: Затраты на БВР, экскавацию и НДПИ являются невозвратными (Sunk Costs) в рамках контура карьера и исключены из критерия принятия решений."
    ws_calc["B3"].font = f_note

    calc_headers = [
        "Показатель / Статья расчета",
        "Ед. изм.",
        "Сценарий 1: В отвал",
        "Сценарий 2: Прямая ОФ",
        "Сценарий 3: ДСК-СМС -> ОФ",
        "Формула / Логика расчета"
    ]
    r = 5
    for idx, text in enumerate(calc_headers):
        c = ws_calc.cell(row=r, column=2 + idx, value=text)
        c.font = f_tbl_hdr
        c.fill = fill_tbl_hdr
        c.border = border_data
        c.alignment = Alignment(horizontal="center" if idx > 0 else "left", wrap_text=True)
    ws_calc.row_dimensions[r].height = 28

    v_vol = cell_map["VOL_BASE"]
    v_moist = cell_map["MOISTURE"]
    v_fe = cell_map["FE_BASE"]
    v_fe_conv = cell_map["FEO_CONV"]
    v_p = cell_map["P2O5_BASE"]
    v_zr = cell_map["ZRO2_BASE"]
    v_dil = cell_map["DILUTION"]
    
    fx = cell_map["FX_RATE"]
    p_fe = cell_map["PRICE_FE"]
    p_p = cell_map["PRICE_P"]
    p_zr = cell_map["PRICE_ZR"]
    
    g_fe_c = cell_map["GRADE_FE_CONC"]
    g_p_c = cell_map["GRADE_P_CONC"]
    g_zr_c = cell_map["GRADE_ZR_CONC"]
    
    rec_fe = cell_map["REC_FE"]
    rec_p = cell_map["REC_P"]
    rec_zr = cell_map["REC_ZR"]
    
    sms_y = cell_map["SMS_YIELD"]
    sms_r_fe = cell_map["SMS_REC_FE"]
    sms_r_p = cell_map["SMS_REC_P"]
    sms_r_zr = cell_map["SMS_REC_ZR"]
    sms_opt = cell_map["SMS_TAIL_OPT"]
    gravel_v = cell_map["GRAVEL_VAL"]
    
    d_w = cell_map["DIST_WASTE"]
    d_of = cell_map["DIST_OF"]
    d_to_sms = cell_map["DIST_TO_SMS"]
    d_sms_of = cell_map["DIST_SMS_OF"]
    d_sms_w = cell_map["DIST_SMS_WASTE"]
    d_sms_g = cell_map["DIST_SMS_GRAVEL"]
    
    c_tr = cell_map["OPEX_TR"]
    c_cr_of = cell_map["OPEX_CRUSH_OF"]
    c_sms = cell_map["OPEX_SMS"]
    c_of_fe = cell_map["OPEX_OF_FE"]
    c_of_p = cell_map["OPEX_OF_P"]
    c_of_zr = cell_map["OPEX_OF_ZR"]

    calc_rows = [
        ("1. МАТЕРИАЛЬНЫЙ БАЛАНС И КАЧЕСТВО ПИТАНИЯ ОФ", None, None, None, None, None, True),
        ("Сырая горная масса из карьера", "т", f"={v_vol}", f"={v_vol}", f"={v_vol}", "Базовая 1 тонна из карьера", fmt_dec2, False),
        ("Сухая масса горной породы из карьера", "т", f"=D7*(1-{v_moist})", f"=E7*(1-{v_moist})", f"=F7*(1-{v_moist})", "Исключение влаги", fmt_dec2, False),
        ("Масса сухой руды, поступающей на ОФ", "т", 0, "=E8", f"=F8*{sms_y}", "В Сц.3 на ОФ идет только предконцентрат СМС", fmt_dec2, False),
        ("Масса хвостов СМС (выход щебня / отвала)", "т", 0, 0, f"=F8*(1-{sms_y})", "Отсекается после 2-х стадий СМС", fmt_dec2, False),
        ("Содержание Fe в питании ОФ (с учетом разубоживания)", "%", 0, f"={v_fe}*{v_fe_conv}*(1-{v_dil})", f"=(E11*{sms_r_fe})/{sms_y}", "Повышение качества за счет СМС", fmt_pct, False),
        ("Содержание P2O5 в питании ОФ (с учетом разубоживания)", "%", 0, f"={v_p}*(1-{v_dil})", f"=(E12*{sms_r_p})/{sms_y}", "Концентрация апатита в предконцентрате", fmt_pct, False),
        ("Содержание ZrO2 в питании ОФ (с учетом разубоживания)", "%", 0, f"={v_zr}*(1-{v_dil})", f"=(E13*{sms_r_zr})/{sms_y}", "Концентрация бадделеита", fmt_pct, False),

        ("2. ВЫПУСК ТОВАРНЫХ КОНЦЕНТРАТОВ (СУХОЙ ВЕС)", None, None, None, None, None, True),
        ("Выпуск ЖРК (железорудного концентрата)", "т", 0, f"=(E9*E11*{rec_fe})/{g_fe_c}", f"=(F9*F11*{rec_fe})/{g_fe_c}", "Масса_ОФ * Fe% * Извлечение / Качество_ЖРК", fmt_dec2, False),
        ("Выпуск АПК (апатитового концентрата)", "т", 0, f"=(E9*E12*{rec_p})/{g_p_c}", f"=(F9*F12*{rec_p})/{g_p_c}", "Масса_ОФ * P2O5% * Извлечение / Качество_АПК", fmt_dec2, False),
        ("Выпуск БК (бадделеитового концентрата)", "т", 0, f"=(E9*E13*{rec_zr})/{g_zr_c}", f"=(F9*F13*{rec_zr})/{g_zr_c}", "Масса_ОФ * ZrO2% * Извлечение / Качество_БК", fmt_dec2, False),

        ("3. ВЫРУЧКА ОТ РЕАЛИЗАЦИИ И ЭКОНОМИЯ (NSR)", None, None, None, None, None, True),
        ("Выручка от реализации ЖРК", "руб.", 0, f"=E15*{p_fe}*{fx}", f"=F15*{p_fe}*{fx}", "Выпуск * Цена($) * Курс", fmt_rub, False),
        ("Выручка от реализации АПК", "руб.", 0, f"=E16*{p_p}*{fx}", f"=F16*{p_p}*{fx}", "Выпуск * Цена($) * Курс", fmt_rub, False),
        ("Выручка от реализации БК", "руб.", 0, f"=E17*{p_zr}*{fx}", f"=F17*{p_zr}*{fx}", "Выпуск * Цена($) * Курс", fmt_rub, False),
        ("Экономия на покупке дорожного щебня (хвосты СМС)", "руб.", 0, 0, f"=IF({sms_opt}=2, F10/(1-{v_moist})*{gravel_v}, 0)", "Активно при выборе опции 2 в Inputs", fmt_rub, False),
        ("ИТОГО ДОХОД / ЦЕННОСТЬ БЛОКА (GROSS VALUE)", "руб.", "=SUM(D19:D22)", "=SUM(E19:E22)", "=SUM(F19:F22)", "Суммарная выручка и экономия", fmt_rub, True),

        ("4. ПЕРЕМЕННЫЕ ОПЕРАЦИОННЫЕ ЗАТРАТЫ (OPEX)", None, None, None, None, None, True),
        ("Транспортировка горной массы и продуктов", "руб.", f"=D7*{d_w}*{c_tr}", f"=E7*{d_of}*{c_tr}", f"=F7*{d_to_sms}*{c_tr} + (F9/(1-{v_moist}))*{d_sms_of}*{c_tr} + (F10/(1-{v_moist}))*IF({sms_opt}=1,{d_sms_w},{d_sms_g})*{c_tr}", "Учет всех плеч перевозки в сырой массе", fmt_rub, False),
        ("Дробление и сухая магнитная сепарация (ДСК-СМС)", "руб.", 0, f"=E8*{c_cr_of}", f"=F8*{c_sms}", "В Сц.2 дробление на ОФ; в Сц.3 2 стадии на ДСК-СМС", fmt_rub, False),
        ("Обогащение на ОФ (измельчение, флотация, доводка)", "руб.", 0, f"=E9*({c_of_fe}+{c_of_p}+{c_of_zr})", f"=F9*({c_of_fe}+{c_of_p}+{c_of_zr})", "Начисляется только на массу, поступившую на ОФ", fmt_rub, False),
        ("ИТОГО ПЕРЕМЕННЫЕ ЗАТРАТЫ (OPEX)", "руб.", "=SUM(D25:D27)", "=SUM(E25:E27)", "=SUM(F25:F27)", "Суммарные релевантные затраты", fmt_rub, True),

        ("5. ЭКОНОМИЧЕСКИЙ ЭФФЕКТ И ВЫБОР АЛГОРИТМА", None, None, None, None, None, True),
        ("Маржинальная прибыль (Net Value = Доход - OPEX)", "руб./т", "=D23-D28", "=E23-E28", "=F23-F28", "Критерий максимизации стоимости", fmt_rub, True),
        ("Прирост прибыли относительно Отвала (Сц.1)", "руб./т", 0, "=E30-$D$30", "=F30-$D$30", "Если > 0, блок выгоднее перерабатывать, чем везти в отвал", fmt_rub, False),
        ("ОПТИМАЛЬНОЕ НАПРАВЛЕНИЕ ДЛЯ БАЗОВОГО БЛОКА", "решение", '=IF(MAX(D30:F30)=D30,"1. В ОТВАЛ",IF(MAX(D30:F30)=E30,"2. НАПРЯМУЮ НА ОФ","3. ЧЕРЕЗ ДСК-СМС"))', "", "", "Автоматический выбор максимального NPV/т", "@", True)
    ]

    curr_r = 6
    for item in calc_rows:
        if len(item) == 7:
            title_sec = item[0]
            ws_calc.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=7)
            c = ws_calc.cell(row=curr_r, column=2, value=title_sec)
            c.font = f_sec_hdr
            c.fill = fill_sec_hdr
            ws_calc.row_dimensions[curr_r].height = 20
        else:
            name, unit, f1, f2, f3, logic, num_fmt, is_tot = item
            row_font = f_total if is_tot else f_formula
            row_fill = fill_total if is_tot else PatternFill(fill_type=None)
            row_border = border_total if is_tot else border_data

            c_name = ws_calc.cell(row=curr_r, column=2, value=name)
            c_name.font = row_font
            if row_fill.fill_type:
                c_name.fill = row_fill
            c_name.border = row_border

            c_u = ws_calc.cell(row=curr_r, column=3, value=unit)
            c_u.font = f_note
            if row_fill.fill_type:
                c_u.fill = row_fill
            c_u.border = row_border
            c_u.alignment = Alignment(horizontal="center")

            for idx_col, formula_val in enumerate((f1, f2, f3)):
                c_val = ws_calc.cell(row=curr_r, column=4 + idx_col, value=formula_val)
                if not is_tot and isinstance(formula_val, str) and "Inputs!" in formula_val:
                    c_val.font = f_link
                else:
                    c_val.font = row_font
                if row_fill.fill_type:
                    c_val.fill = row_fill
                c_val.border = row_border
                c_val.number_format = num_fmt
                if num_fmt == "@":
                    c_val.alignment = Alignment(horizontal="center")
            c_log = ws_calc.cell(row=curr_r, column=7, value=logic)
            c_log.font = f_note
            c_log.border = border_data

            if curr_r == 32:
                ws_calc.merge_cells(start_row=curr_r, start_column=4, end_row=curr_r, end_column=6)
                ws_calc.cell(row=curr_r, column=4).fill = fill_sms

        curr_r += 1

    ws_calc.column_dimensions["A"].width = 3
    ws_calc.column_dimensions["B"].width = 48
    ws_calc.column_dimensions["C"].width = 10
    ws_calc.column_dimensions["D"].width = 18
    ws_calc.column_dimensions["E"].width = 20
    ws_calc.column_dimensions["F"].width = 22
    ws_calc.column_dimensions["G"].width = 45

    ws_mat.views.sheetView[0].showGridLines = True
    ws_mat["B2"] = "МАТРИЧНЫЙ АНАЛИЗ CUT-OFF GRADE: ЧУВСТВИТЕЛЬНОСТЬ К СОДЕРЖАНИЯМ Fe И P2O5"
    ws_mat["B2"].font = f_title

    ws_mat["B3"] = "По строкам: Содержание Fe в исходной руде (%). По столбцам: Содержание P2O5 в исходной руде (%). Все ячейки пересчитываются динамически от листа Inputs."
    ws_mat["B3"].font = f_note

    fe_steps = (0.06, 0.08, 0.10, 0.12, 0.14, 0.16, 0.18, 0.20, 0.22, 0.24, 0.26, 0.28)
    p_steps = (0.005, 0.010, 0.015, 0.020, 0.025, 0.030, 0.040, 0.050, 0.060, 0.070)

    def build_matrix(start_row, title, mode):
        ws_mat.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=2 + len(p_steps))
        hdr = ws_mat.cell(row=start_row, column=2, value=title)
        hdr.font = f_sec_hdr
        hdr.fill = fill_sec_hdr
        ws_mat.row_dimensions[start_row].height = 22

        r_col_hdr = start_row + 1
        corner = ws_mat.cell(row=r_col_hdr, column=2, value="Fe \\ P2O5")
        corner.font = f_tbl_hdr
        corner.fill = fill_tbl_hdr
        corner.border = border_data
        corner.alignment = Alignment(horizontal="center")

        for j, p_val in enumerate(p_steps):
            c = ws_mat.cell(row=r_col_hdr, column=3 + j, value=p_val)
            c.font = f_input
            c.fill = fill_tbl_hdr
            c.border = border_data
            c.number_format = fmt_pct
            c.alignment = Alignment(horizontal="center")

        for i, fe_val in enumerate(fe_steps):
            curr_row = r_col_hdr + 1 + i
            c_fe = ws_mat.cell(row=curr_row, column=2, value=fe_val)
            c_fe.font = f_input
            c_fe.fill = fill_tbl_hdr
            c_fe.border = border_data
            c_fe.number_format = fmt_pct
            c_fe.alignment = Alignment(horizontal="center")

            fe_ref = f"$B{curr_row}"
            for j in range(len(p_steps)):
                col_let = get_column_letter(3 + j)
                p_ref = f"{col_let}${r_col_hdr}"

                rev_fe_sc2 = f"((1-{v_moist})*{fe_ref}*{v_fe_conv}*(1-{v_dil})*{rec_fe}/{g_fe_c}*{p_fe}*{fx})"
                rev_p_sc2  = f"((1-{v_moist})*{p_ref}*(1-{v_dil})*{rec_p}/{g_p_c}*{p_p}*{fx})"
                profit_sc2 = f"({rev_fe_sc2} + {rev_p_sc2} + Unit_Economics_Calc!$E$21 - Unit_Economics_Calc!$E$28)"

                rev_fe_sc3 = f"((1-{v_moist})*{fe_ref}*{v_fe_conv}*(1-{v_dil})*{sms_r_fe}*{rec_fe}/{g_fe_c}*{p_fe}*{fx})"
                rev_p_sc3  = f"((1-{v_moist})*{p_ref}*(1-{v_dil})*{sms_r_p}*{rec_p}/{g_p_c}*{p_p}*{fx})"
                profit_sc3 = f"({rev_fe_sc3} + {rev_p_sc3} + Unit_Economics_Calc!$F$21 + Unit_Economics_Calc!$F$22 - Unit_Economics_Calc!$F$28)"

                profit_sc1 = "Unit_Economics_Calc!$D$30"

                if mode == "DECISION":
                    formula = f'=IF(MAX({profit_sc1},{profit_sc2},{profit_sc3})={profit_sc1},"ОТВАЛ",IF(MAX({profit_sc1},{profit_sc2},{profit_sc3})={profit_sc2},"ОФ","СМС"))'
                    num_f = "@"
                elif mode == "SMS_VS_OF":
                    formula = f"={profit_sc3}-{profit_sc2}"
                    num_f = fmt_rub
                elif mode == "NET_MAX":
                    formula = f"=MAX({profit_sc1},{profit_sc2},{profit_sc3})"
                    num_f = fmt_rub

                cell = ws_mat.cell(row=curr_row, column=3 + j, value=formula)
                cell.font = f_formula
                cell.border = border_data
                cell.number_format = num_f
                if mode == "DECISION":
                    cell.alignment = Alignment(horizontal="center")

        return r_col_hdr + 1, r_col_hdr + len(fe_steps)

    r1_start, r1_end = build_matrix(5, "МАТРИЦА 1: КАРТА ПРИНЯТИЯ РЕШЕНИЙ (CUT-OFF GRADE: ОТВАЛ / ОФ / ДСК-СМС)", "DECISION")
    end_col_let = get_column_letter(2 + len(p_steps))
    range_m1 = f"C{r1_start}:{end_col_let}{r1_end}"

    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=('"ОТВАЛ"',), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=('"ОФ"',), fill=fill_of, font=Font(bold=True, color="1F4E78")))
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=('"СМС"',), fill=fill_sms, font=Font(bold=True, color="006100")))

    r2_start, r2_end = build_matrix(r1_end + 3, "МАТРИЦА 2: МАКСИМАЛЬНАЯ МАРЖИНАЛЬНАЯ ПРИБЫЛЬ НА 1 Т ПОРОДЫ (РУБ./Т)", "NET_MAX")
    range_m2 = f"C{r2_start}:{end_col_let}{r2_end}"
    ws_mat.conditional_formatting.add(range_m2, ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="num", mid_value=0, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"
    ))

    r3_start, r3_end = build_matrix(r2_end + 3, "МАТРИЦА 3: ЭФФЕКТ ОТ РУДОРАЗБОРКИ НА ДСК-СМС ПРОТИВ ПРЯМОЙ ОФ (РУБ./Т)", "SMS_VS_OF")
    range_m3 = f"C{r3_start}:{end_col_let}{r3_end}"
    ws_mat.conditional_formatting.add(range_m3, ColorScaleRule(
        start_type="min", start_color="D9E1F2",
        mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="max", end_color="70AD47"
    ))

    ws_mat.column_dimensions["A"].width = 3
    ws_mat.column_dimensions["B"].width = 14
    for j in range(len(p_steps)):
        ws_mat.column_dimensions[get_column_letter(3 + j)].width = 13

    chart = LineChart()
    chart.title = "Зависимость удельной прибыли (руб./т) от содержания Fe при разном P2O5"
    chart.style = 10
    chart.y_axis.title = "Прибыль, руб./т"
    chart.x_axis.title = "Содержание Fe в исходной руде, %"
    chart.width = 22
    chart.height = 12

    data_ref = Reference(ws_mat, min_col=3, min_row=r2_start - 1, max_col=2 + len(p_steps), max_row=r2_end)
    cats_ref = Reference(ws_mat, min_col=2, min_row=r2_start, max_row=r2_end)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    ws_mat.add_chart(chart, f"{get_column_letter(4 + len(p_steps))}5")

    wb.save(filename)
    print(f"Файл ФЭМ успешно сгенерирован: {filename}")

if __name__ == "__main__":
    create_cut_off_fem()