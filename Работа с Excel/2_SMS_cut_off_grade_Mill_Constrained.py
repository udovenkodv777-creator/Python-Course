import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart import LineChart, Reference


def create_cut_off_fem(filename="Cut_Off_Grade_MAR_FEM_v3.xlsx"):
    wb = openpyxl.Workbook()

    ws_in = wb.active
    ws_in.title = "Inputs"
    ws_calc = wb.create_sheet(title="Unit_Economics_Calc")
    ws_mat = wb.create_sheet(title="CutOff_Matrices_Visual")

    # =========================================================================
    # СТИЛИ ПО СТАНДАРТАМ BIG4 / MCKINSEY
    # =========================================================================
    font_family = "Calibri"

    f_title = Font(name=font_family, size=14, bold=True, color="1F4E78")
    f_sec_hdr = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    f_tbl_hdr = Font(name=font_family, size=10, bold=True, color="1F4E78")
    f_input = Font(name=font_family, size=10, bold=False, color="0000FF")
    f_formula = Font(name=font_family, size=10, bold=False, color="000000")
    f_link = Font(name=font_family, size=10, bold=False, color="008000")
    f_sub_link = Font(name=font_family, size=9, italic=True, color="008000")
    f_sub_lbl = Font(name=font_family, size=9, italic=True, color="595959")
    f_total = Font(name=font_family, size=10, bold=True, color="000000")
    f_note = Font(name=font_family, size=9, italic=True, color="595959")

    fill_sec_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_tbl_hdr = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_sub = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    fill_waste = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    fill_of = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_sms = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    thin_side = Side(border_style="thin", color="BFBFBF")
    double_side = Side(border_style="double", color="000000")
    top_thin = Side(border_style="thin", color="000000")
    highlight_side = Side(border_style="thick", color="FF0000")

    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(top=top_thin, bottom=double_side, left=thin_side, right=thin_side)
    border_highlight = Border(left=highlight_side, right=highlight_side, top=highlight_side, bottom=highlight_side)

    fmt_int = "#,##0"
    fmt_dec2 = "#,##0.00"
    fmt_dec4 = "#,##0.0000"
    fmt_pct = "0.00%"
    fmt_rub = '#,##0.00" ₽"'
    fmt_usd = '$#,##0.00'

    # =========================================================================
    # ЛИСТ 1: INPUTS (ИСХОДНЫЕ ДАННЫЕ, УРАВНЕНИЯ И ГОРНЫЕ УДР)
    # =========================================================================
    for sv in ws_in.views.sheetView:
        sv.showGridLines = True

    ws_in["B2"] = "ФЭМ CUT-OFF GRADE (МАР): РАСЧЕТ НА 1 Т ПИТАНИЯ ОФ (MILL-CONSTRAINED)"
    ws_in["B2"].font = f_title

    ws_in["F2"] = "ЛЕГЕНДА ФОРМАТИРОВАНИЯ (СТАНДАРТ BIG4 / McKinsey)"
    ws_in["F2"].font = f_tbl_hdr

    legend_items = (
        ("F3", "Ручной ввод (Inputs)", f_input, fill_input, "Изменяемые параметры и коэффициенты"),
        ("F4", "Формула на текущем листе", f_formula, PatternFill(fill_type=None), "Внутренние расчеты"),
        ("F5", "Ссылка на другой лист", f_link, PatternFill(fill_type=None), "Импорт данных из Inputs/Calc"),
        ("F6", "Итоговые / ключевые строки", f_total, fill_total, "Результирующие показатели")
    )
    for cell_ref, text, font_style, fill_style, desc in legend_items:
        ws_in[cell_ref] = text
        ws_in[cell_ref].font = font_style
        if fill_style.fill_type:
            ws_in[cell_ref].fill = fill_style
        ws_in[cell_ref].border = border_data

        desc_cell = ws_in[f"G{cell_ref[1:]}"]
        desc_cell.value = desc
        desc_cell.font = f_note

    inputs_data = (
        ("1. ОБЩИЕ ПАРАМЕТРЫ И БАЗОВАЯ ХИМИЯ БЛОКА", None, None, None, None, None, None),
        (None, "Целевой объем питания ОФ (сырая масса на входе ОФ)", 1.0, "т", fmt_dec2, "В Сц.3 добыча из карьера = 1 т / Выход предконцентрата", "VOL_BASE"),
        (None, "Содержание влаги в исходной руде", 0.035, "%", fmt_pct, "Пересчет сырой массы в сухую", "MOISTURE"),
        (None, "Базовое содержание Fe (общее / из FeO) в руде", 0.030, "%", fmt_pct, "Базовое значение для Calc и подсветки матриц", "FE_BASE"),
        (None, "Коэффициент пересчета FeO в Fe (если ввод в FeO)", 1.000, "доля", fmt_dec2, "1.00 если ввод сразу в Fe; 0.777 если ввод в FeO", "FEO_CONV"),
        (None, "Базовое содержание P2O5 в руде", 0.010, "%", fmt_pct, "Базовое значение для Calc и подсветки матриц", "P2O5_BASE"),
        (None, "Содержание ZrO2 в руде (попутный компонент)", 0.0005, "%", fmt_pct, "Базовое содержание ZrO2 = 0.05% (const)", "ZRO2_BASE"),
        (None, "Разубоживание при выемке в карьере (обнулено)", 0.000, "%", fmt_pct, "Установлено 0% по условию задачи", "DILUTION"),

        ("2. МАКРОЭКОНОМИКА И ЦЕНЫ РЕАЛИЗАЦИИ (FCA/FOB)", None, None, None, None, None, None),
        (None, "Курс доллара США", 92.50, "руб./$", fmt_dec2, "Курс пересчета валютной выручки", "FX_RATE"),
        (None, "Цена ЖРК (железорудный концентрат)", 115.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_FE"),
        (None, "Цена АПК (апатитовый концентрат)", 165.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_P"),
        (None, "Цена БК (бадделеитовый концентрат)", 3200.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_ZR"),

        ("3. КАЧЕСТВО КОНЦЕНТРАТОВ И УРАВНЕНИЯ ИЗВЛЕЧЕНИЯ НА ОФ", None, None, None, None, None, None),
        (None, "Содержание Fe в товарном ЖРК", 0.655, "%", fmt_pct, "Базисное качество ЖРК", "GRADE_FE_CONC"),
        (None, "Содержание P2O5 в товарном АПК", 0.390, "%", fmt_pct, "Базисное качество АПК", "GRADE_P_CONC"),
        (None, "Содержание ZrO2 в товарном БК", 0.985, "%", fmt_pct, "Базисное качество БК", "GRADE_ZR_CONC"),
        (None, "ЖРК: Асимптота 'a' (105.5884 - 432.9662/(Fe%*100))/100", 105.5884, "коэф.", fmt_dec4, "Верхний предел гиперболы извлечения ЖРК (%)", "OF_FE_A"),
        (None, "ЖРК: Коэффициент кривизны 'b'", 432.9662, "коэф.", fmt_dec4, "Делится на (Fe% * 100) в питании ОФ", "OF_FE_B"),
        (None, "АПК: Асимптота 'a' (91.5199 - 155.7349/(P2O5%*100))/100", 91.5199, "коэф.", fmt_dec4, "Верхний предел гиперболы извлечения АПК (%)", "OF_P_A"),
        (None, "АПК: Коэффициент кривизны 'b'", 155.7349, "коэф.", fmt_dec4, "Делится на (P2O5% * 100) в питании ОФ", "OF_P_B"),
        (None, "БК: Асимптота 'a' (26.1905 - 1.0972/(ZrO2%*100))/100", 26.1905, "коэф.", fmt_dec4, "Верхний предел гиперболы извлечения БК (%)", "OF_ZR_A"),
        (None, "БК: Коэффициент кривизны 'b'", 1.0972, "коэф.", fmt_dec4, "Делится на (ZrO2% * 100) в питании ОФ", "OF_ZR_B"),
        (None, "Максимальный потолок извлечения Fe на ОФ", 0.950, "%", fmt_pct, "Физическое ограничение извлечения ЖРК", "MAX_REC_FE"),
        (None, "Максимальный потолок извлечения P2O5 на ОФ", 0.920, "%", fmt_pct, "Физическое ограничение извлечения АПК", "MAX_REC_P"),
        (None, "Максимальный потолок извлечения ZrO2 на ОФ", 0.650, "%", fmt_pct, "Физическое ограничение извлечения БК", "MAX_REC_ZR"),

        ("4. ПАРАМЕТРЫ РУДОРАЗБОРКИ НА ДСК-СМС (2 СТАДИИ + ГРОХОТ)", None, None, None, None, None, None),
        (None, "Выход магн. продукта (2.1): коэф. 'a' (3.0134*Fe%*100 - 12.079)", 3.0134, "коэф.", fmt_dec4, "Наклон прямой выхода магнитного продукта", "SMS_YM_A"),
        (None, "Выход магн. продукта (2.1): свободный член 'b'", 12.0790, "коэф.", fmt_dec4, "Вычитается в формуле выхода магнитного продукта", "SMS_YM_B"),
        (None, "Выход объед. предконцентрата (2.3): свободный член 'b'", 14.0790, "коэф.", fmt_dec4, "В формуле (3.0134*Fe%*100 - 14.079 + ВыхНемаг*0.6432)/100", "SMS_YC_B"),
        (None, "Доля подрешетного класса -20+0мм в немагн. продукте (2.3)", 0.6432, "доля", fmt_dec4, "64.32% немагнитного продукта идет в предконцентрат", "SMS_SUB_SHARE"),
        (None, "Извлечение Fe в магн. продукт (2.4): коэф. 'a'", 3.6999, "коэф.", fmt_dec4, "В формуле (3.6999*Fe%*100 - 3.9471)/100", "SMS_RFE_M_A"),
        (None, "Извлечение Fe в магн. продукт (2.4): свободный член 'b'", 3.9471, "коэф.", fmt_dec4, "Вычитается в формуле извлечения Fe в магн. продукт", "SMS_RFE_M_B"),
        (None, "Максимальный потолок извлечения Fe в магн. продукт СМС", 0.780, "%", fmt_pct, "Технологический предел магнитной сепарации (78%)", "MAX_SMS_RFE_M"),
        (None, "Доля извлечения Fe из немагн. продукта в предконцентрат (2.6)", 0.7161, "%", fmt_pct, "71.61% от Fe в немагнитном продукте уходит на ОФ", "SMS_RFE_NM_SHARE"),
        (None, "Извлечение P2O5 в предконцентрат (2.7): коэф. 'a'", 1.1843, "коэф.", fmt_dec4, "В формуле (1.1843*P2O5%*100 + 18.814 + 58.94)/100", "SMS_RP_A"),
        (None, "Извлечение P2O5 в предконцентрат (2.7): слагаемое 'b1'", 18.8140, "коэф.", fmt_dec4, "Первая константа в формуле извлечения P2O5", "SMS_RP_B1"),
        (None, "Извлечение P2O5 в предконцентрат (2.7): слагаемое 'b2'", 58.9400, "коэф.", fmt_dec4, "Вторая константа в формуле извлечения P2O5", "SMS_RP_B2"),
        (None, "Извлечение ZrO2 в объединенный предконцентрат СМС", 0.850, "%", fmt_pct, "Доля бадделеита, уходящая на ОФ", "SMS_REC_ZR"),
        (None, "Максимальный потолок выхода предконцентрата ДСК-СМС", 0.920, "%", fmt_pct, "Минимум 8% крупной породы всегда отсекается в отвал", "MAX_SMS_YIELD"),
        (None, "Опция хвостов СМС (+20мм) (1 = в отвал; 2 = щебень для дорог)", 1, "код", fmt_int, "1 - вывоз в отвал; 2 - штабель щебня + экономия", "SMS_TAIL_OPT"),
        (None, "Экономия / ценность щебня из хвостов СМС (+20мм)", 350.00, "руб./т", fmt_rub, "Альтернативная стоимость покупки щебня для дорог", "GRAVEL_VAL"),

        ("5. ЛОГИСТИКА: ПЛЕЧИ ТРАНСПОРТИРОВКИ", None, None, None, None, None, None),
        (None, "Плечо 1: Карьер -> Отвал вскрыши (Сценарий 1)", 3.20, "км", fmt_dec2, "Вывоз пустой породы", "DIST_WASTE"),
        (None, "Плечо 2: Карьер -> ОФ напрямую (Сценарий 2)", 6.50, "км", fmt_dec2, "Прямая доставка руды на фабрику", "DIST_OF"),
        (None, "Плечо 3а: Карьер -> ДСК-СМС (Сценарий 3)", 2.10, "км", fmt_dec2, "Доставка всего добытого объема на корпус СМС", "DIST_TO_SMS"),
        (None, "Плечо 3б: ДСК-СМС -> ОФ (объединенный предконцентрат)", 4.80, "км", fmt_dec2, "Доставка 1 т предконцентрата на ОФ", "DIST_SMS_OF"),
        (None, "Плечо 3в (опция 1): ДСК-СМС -> Отвал хвостов (+20мм)", 1.80, "км", fmt_dec2, "Вывоз хвостов СМС в отвал", "DIST_SMS_WASTE"),
        (None, "Плечо 3г (опция 2): ДСК-СМС -> Штабель щебня (+20мм)", 0.90, "км", fmt_dec2, "Доставка хвостов СМС на склад дорожного щебня", "DIST_SMS_GRAVEL"),

        ("6. УДЕЛЬНЫЕ ОПЕРАЦИОННЫЕ РАСХОДЫ: ДОБЫЧА (НДПИ, БВР, ЭКСКАВАЦИЯ)", None, None, None, None, None, None),
        (None, "УДР НДПИ (налог на добычу полезного ископаемого)", 120.00, "руб./т руды", fmt_rub, "Начисляется на весь добытый объем руды (Сц.2 и Сц.3)", "OPEX_NDPI"),
        (None, "УДР БВР (буровзрывные работы в карьере)", 85.00, "руб./т г.м.", fmt_rub, "Начисляется на всю взорванную горную массу (Сц.1-3)", "OPEX_BVR"),
        (None, "УДР Экскавации (погрузка в карьерный транспорт)", 105.00, "руб./т г.м.", fmt_rub, "Начисляется на всю отгруженную горную массу (Сц.1-3)", "OPEX_EXC"),

        ("7. УДЕЛЬНЫЕ ОПЕРАЦИОННЫЕ РАСХОДЫ: ТРАНСПОРТ И ОБОГАЩЕНИЕ", None, None, None, None, None, None),
        (None, "УДР транспортировки горной массы / продуктов", 28.50, "руб./т*км", fmt_rub, "Карьерный автотранспорт / конвейер", "OPEX_TR"),
        (None, "УДР дробления на ОФ (стандартная схема)", 110.00, "руб./т", fmt_rub, "Крупное, среднее и мелкое дробление на ОФ", "OPEX_CRUSH_OF"),
        (None, "УДР передела ДСК-СМС (2 стадии дробления + СМС + грохот)", 145.00, "руб./т", fmt_rub, "На 1 т сухой руды, поступившей на ДСК-СМС", "OPEX_SMS"),
        (None, "УДР переработки ЖРК (измельчение -> концентрат)", 420.00, "руб./т руды", fmt_rub, "На 1 т сухой руды, поступившей на ОФ", "OPEX_OF_FE"),
        (None, "УДР переработки АПК (флотация -> концентрат)", 290.00, "руб./т руды", fmt_rub, "На 1 т сухой руды, поступившей на ОФ", "OPEX_OF_P"),
        (None, "УДР переработки БК (гравитация/доводка -> конц.)", 85.00, "руб./т руды", fmt_rub, "На 1 т сухой руды, поступившей на ОФ", "OPEX_OF_ZR")
    )

    row_idx = 8
    cell_map = {}

    headers = ("Параметр / Статья", "Значение", "Ед. изм.", "Методическое примечание")
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
    ws_in.column_dimensions["B"].width = 60
    ws_in.column_dimensions["C"].width = 16
    ws_in.column_dimensions["D"].width = 14
    ws_in.column_dimensions["E"].width = 55
    ws_in.column_dimensions["F"].width = 32
    ws_in.column_dimensions["G"].width = 35
    # =========================================================================
    # ЛИСТ 2: UNIT_ECONOMICS_CALC (БАЛАНС НА 1 Т ПИТАНИЯ ОФ + ДОП. ДОБЫЧА В СЦ.3)
    # =========================================================================
    for sv in ws_calc.views.sheetView:
        sv.showGridLines = True

    ws_calc.sheet_properties.outlinePr.summaryBelow = False

    ws_calc["B2"] = "МОДЕЛЬ UNIT-ЭКОНОМИКИ НА 1 Т ПИТАНИЯ ОФ (С УЧЕТОМ ДОП. ДОБЫЧИ ДЛЯ ДСК-СМС)"
    ws_calc["B2"].font = f_title

    ws_calc["B3"] = "В Сценарии 3 на вход ОФ подается ровно 1,00 т предконцентрата, поэтому объем добычи из карьера увеличивается пропорционально (1 / Выход СМС)."
    ws_calc["B3"].font = f_note

    calc_headers = (
        "Показатель / Статья расчета",
        "Ед. изм.",
        "Сценарий 1: В отвал",
        "Сценарий 2: Прямая ОФ",
        "Сценарий 3: ДСК-СМС -> ОФ",
        "Формула / Логика расчета"
    )
    r = 5
    for idx, text in enumerate(calc_headers):
        c = ws_calc.cell(row=r, column=2 + idx, value=text)
        c.font = f_tbl_hdr
        c.fill = fill_tbl_hdr
        c.border = border_data
        c.alignment = Alignment(horizontal="center" if idx > 0 else "left", wrap_text=True)
    ws_calc.row_dimensions[r].height = 28

    curr_r = 6

    def add_section(title):
        nonlocal curr_r
        ws_calc.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=7)
        c = ws_calc.cell(row=curr_r, column=2, value=title)
        c.font = f_sec_hdr
        c.fill = fill_sec_hdr
        ws_calc.row_dimensions[curr_r].height = 20
        curr_r += 1

    def add_row(name, unit, f1, f2, f3, logic, num_fmt, row_type="NORMAL"):
        nonlocal curr_r
        row_idx_now = curr_r
        if row_type == "TOTAL":
            r_font = f_total
            r_fill = fill_total
            r_border = border_total
        elif row_type == "SUB":
            r_font = f_sub_lbl
            r_fill = fill_sub
            r_border = border_data
            ws_calc.row_dimensions[row_idx_now].outline_level = 1
            ws_calc.row_dimensions[row_idx_now].hidden = True
        else:
            r_font = f_formula
            r_fill = PatternFill(fill_type=None)
            r_border = border_data

        c_name = ws_calc.cell(row=row_idx_now, column=2, value=name)
        c_name.font = r_font
        if r_fill.fill_type:
            c_name.fill = r_fill
        c_name.border = r_border

        c_u = ws_calc.cell(row=row_idx_now, column=3, value=unit)
        c_u.font = f_note
        if r_fill.fill_type:
            c_u.fill = r_fill
        c_u.border = r_border
        c_u.alignment = Alignment(horizontal="center")

        for idx_col, val in enumerate((f1, f2, f3)):
            c_val = ws_calc.cell(row=row_idx_now, column=4 + idx_col, value=val)
            if row_type == "SUB":
                c_val.font = f_sub_link if (isinstance(val, str) and "Inputs!" in val) else f_sub_lbl
            elif row_type != "TOTAL" and isinstance(val, str) and "Inputs!" in val:
                c_val.font = f_link
            else:
                c_val.font = r_font

            if r_fill.fill_type:
                c_val.fill = r_fill
            c_val.border = r_border
            c_val.number_format = num_fmt
            if num_fmt == "@":
                c_val.alignment = Alignment(horizontal="center")

        c_log = ws_calc.cell(row=row_idx_now, column=7, value=logic)
        c_log.font = f_note
        if r_fill.fill_type:
            c_log.fill = r_fill
        c_log.border = r_border

        curr_r += 1
        return row_idx_now

    # --- СЕКЦИЯ 1: МАТЕРИАЛЬНЫЙ БАЛАНС И КАЧЕСТВО ПИТАНИЯ ОФ ---
    add_section("1. МАТЕРИАЛЬНЫЙ БАЛАНС (ДОБЫЧА ИЗ КАРЬЕРА ДЛЯ ПОДАЧИ 1,00 Т НА ОФ)")

    r_fe_raw = add_row("Исходное содержание Fe в руде из карьера", "%", 0, 0, 0, "Качество недровой руды до обогащения", fmt_pct)
    r_fe_b_l = add_row("   в т.ч. базовое содержание Fe в руде (ссылка)", "%", 0, f"={cell_map['FE_BASE']}", f"={cell_map['FE_BASE']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_fe_c_l = add_row("   в т.ч. коэффициент пересчета FeO -> Fe (ссылка)", "доля", 0, f"={cell_map['FEO_CONV']}", f"={cell_map['FEO_CONV']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_dil_l = add_row("   в т.ч. разубоживание при добыче (ссылка)", "%", 0, f"={cell_map['DILUTION']}", f"={cell_map['DILUTION']}", "Импорт из Inputs (обнулено = 0%)", fmt_pct, "SUB")
    ws_calc[f"E{r_fe_raw}"] = f"=E{r_fe_b_l}*E{r_fe_c_l}*(1-E{r_dil_l})"
    ws_calc[f"F{r_fe_raw}"] = f"=F{r_fe_b_l}*F{r_fe_c_l}*(1-F{r_dil_l})"

    r_p_raw = add_row("Исходное содержание P2O5 в руде из карьера", "%", 0, 0, 0, "Качество недровой руды до обогащения", fmt_pct)
    r_p_b_l = add_row("   в т.ч. базовое содержание P2O5 в руде (ссылка)", "%", 0, f"={cell_map['P2O5_BASE']}", f"={cell_map['P2O5_BASE']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_p_raw}"] = f"=E{r_p_b_l}*(1-E{r_dil_l})"
    ws_calc[f"F{r_p_raw}"] = f"=F{r_p_b_l}*(1-F{r_dil_l})"

    r_zr_raw = add_row("Исходное содержание ZrO2 в руде из карьера", "%", 0, 0, 0, "Качество недровой руды до обогащения", fmt_pct)
    r_zr_b_l = add_row("   в т.ч. базовое содержание ZrO2 в руде (ссылка)", "%", 0, f"={cell_map['ZRO2_BASE']}", f"={cell_map['ZRO2_BASE']}", "Импорт из Inputs (0.05%)", fmt_pct, "SUB")
    ws_calc[f"E{r_zr_raw}"] = f"=E{r_zr_b_l}*(1-E{r_dil_l})"
    ws_calc[f"F{r_zr_raw}"] = f"=F{r_zr_b_l}*(1-F{r_dil_l})"

    r_sms_yield_tot = add_row("Выход объединенного предконцентрата ДСК-СМС (шаг 2.3)", "%", 0, 1.0, 0, "Доля массы руды, уходящей после СМС на ОФ", fmt_pct)
    r_ym_calc = add_row("   2.1. Выход магнитного продукта СМС (расчет)", "%", 0, 0, 0, "MAX(0, (3.0134*Fe%*100 - 12.079)/100)", fmt_pct, "SUB")
    r_ym_a_l = add_row("      коэф. 'a' выхода магн. продукта (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_YM_A']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_ym_b_l = add_row("      свободный член 'b' выхода магн. продукта (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_YM_B']}", "Импорт из Inputs", fmt_dec4, "SUB")
    ws_calc[f"F{r_ym_calc}"] = f"=MAX(0, MIN(1, (F{r_ym_a_l}*F{r_fe_raw}*100 - F{r_ym_b_l})/100))"

    r_ynm_calc = add_row("   2.2. Выход немагнитного продукта СМС (расчет)", "%", 0, 0, f"=1-F{r_ym_calc}", "100% - Выход магнитного продукта", fmt_pct, "SUB")
    r_yc_b_l = add_row("      свободный член 'b' объед. предконцентрата (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_YC_B']}", "Импорт из Inputs (14.079)", fmt_dec4, "SUB")
    r_sub_sh_l = add_row("      доля класса -20+0мм в немагн. продукте (ссылка)", "доля", 0, 0, f"={cell_map['SMS_SUB_SHARE']}", "Импорт из Inputs (0.6432)", fmt_dec4, "SUB")
    r_max_yc_l = add_row("      максимальный потолок выхода предконцентрата (ссылка)", "%", 0, 0, f"={cell_map['MAX_SMS_YIELD']}", "Импорт из Inputs (92%)", fmt_pct, "SUB")
    ws_calc[f"F{r_sms_yield_tot}"] = f"=MAX(0.01, MIN(F{r_max_yc_l}, (F{r_ym_a_l}*F{r_fe_raw}*100 - F{r_yc_b_l} + (F{r_ynm_calc}*100*F{r_sub_sh_l}))/100))"

    r_target_of_wet = add_row("Целевой поток сырой массы на входе ОФ / базовый блок", "т", f"={cell_map['VOL_BASE']}", f"={cell_map['VOL_BASE']}", f"={cell_map['VOL_BASE']}", "Линк: 1,00 т на входе ОФ в Сц.2 и Сц.3", fmt_dec2)
    r_wet_mined = add_row("Необходимый объем добычи сырой руды из карьера", "т", f"=D{r_target_of_wet}", f"=E{r_target_of_wet}", f"=F{r_target_of_wet}/F{r_sms_yield_tot}", "В Сц.3 добывается 1 / Выход СМС для загрузки ОФ", fmt_dec2)
    r_moist_l = add_row("   в т.ч. влажность исходной руды (ссылка)", "%", f"={cell_map['MOISTURE']}", f"={cell_map['MOISTURE']}", f"={cell_map['MOISTURE']}", "Импорт из Inputs", fmt_pct, "SUB")

    r_dry_mined = add_row("Добытая сухая горная масса из карьера", "т", f"=D{r_wet_mined}*(1-D{r_moist_l})", f"=E{r_wet_mined}*(1-E{r_moist_l})", f"=F{r_wet_mined}*(1-F{r_moist_l})", "Сухой вес добытого потока", fmt_dec2)
    r_of_mass = add_row("Масса сухой руды, поступающей на ОФ (одинакова в Сц.2 и Сц.3)", "т", 0, f"=E{r_dry_mined}", f"=F{r_dry_mined}*F{r_sms_yield_tot}", "Равна 0,965 т сухой руды в обоих сценариях переработки", fmt_dec2)
    r_tail_mass = add_row("Масса отсеченных хвостов СМС (+20мм в отвал / на щебень)", "т", 0, 0, f"=F{r_dry_mined}-F{r_of_mass}", "Дополнительная пустая порода, удаленная на СМС", fmt_dec2)

    r_fe_feed = add_row("Содержание Fe в питании ОФ", "%", 0, f"=E{r_fe_raw}", 0, "Качество обогащенного предконцентрата по Fe", fmt_pct)
    r_rfe_m_calc = add_row("   2.4. Извлечение Fe в магнитный продукт (расчет)", "%", 0, 0, 0, "MAX(0, MIN(0.78, (3.6999*Fe%*100 - 3.9471)/100))", fmt_pct, "SUB")
    r_rfe_ma_l = add_row("      коэф. 'a' извлечения Fe в магн. продукт (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_RFE_M_A']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_rfe_mb_l = add_row("      свободный член 'b' извлечения Fe (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_RFE_M_B']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_max_rfem_l = add_row("      потолок извлечения Fe в магн. продукт (ссылка)", "%", 0, 0, f"={cell_map['MAX_SMS_RFE_M']}", "Импорт из Inputs (78%)", fmt_pct, "SUB")
    ws_calc[f"F{r_rfe_m_calc}"] = f"=MAX(0, MIN(F{r_max_rfem_l}, (F{r_rfe_ma_l}*F{r_fe_raw}*100 - F{r_rfe_mb_l})/100))"

    r_rfe_nm_calc = add_row("   2.5. Извлечение Fe в немагнитный продукт (расчет)", "%", 0, 0, f"=1-F{r_rfe_m_calc}", "100% - Извлечение в магнитный продукт", fmt_pct, "SUB")
    r_rfe_nm_sh_l = add_row("      доля извлечения Fe из немагн. в предконц. (ссылка)", "%", 0, 0, f"={cell_map['SMS_RFE_NM_SHARE']}", "Импорт из Inputs (71.61%)", fmt_pct, "SUB")
    r_rfe_tot_sms = add_row("   2.6. Извлечение Fe в объед. предконцентрат (расчет)", "%", 0, 1.0, f"=MIN(1, F{r_rfe_m_calc} + F{r_rfe_nm_calc}*F{r_rfe_nm_sh_l})", "Суммарное извлечение Fe после СМС и грохота", fmt_pct, "SUB")
    ws_calc[f"F{r_fe_feed}"] = f"=(F{r_fe_raw}*F{r_rfe_tot_sms})/F{r_sms_yield_tot}"

    r_p_feed = add_row("Содержание P2O5 в питании ОФ", "%", 0, f"=E{r_p_raw}", 0, "Качество обогащенного предконцентрата по P2O5", fmt_pct)
    r_rp_tot_sms = add_row("   2.7. Извлечение P2O5 в объед. предконцентрат (расчет)", "%", 0, 1.0, 0, "MIN(1, (1.1843*P2O5%*100 + 18.814 + 58.94)/100)", fmt_pct, "SUB")
    r_rp_a_l = add_row("      коэф. 'a' извлечения P2O5 в предконц. (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_RP_A']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_rp_b1_l = add_row("      слагаемое 'b1' извлечения P2O5 (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_RP_B1']}", "Импорт из Inputs (18.814)", fmt_dec4, "SUB")
    r_rp_b2_l = add_row("      слагаемое 'b2' извлечения P2O5 (ссылка)", "коэф.", 0, 0, f"={cell_map['SMS_RP_B2']}", "Импорт из Inputs (58.94)", fmt_dec4, "SUB")
    ws_calc[f"F{r_rp_tot_sms}"] = f"=MAX(0, MIN(1, (F{r_rp_a_l}*F{r_p_raw}*100 + F{r_rp_b1_l} + F{r_rp_b2_l})/100))"
    ws_calc[f"F{r_p_feed}"] = f"=(F{r_p_raw}*F{r_rp_tot_sms})/F{r_sms_yield_tot}"

    r_zr_feed = add_row("Содержание ZrO2 в питании ОФ", "%", 0, f"=E{r_zr_raw}", 0, "Качество обогащенного предконцентрата по ZrO2", fmt_pct)
    r_sms_rzr_l = add_row("   в т.ч. извлечение ZrO2 в предконцентрат СМС (ссылка)", "%", 0, 1.0, f"={cell_map['SMS_REC_ZR']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"F{r_zr_feed}"] = f"=(F{r_zr_raw}*F{r_sms_rzr_l})/F{r_sms_yield_tot}"

    # --- СЕКЦИЯ 2: ВЫПУСК ТОВАРНЫХ КОНЦЕНТРАТОВ (ГИПЕРБОЛИЧЕСКИЕ УРАВНЕНИЯ ОФ) ---
    add_section("2. ВЫПУСК ТОВАРНЫХ КОНЦЕНТРАТОВ НА ОФ (ПРИ ОДИНАКОВОЙ ЗАГРУЗКЕ ОФ = 1 Т)")
    r_rec_fe_calc = add_row("Сквозное извлечение Fe в ЖРК на ОФ (шаг 1)", "%", 0, 0, 0, "MAX(0, MIN(Потолок, (105.5884 - 432.9662/(Fe%*100))/100))", fmt_pct)
    r_offe_a_l = add_row("   в т.ч. асимптота 'a' извлечения ЖРК (ссылка)", "коэф.", 0, f"={cell_map['OF_FE_A']}", f"={cell_map['OF_FE_A']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_offe_b_l = add_row("   в т.ч. коэф. кривизны 'b' извлечения ЖРК (ссылка)", "коэф.", 0, f"={cell_map['OF_FE_B']}", f"={cell_map['OF_FE_B']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_max_fe_l = add_row("   в т.ч. потолок извлечения Fe на ОФ (ссылка)", "%", 0, f"={cell_map['MAX_REC_FE']}", f"={cell_map['MAX_REC_FE']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_rec_fe_calc}"] = f"=IFERROR(MAX(0, MIN(E{r_max_fe_l}, (E{r_offe_a_l} - E{r_offe_b_l}/(E{r_fe_feed}*100))/100)), 0)"
    ws_calc[f"F{r_rec_fe_calc}"] = f"=IFERROR(MAX(0, MIN(F{r_max_fe_l}, (F{r_offe_a_l} - F{r_offe_b_l}/(F{r_fe_feed}*100))/100)), 0)"

    r_prod_fe = add_row("Выпуск ЖРК (железорудного концентрата)", "т", 0, 0, 0, "Прирост в Сц.3 за счет более богатого питания ОФ", fmt_dec2)
    r_g_fe_l = add_row("   в т.ч. содержание Fe в товарном ЖРК (ссылка)", "%", 0, f"={cell_map['GRADE_FE_CONC']}", f"={cell_map['GRADE_FE_CONC']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_prod_fe}"] = f"=(E{r_of_mass}*E{r_fe_feed}*E{r_rec_fe_calc})/E{r_g_fe_l}"
    ws_calc[f"F{r_prod_fe}"] = f"=(F{r_of_mass}*F{r_fe_feed}*F{r_rec_fe_calc})/F{r_g_fe_l}"

    r_rec_p_calc = add_row("Сквозное извлечение P2O5 в АПК на ОФ (шаг 1)", "%", 0, 0, 0, "MAX(0, MIN(Потолок, (91.5199 - 155.7349/(P2O5%*100))/100))", fmt_pct)
    r_ofp_a_l = add_row("   в т.ч. асимптота 'a' извлечения АПК (ссылка)", "коэф.", 0, f"={cell_map['OF_P_A']}", f"={cell_map['OF_P_A']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_ofp_b_l = add_row("   в т.ч. коэф. кривизны 'b' извлечения АПК (ссылка)", "коэф.", 0, f"={cell_map['OF_P_B']}", f"={cell_map['OF_P_B']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_max_p_l = add_row("   в т.ч. потолок извлечения P2O5 на ОФ (ссылка)", "%", 0, f"={cell_map['MAX_REC_P']}", f"={cell_map['MAX_REC_P']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_rec_p_calc}"] = f"=IFERROR(MAX(0, MIN(E{r_max_p_l}, (E{r_ofp_a_l} - E{r_ofp_b_l}/(E{r_p_feed}*100))/100)), 0)"
    ws_calc[f"F{r_rec_p_calc}"] = f"=IFERROR(MAX(0, MIN(F{r_max_p_l}, (F{r_ofp_a_l} - F{r_ofp_b_l}/(F{r_p_feed}*100))/100)), 0)"

    r_prod_p = add_row("Выпуск АПК (апатитового концентрата)", "т", 0, 0, 0, "Прирост в Сц.3 за счет более богатого питания ОФ", fmt_dec2)
    r_g_p_l = add_row("   в т.ч. содержание P2O5 в товарном АПК (ссылка)", "%", 0, f"={cell_map['GRADE_P_CONC']}", f"={cell_map['GRADE_P_CONC']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_prod_p}"] = f"=(E{r_of_mass}*E{r_p_feed}*E{r_rec_p_calc})/E{r_g_p_l}"
    ws_calc[f"F{r_prod_p}"] = f"=(F{r_of_mass}*F{r_p_feed}*F{r_rec_p_calc})/F{r_g_p_l}"

    r_rec_zr_calc = add_row("Сквозное извлечение ZrO2 в БК на ОФ (шаг 1)", "%", 0, 0, 0, "MAX(0, MIN(Потолок, (26.1905 - 1.0972/(ZrO2%*100))/100))", fmt_pct)
    r_ofzr_a_l = add_row("   в т.ч. асимптота 'a' извлечения БК (ссылка)", "коэф.", 0, f"={cell_map['OF_ZR_A']}", f"={cell_map['OF_ZR_A']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_ofzr_b_l = add_row("   в т.ч. коэф. кривизны 'b' извлечения БК (ссылка)", "коэф.", 0, f"={cell_map['OF_ZR_B']}", f"={cell_map['OF_ZR_B']}", "Импорт из Inputs", fmt_dec4, "SUB")
    r_max_zr_l = add_row("   в т.ч. потолок извлечения ZrO2 на ОФ (ссылка)", "%", 0, f"={cell_map['MAX_REC_ZR']}", f"={cell_map['MAX_REC_ZR']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_rec_zr_calc}"] = f"=IFERROR(MAX(0, MIN(E{r_max_zr_l}, (E{r_ofzr_a_l} - E{r_ofzr_b_l}/(E{r_zr_feed}*100))/100)), 0)"
    ws_calc[f"F{r_rec_zr_calc}"] = f"=IFERROR(MAX(0, MIN(F{r_max_zr_l}, (F{r_ofzr_a_l} - F{r_ofzr_b_l}/(F{r_zr_feed}*100))/100)), 0)"

    r_prod_zr = add_row("Выпуск БК (бадделеитового концентрата)", "т", 0, 0, 0, "Прирост в Сц.3 за счет более богатого питания ОФ", fmt_dec2)
    r_g_zr_l = add_row("   в т.ч. содержание ZrO2 в товарном БК (ссылка)", "%", 0, f"={cell_map['GRADE_ZR_CONC']}", f"={cell_map['GRADE_ZR_CONC']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_prod_zr}"] = f"=(E{r_of_mass}*E{r_zr_feed}*E{r_rec_zr_calc})/E{r_g_zr_l}"
    ws_calc[f"F{r_prod_zr}"] = f"=(F{r_of_mass}*F{r_zr_feed}*F{r_rec_zr_calc})/F{r_g_zr_l}"

    # --- СЕКЦИЯ 3: ВЫРУЧКА ОТ РЕАЛИЗАЦИИ И ЭКОНОМИЯ (NSR) ---
    add_section("3. ВЫРУЧКА ОТ РЕАЛИЗАЦИИ И ЭКОНОМИЯ (NSR)")
    r_fx_l = add_row("   Курс доллара США (ссылка для пересчета)", "руб./$", f"={cell_map['FX_RATE']}", f"={cell_map['FX_RATE']}", f"={cell_map['FX_RATE']}", "Импорт курса из Inputs", fmt_dec2, "SUB")

    r_rev_fe = add_row("Выручка от реализации ЖРК", "руб.", 0, 0, 0, "Выпуск ЖРК * Цена ЖРК * Курс", fmt_rub)
    r_pfe_l = add_row("   в т.ч. цена реализации ЖРК (ссылка)", "$/т", 0, f"={cell_map['PRICE_FE']}", f"={cell_map['PRICE_FE']}", "Импорт из Inputs", fmt_usd, "SUB")
    ws_calc[f"E{r_rev_fe}"] = f"=E{r_prod_fe}*E{r_pfe_l}*E{r_fx_l}"
    ws_calc[f"F{r_rev_fe}"] = f"=F{r_prod_fe}*F{r_pfe_l}*F{r_fx_l}"

    r_rev_p = add_row("Выручка от реализации АПК", "руб.", 0, 0, 0, "Выпуск АПК * Цена АПК * Курс", fmt_rub)
    r_pp_l = add_row("   в т.ч. цена реализации АПК (ссылка)", "$/т", 0, f"={cell_map['PRICE_P']}", f"={cell_map['PRICE_P']}", "Импорт из Inputs", fmt_usd, "SUB")
    ws_calc[f"E{r_rev_p}"] = f"=E{r_prod_p}*E{r_pp_l}*E{r_fx_l}"
    ws_calc[f"F{r_rev_p}"] = f"=F{r_prod_p}*F{r_pp_l}*F{r_fx_l}"

    r_rev_zr = add_row("Выручка от реализации БК", "руб.", 0, 0, 0, "Выпуск БК * Цена БК * Курс", fmt_rub)
    r_pzr_l = add_row("   в т.ч. цена реализации БК (ссылка)", "$/т", 0, f"={cell_map['PRICE_ZR']}", f"={cell_map['PRICE_ZR']}", "Импорт из Inputs", fmt_usd, "SUB")
    ws_calc[f"E{r_rev_zr}"] = f"=E{r_prod_zr}*E{r_pzr_l}*E{r_fx_l}"
    ws_calc[f"F{r_rev_zr}"] = f"=F{r_prod_zr}*F{r_pzr_l}*F{r_fx_l}"

    r_rev_gravel = add_row("Экономия на покупке дорожного щебня (хвосты +20мм)", "руб.", 0, 0, 0, "Активно при выборе опции 2 в Inputs", fmt_rub)
    r_opt_l = add_row("   в т.ч. опция использования хвостов СМС (ссылка)", "код", 0, 0, f"={cell_map['SMS_TAIL_OPT']}", "Импорт из Inputs (1 = отвал, 2 = щебень)", fmt_int, "SUB")
    r_gval_l = add_row("   в т.ч. ценность/экономия на щебне (ссылка)", "руб./т", 0, 0, f"={cell_map['GRAVEL_VAL']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"F{r_rev_gravel}"] = f"=IF(F{r_opt_l}=2, (F{r_tail_mass}/(1-F{r_moist_l}))*F{r_gval_l}, 0)"

    r_gross_rev = add_row("ИТОГО ДОХОД / ЦЕННОСТЬ БЛОКА (GROSS VALUE)", "руб.", f"=D{r_rev_fe}+D{r_rev_p}+D{r_rev_zr}+D{r_rev_gravel}", f"=E{r_rev_fe}+E{r_rev_p}+E{r_rev_zr}+E{r_rev_gravel}", f"=F{r_rev_fe}+F{r_rev_p}+F{r_rev_zr}+F{r_rev_gravel}", "Сумма выручки от концентратов и экономии", fmt_rub, "TOTAL")

    # --- СЕКЦИЯ 4: ПЕРЕМЕННЫЕ ОПЕРАЦИОННЫЕ ЗАТРАТЫ (OPEX С УЧЕТОМ ДОП. ДОБЫЧИ) ---
    add_section("4. ОПЕРАЦИОННЫЕ ЗАТРАТЫ (ДОБЫЧА ДОП. ОБЪЕМА, ТРАНСПОРТ И ОБОГАЩЕНИЕ)")

    r_opex_ndpi = add_row("НДПИ (на доп. объем добычи руды в Сц.3)", "руб.", 0, 0, 0, "На базовую 1 т — Sunk cost; в Сц.3 начисляется на (Добыча - 1 т)", fmt_rub)
    r_cndpi_l = add_row("   в т.ч. ставка / УДР НДПИ на 1 т добытой руды (ссылка)", "руб./т", 0, 0, f"={cell_map['OPEX_NDPI']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"F{r_opex_ndpi}"] = f"=(F{r_wet_mined}-F{r_target_of_wet})*F{r_cndpi_l}"

    r_opex_bvr = add_row("Буровзрывные работы (на доп. объем добычи в Сц.3)", "руб.", 0, 0, 0, "На базовую 1 т — Sunk cost; в Сц.3 начисляется на (Добыча - 1 т)", fmt_rub)
    r_cbvr_l = add_row("   в т.ч. УДР БВР на 1 т горной массы (ссылка)", "руб./т", 0, 0, f"={cell_map['OPEX_BVR']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"F{r_opex_bvr}"] = f"=(F{r_wet_mined}-F{r_target_of_wet})*F{r_cbvr_l}"

    r_opex_exc = add_row("Экскавация (на доп. объем добычи в Сц.3)", "руб.", 0, 0, 0, "На базовую 1 т — Sunk cost; в Сц.3 начисляется на (Добыча - 1 т)", fmt_rub)
    r_cexc_l = add_row("   в т.ч. УДР экскавации на 1 т горной массы (ссылка)", "руб./т", 0, 0, f"={cell_map['OPEX_EXC']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"F{r_opex_exc}"] = f"=(F{r_wet_mined}-F{r_target_of_wet})*F{r_cexc_l}"

    r_opex_tr = add_row("Транспортировка горной массы и продуктов", "руб.", 0, 0, 0, "Учитывает перевозку доп. добытой руды на ДСК-СМС", fmt_rub)
    r_d_main_l = add_row("   в т.ч. плечо 1-го этапа (Отвал / ОФ / ДСК-СМС) (ссылка)", "км", f"={cell_map['DIST_WASTE']}", f"={cell_map['DIST_OF']}", f"={cell_map['DIST_TO_SMS']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_d_sms_of_l = add_row("   в т.ч. плечо 3б: ДСК-СМС -> ОФ (предконцентрат) (ссылка)", "км", 0, 0, f"={cell_map['DIST_SMS_OF']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_d_sms_tw_l = add_row("   в т.ч. плечо 3в: ДСК-СМС -> Отвал хвостов (ссылка)", "км", 0, 0, f"={cell_map['DIST_SMS_WASTE']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_d_sms_tg_l = add_row("   в т.ч. плечо 3г: ДСК-СМС -> Штабель щебня (ссылка)", "км", 0, 0, f"={cell_map['DIST_SMS_GRAVEL']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_ctr_l = add_row("   в т.ч. УДР транспортировки горной массы (ссылка)", "руб./т*км", f"={cell_map['OPEX_TR']}", f"={cell_map['OPEX_TR']}", f"={cell_map['OPEX_TR']}", "Импорт из Inputs", fmt_rub, "SUB")

    ws_calc[f"D{r_opex_tr}"] = f"=D{r_wet_mined}*D{r_d_main_l}*D{r_ctr_l}"
    ws_calc[f"E{r_opex_tr}"] = f"=E{r_wet_mined}*E{r_d_main_l}*E{r_ctr_l}"
    ws_calc[f"F{r_opex_tr}"] = f"=F{r_wet_mined}*F{r_d_main_l}*F{r_ctr_l} + (F{r_of_mass}/(1-F{r_moist_l}))*F{r_d_sms_of_l}*F{r_ctr_l} + (F{r_tail_mass}/(1-F{r_moist_l}))*IF(F{r_opt_l}=1,F{r_d_sms_tw_l},F{r_d_sms_tg_l})*F{r_ctr_l}"

    r_opex_crush = add_row("Дробление и сухая магнитная сепарация (ДСК-СМС)", "руб.", 0, 0, 0, "В Сц.3 начисляется на весь добытый поток питания ДСК-СМС", fmt_rub)
    r_ccr_l = add_row("   в т.ч. УДР дробления / ДСК-СМС на 1 т сухой руды (ссылка)", "руб./т", 0, f"={cell_map['OPEX_CRUSH_OF']}", f"={cell_map['OPEX_SMS']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"E{r_opex_crush}"] = f"=E{r_dry_mined}*E{r_ccr_l}"
    ws_calc[f"F{r_opex_crush}"] = f"=F{r_dry_mined}*F{r_ccr_l}"

    r_opex_of = add_row("Обогащение на ОФ (измельчение, флотация, доводка)", "руб.", 0, 0, 0, "Начисляется на 0,965 т сухой руды на входе ОФ", fmt_rub)
    r_cof_fe_l = add_row("   в т.ч. УДР переработки ЖРК на ОФ (ссылка)", "руб./т", 0, f"={cell_map['OPEX_OF_FE']}", f"={cell_map['OPEX_OF_FE']}", "Импорт из Inputs", fmt_rub, "SUB")
    r_cof_p_l = add_row("   в т.ч. УДР переработки АПК на ОФ (ссылка)", "руб./т", 0, f"={cell_map['OPEX_OF_P']}", f"={cell_map['OPEX_OF_P']}", "Импорт из Inputs", fmt_rub, "SUB")
    r_cof_zr_l = add_row("   в т.ч. УДР переработки БК на ОФ (ссылка)", "руб./т", 0, f"={cell_map['OPEX_OF_ZR']}", f"={cell_map['OPEX_OF_ZR']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"E{r_opex_of}"] = f"=E{r_of_mass}*(E{r_cof_fe_l}+E{r_cof_p_l}+E{r_cof_zr_l})"
    ws_calc[f"F{r_opex_of}"] = f"=F{r_of_mass}*(F{r_cof_fe_l}+F{r_cof_p_l}+F{r_cof_zr_l})"

    r_tot_opex = add_row("ИТОГО ОПЕРАЦИОННЫЕ ЗАТРАТЫ (OPEX)", "руб.", f"=D{r_opex_ndpi}+D{r_opex_bvr}+D{r_opex_exc}+D{r_opex_tr}+D{r_opex_crush}+D{r_opex_of}", f"=E{r_opex_ndpi}+E{r_opex_bvr}+E{r_opex_exc}+E{r_opex_tr}+E{r_opex_crush}+E{r_opex_of}", f"=F{r_opex_ndpi}+F{r_opex_bvr}+F{r_opex_exc}+F{r_opex_tr}+F{r_opex_crush}+F{r_opex_of}", "Полные затраты включая НДПИ, БВР и экскавацию", fmt_rub, "TOTAL")

    # --- СЕКЦИЯ 5: ЭКОНОМИЧЕСКИЙ ЭФФЕКТ И ВЫБОР АЛГОРИТМА ---
    add_section("5. ЭКОНОМИЧЕСКИЙ ЭФФЕКТ И ВЫБОР АЛГОРИТМА")
    r_net_profit = add_row("Маржинальная прибыль (Net Value = Доход - OPEX)", "руб./т ОФ", f"=D{r_gross_rev}-D{r_tot_opex}", f"=E{r_gross_rev}-E{r_tot_opex}", f"=F{r_gross_rev}-F{r_tot_opex}", "Критерий максимизации стоимости на 1 т питания ОФ", fmt_rub, "TOTAL")
    r_diff_waste = add_row("Прирост прибыли относительно Отвала (Сц.1)", "руб./т ОФ", 0, f"=E{r_net_profit}-$D${r_net_profit}", f"=F{r_net_profit}-$D${r_net_profit}", "Если > 0, выгоднее перерабатывать, чем везти в отвал", fmt_rub)
    r_diff_sms = add_row("Эффект от рудоразборки на ДСК-СМС против прямой ОФ", "руб./т ОФ", 0, 0, f"=F{r_net_profit}-E{r_net_profit}", "Сальдо прироста выручки и доп. затрат на добычу", fmt_rub)
    r_decision = add_row("ОПТИМАЛЬНОЕ НАПРАВЛЕНИЕ ДЛЯ БАЗОВОГО БЛОКА", "решение", f'=IF(MAX(D{r_net_profit}:F{r_net_profit})=D{r_net_profit},"1. В ОТВАЛ",IF(MAX(D{r_net_profit}:F{r_net_profit})=E{r_net_profit},"2. НАПРЯМУЮ НА ОФ","3. ЧЕРЕЗ ДСК-СМС"))', "", "", "Автоматический выбор максимального Net Value", "@", "TOTAL")

    ws_calc.merge_cells(start_row=r_decision, start_column=4, end_row=r_decision, end_column=6)
    ws_calc.cell(row=r_decision, column=4).fill = fill_sms

    ws_calc.column_dimensions["A"].width = 3
    ws_calc.column_dimensions["B"].width = 58
    ws_calc.column_dimensions["C"].width = 11
    ws_calc.column_dimensions["D"].width = 18
    ws_calc.column_dimensions["E"].width = 20
    ws_calc.column_dimensions["F"].width = 22
    ws_calc.column_dimensions["G"].width = 50

    # =========================================================================
    # ЛИСТ 3: CUTOFF_MATRICES_VISUAL (6 МАТРИЦ, ДОП. ДОБЫЧА В СЦ.3, ПОДСВЕТКА)
    # =========================================================================
    for sv in ws_mat.views.sheetView:
        sv.showGridLines = True

    ws_mat["B2"] = "МАТРИЧНЫЙ АНАЛИЗ CUT-OFF GRADE: ЧУВСТВИТЕЛЬНОСТЬ НА 1 Т ПИТАНИЯ ОФ"
    ws_mat["B2"].font = f_title

    ws_mat["B3"] = "В Сценарии 3 учтены доп. затраты на добычу (НДПИ, БВР, экскавация) и транспорт для получения 1 т питания ОФ. Красная рамка — базовый блок."
    ws_mat["B3"].font = f_note

    fe_steps = (0.03, 0.05, 0.07, 0.09, 0.10, 0.11, 0.12, 0.13, 0.14, 0.15, 0.16, 0.17, 0.19, 0.21, 0.23, 0.25)
    p_steps = (0.010, 0.015, 0.020, 0.025, 0.030, 0.035, 0.040, 0.045, 0.050, 0.055, 0.060, 0.068, 0.076, 0.084, 0.092)

    v_moist = cell_map["MOISTURE"]
    v_fe_conv = cell_map["FEO_CONV"]
    v_dil = cell_map["DILUTION"]
    v_zr = cell_map["ZRO2_BASE"]

    of_fe_a = cell_map["OF_FE_A"]
    of_fe_b = cell_map["OF_FE_B"]
    of_p_a = cell_map["OF_P_A"]
    of_p_b = cell_map["OF_P_B"]
    of_zr_a = cell_map["OF_ZR_A"]
    of_zr_b = cell_map["OF_ZR_B"]

    max_rec_fe = cell_map["MAX_REC_FE"]
    max_rec_p = cell_map["MAX_REC_P"]
    max_rec_zr = cell_map["MAX_REC_ZR"]

    g_fe_c = cell_map["GRADE_FE_CONC"]
    g_p_c = cell_map["GRADE_P_CONC"]
    g_zr_c = cell_map["GRADE_ZR_CONC"]

    p_fe = cell_map["PRICE_FE"]
    p_p = cell_map["PRICE_P"]
    p_zr = cell_map["PRICE_ZR"]
    fx = cell_map["FX_RATE"]

    sms_ym_a = cell_map["SMS_YM_A"]
    sms_ym_b = cell_map["SMS_YM_B"]
    sms_yc_b = cell_map["SMS_YC_B"]
    sms_sub_sh = cell_map["SMS_SUB_SHARE"]
    sms_rfe_ma = cell_map["SMS_RFE_M_A"]
    sms_rfe_mb = cell_map["SMS_RFE_M_B"]
    max_sms_rfem = cell_map["MAX_SMS_RFE_M"]
    sms_rfe_nm_sh = cell_map["SMS_RFE_NM_SHARE"]
    sms_rp_a = cell_map["SMS_RP_A"]
    sms_rp_b1 = cell_map["SMS_RP_B1"]
    sms_rp_b2 = cell_map["SMS_RP_B2"]
    sms_r_zr = cell_map["SMS_REC_ZR"]
    max_sms_y = cell_map["MAX_SMS_YIELD"]

    sms_opt = cell_map["SMS_TAIL_OPT"]
    gravel_v = cell_map["GRAVEL_VAL"]
    d_to_sms = cell_map["DIST_TO_SMS"]
    d_sms_of = cell_map["DIST_SMS_OF"]
    d_sms_w = cell_map["DIST_SMS_WASTE"]
    d_sms_g = cell_map["DIST_SMS_GRAVEL"]

    c_ndpi = cell_map["OPEX_NDPI"]
    c_bvr = cell_map["OPEX_BVR"]
    c_exc = cell_map["OPEX_EXC"]
    c_tr = cell_map["OPEX_TR"]
    c_sms = cell_map["OPEX_SMS"]
    c_of_fe = cell_map["OPEX_OF_FE"]
    c_of_p = cell_map["OPEX_OF_P"]
    c_of_zr = cell_map["OPEX_OF_ZR"]

    fe_base_ref = cell_map["FE_BASE"]
    p_base_ref = cell_map["P2O5_BASE"]
    left_col = 2
    gap_col = left_col + len(p_steps) + 1
    right_col = gap_col + 1

    r1_start = 7
    r1_end = r1_start + len(fe_steps) - 1
    row_m2 = r1_end + 3
    row_m3 = row_m2 + len(fe_steps) + 4

    def build_matrix(start_row, start_col, title, mode):
        first_data_col = start_col + 1
        last_data_col = start_col + len(p_steps)
        fe_col_let = get_column_letter(start_col)

        ws_mat.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=last_data_col)
        hdr = ws_mat.cell(row=start_row, column=start_col, value=title)
        hdr.font = f_sec_hdr
        hdr.fill = fill_sec_hdr
        ws_mat.row_dimensions[start_row].height = 22

        r_col_hdr = start_row + 1
        corner = ws_mat.cell(row=r_col_hdr, column=start_col, value="Fe \\ P2O5")
        corner.font = f_tbl_hdr
        corner.fill = fill_tbl_hdr
        corner.border = border_data
        corner.alignment = Alignment(horizontal="center")

        for j, p_val in enumerate(p_steps):
            c = ws_mat.cell(row=r_col_hdr, column=first_data_col + j, value=p_val)
            c.font = f_input
            c.fill = fill_tbl_hdr
            c.border = border_data
            c.number_format = fmt_pct
            c.alignment = Alignment(horizontal="center")

        for i, fe_val in enumerate(fe_steps):
            curr_row = r_col_hdr + 1 + i
            c_fe = ws_mat.cell(row=curr_row, column=start_col, value=fe_val)
            c_fe.font = f_input
            c_fe.fill = fill_tbl_hdr
            c_fe.border = border_data
            c_fe.number_format = fmt_pct
            c_fe.alignment = Alignment(horizontal="center")

            fe_ref = f"${fe_col_let}{curr_row}"
            for j in range(len(p_steps)):
                col_let = get_column_letter(first_data_col + j)
                p_ref = f"{col_let}${r_col_hdr}"

                fe_raw = f"({fe_ref}*{v_fe_conv}*(1-{v_dil}))"
                p_raw = f"({p_ref}*(1-{v_dil}))"
                zr_raw = f"({v_zr}*(1-{v_dil}))"

                rec_fe_2 = f"MAX(0,MIN({max_rec_fe},({of_fe_a}-{of_fe_b}/({fe_raw}*100))/100))"
                rec_p_2 = f"MAX(0,MIN({max_rec_p},({of_p_a}-{of_p_b}/({p_raw}*100))/100))"
                rec_zr_2 = f"MAX(0,MIN({max_rec_zr},({of_zr_a}-{of_zr_b}/({zr_raw}*100))/100))"

                rev_2 = f"((1-{v_moist})*({fe_raw}*{rec_fe_2}/{g_fe_c}*{p_fe} + {p_raw}*{rec_p_2}/{g_p_c}*{p_p} + {zr_raw}*{rec_zr_2}/{g_zr_c}*{p_zr})*{fx})"
                profit_sc2 = f"({rev_2} - Unit_Economics_Calc!$E${r_tot_opex})"

                ym_3 = f"MAX(0,MIN(1,({sms_ym_a}*{fe_raw}*100-{sms_ym_b})/100))"
                ynm_3 = f"(1-{ym_3})"
                yc_3 = f"MAX(0.01,MIN({max_sms_y},({sms_ym_a}*{fe_raw}*100-{sms_yc_b}+({ynm_3}*100*{sms_sub_sh}))/100))"

                rfe_m_3 = f"MAX(0,MIN({max_sms_rfem},({sms_rfe_ma}*{fe_raw}*100-{sms_rfe_mb})/100))"
                rfe_tot_3 = f"MIN(1,{rfe_m_3}+(1-{rfe_m_3})*{sms_rfe_nm_sh})"
                rp_tot_3 = f"MAX(0,MIN(1,({sms_rp_a}*{p_raw}*100+{sms_rp_b1}+{sms_rp_b2})/100))"

                fe_of_3 = f"({fe_raw}*{rfe_tot_3}/{yc_3})"
                p_of_3 = f"({p_raw}*{rp_tot_3}/{yc_3})"
                zr_of_3 = f"({zr_raw}*{sms_r_zr}/{yc_3})"

                rec_fe_3 = f"MAX(0,MIN({max_rec_fe},({of_fe_a}-{of_fe_b}/({fe_of_3}*100))/100))"
                rec_p_3 = f"MAX(0,MIN({max_rec_p},({of_p_a}-{of_p_b}/({p_of_3}*100))/100))"
                rec_zr_3 = f"MAX(0,MIN({max_rec_zr},({of_zr_a}-{of_zr_b}/({zr_of_3}*100))/100))"

                rev_conc_3 = f"((1-{v_moist})*({fe_of_3}*{rec_fe_3}/{g_fe_c}*{p_fe} + {p_of_3}*{rec_p_3}/{g_p_c}*{p_p} + {zr_of_3}*{rec_zr_3}/{g_zr_c}*{p_zr})*{fx})"
                rev_gravel_3 = f"IF({sms_opt}=2,((1/{yc_3})-1)*{gravel_v},0)"
                opex_3 = f"(((1/{yc_3})-1)*({c_ndpi}+{c_bvr}+{c_exc}) + (1/{yc_3})*{d_to_sms}*{c_tr} + {d_sms_of}*{c_tr} + ((1/{yc_3})-1)*IF({sms_opt}=1,{d_sms_w},{d_sms_g})*{c_tr} + (1-{v_moist})*((1/{yc_3})*{c_sms} + ({c_of_fe}+{c_of_p}+{c_of_zr})))"
                profit_sc3 = f"({rev_conc_3} + {rev_gravel_3} - {opex_3})"

                profit_sc1 = f"Unit_Economics_Calc!$D${r_net_profit}"

                r_col_let = get_column_letter(right_col + 1 + j)
                c_sc1 = f"{r_col_let}{7 + i}"
                c_sc2 = f"{r_col_let}{7 + len(fe_steps) + 4 + i}"
                c_sc3 = f"{r_col_let}{7 + (len(fe_steps) + 4)*2 + i}"

                if mode == "DECISION":
                    formula = f"=IF(MAX({c_sc1},{c_sc2},{c_sc3})={c_sc1},1,IF(MAX({c_sc1},{c_sc2},{c_sc3})={c_sc2},2,3))"
                    num_f = '[=1]"ОТВАЛ";[=2]"ОФ";"СМС"'
                elif mode == "NET_MAX":
                    formula = f"=MAX({c_sc1},{c_sc2},{c_sc3})"
                    num_f = fmt_rub
                elif mode == "SMS_VS_OF":
                    formula = f"={c_sc3}-{c_sc2}"
                    num_f = fmt_rub
                elif mode == "SC1_ONLY":
                    formula = f"={profit_sc1}"
                    num_f = fmt_rub
                elif mode == "SC2_ONLY":
                    formula = f"={profit_sc2}"
                    num_f = fmt_rub
                elif mode == "SC3_ONLY":
                    formula = f"={profit_sc3}"
                    num_f = fmt_rub

                cell = ws_mat.cell(row=curr_row, column=first_data_col + j, value=formula)
                cell.font = f_formula
                cell.border = border_data
                cell.number_format = num_f
                if mode == "DECISION":
                    cell.alignment = Alignment(horizontal="center")

        r_data_start = r_col_hdr + 1
        r_data_end = r_col_hdr + len(fe_steps)
        m_range = f"{get_column_letter(first_data_col)}{r_data_start}:{get_column_letter(last_data_col)}{r_data_end}"
        return r_data_start, r_data_end, m_range
    _, _, range_m4_1 = build_matrix(5, right_col, "МАТРИЦА 4_1: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ — СЦЕНАРИЙ 1: В ОТВАЛ (РУБ./Т)", "SC1_ONLY")
    ws_mat.conditional_formatting.add(range_m4_1, CellIsRule(operator="lessThan", formula=("0",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m4_1, CellIsRule(operator="greaterThanOrEqual", formula=("0",), fill=fill_sms, font=Font(bold=True, color="006100")))

    _, _, range_m4_2 = build_matrix(row_m2, right_col, "МАТРИЦА 4_2: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ — СЦЕНАРИЙ 2: ПРЯМАЯ ОФ (РУБ./Т ОФ)", "SC2_ONLY")
    ws_mat.conditional_formatting.add(range_m4_2, CellIsRule(operator="lessThan", formula=("0",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m4_2, CellIsRule(operator="greaterThanOrEqual", formula=("0",), fill=fill_of, font=Font(bold=True, color="1F4E78")))

    _, _, range_m4_3 = build_matrix(row_m3, right_col, "МАТРИЦА 4_3: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ — СЦЕНАРИЙ 3: ДСК-СМС -> ОФ (РУБ./Т ОФ)", "SC3_ONLY")
    ws_mat.conditional_formatting.add(range_m4_3, CellIsRule(operator="lessThan", formula=("0",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m4_3, CellIsRule(operator="greaterThanOrEqual", formula=("0",), fill=fill_sms, font=Font(bold=True, color="006100")))

    r1_start, r1_end, range_m1 = build_matrix(5, left_col, "МАТРИЦА 1: КАРТА ПРИНЯТИЯ РЕШЕНИЙ (CUT-OFF GRADE: ОТВАЛ / ОФ / ДСК-СМС)", "DECISION")
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=("1",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=("2",), fill=fill_of, font=Font(bold=True, color="1F4E78")))
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=("3",), fill=fill_sms, font=Font(bold=True, color="006100")))

    r2_start, r2_end, range_m2 = build_matrix(row_m2, left_col, "МАТРИЦА 2: МАКСИМАЛЬНАЯ МАРЖИНАЛЬНАЯ ПРИБЫЛЬ НА 1 Т ПИТАНИЯ ОФ (РУБ./Т ОФ)", "NET_MAX")
    ws_mat.conditional_formatting.add(range_m2, ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="num", mid_value=0, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"
    ))

    r3_start, r3_end, range_m3 = build_matrix(row_m3, left_col, "МАТРИЦА 3: ЭФФЕКТ ОТ РУДОРАЗБОРКИ НА ДСК-СМС ПРОТИВ ПРЯМОЙ ОФ (РУБ./Т ОФ)", "SMS_VS_OF")
    ws_mat.conditional_formatting.add(range_m3, ColorScaleRule(
        start_type="min", start_color="D9E1F2",
        mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="max", end_color="70AD47"
    ))

    for m_start_r, m_start_c, m_rng in (
        (5, right_col, range_m4_1),
        (row_m2, right_col, range_m4_2),
        (row_m3, right_col, range_m4_3),
        (5, left_col, range_m1),
        (row_m2, left_col, range_m2),
        (row_m3, left_col, range_m3),
    ):
        c_fe_let = get_column_letter(m_start_c)
        c_p_let = get_column_letter(m_start_c + 1)
        r_hdr = m_start_r + 1
        r_dat = m_start_r + 2
        hl_f = f"AND(ROUND(${c_fe_let}{r_dat},4)=ROUND({fe_base_ref},4),ROUND({c_p_let}${r_hdr},4)=ROUND({p_base_ref},4))"
        ws_mat.conditional_formatting.add(
            m_rng,
            FormulaRule(
                formula=(hl_f,),
                stopIfTrue=False,
                border=border_highlight,
                font=Font(name=font_family, size=10, bold=True, underline="single", color="9C0006")
            )
        )

    ws_mat.column_dimensions["A"].width = 3
    ws_mat.column_dimensions[get_column_letter(left_col)].width = 13
    for j in range(len(p_steps)):
        ws_mat.column_dimensions[get_column_letter(left_col + 1 + j)].width = 11.5

    ws_mat.column_dimensions[get_column_letter(gap_col)].width = 4
    ws_mat.column_dimensions[get_column_letter(right_col)].width = 13
    for j in range(len(p_steps)):
        ws_mat.column_dimensions[get_column_letter(right_col + 1 + j)].width = 11.5

    far_right_col = right_col + len(p_steps) + 2
    ws_mat.column_dimensions[get_column_letter(far_right_col - 1)].width = 4

    chart = LineChart()
    chart.title = "Зависимость максимальной прибыли (руб./т ОФ) от Fe при разном P2O5"
    chart.style = 10
    chart.y_axis.title = "Прибыль, руб./т ОФ"
    chart.x_axis.title = "Содержание Fe в исходной руде, %"
    chart.width = 22
    chart.height = 12

    data_ref = Reference(ws_mat, min_col=left_col + 1, min_row=r2_start - 1, max_col=left_col + len(p_steps), max_row=r2_end)
    cats_ref = Reference(ws_mat, min_col=left_col, min_row=r2_start, max_row=r2_end)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    ws_mat.add_chart(chart, f"{get_column_letter(far_right_col)}5")

    stat_col_start = far_right_col
    stat_col_val = stat_col_start + 1
    stat_row = 31

    ws_mat.merge_cells(start_row=stat_row, start_column=stat_col_start, end_row=stat_row, end_column=stat_col_val)
    s_hdr = ws_mat.cell(row=stat_row, column=stat_col_start, value="СТАТИСТИКА ЗОНЫ РУДОРАЗБОРКИ (ДСК-СМС)")
    s_hdr.font = f_sec_hdr
    s_hdr.fill = fill_sec_hdr
    ws_mat.row_dimensions[stat_row].height = 22

    fe_col_range = f"${get_column_letter(left_col)}${r1_start}:${get_column_letter(left_col)}${r1_end}"
    p_row_range = f"${get_column_letter(left_col + 1)}${r1_start - 1}:${get_column_letter(left_col + len(p_steps))}${r1_start - 1}"

    stats_rows = (
        ("Количество сочетаний (ячеек) для СМС", f"=COUNTIF({range_m1},3)", fmt_int),
        ("Доля матрицы, эффективная для СМС", f"=COUNTIF({range_m1},3)/COUNTA({range_m1})", fmt_pct),
        ("Медиана содержания Fe в руде (зона СМС)", f"=IFERROR(MEDIAN(IF({range_m1}=3,{fe_col_range})),0)", fmt_pct),
        ("Медиана содержания P2O5 в руде (зона СМС)", f"=IFERROR(MEDIAN(IF({range_m1}=3,{p_row_range})),0)", fmt_pct),
        ("Среднее содержание Fe в руде (зона СМС)", f"=IFERROR(SUMPRODUCT(({range_m1}=3)*{fe_col_range})/COUNTIF({range_m1},3),0)", fmt_pct),
        ("Среднее содержание P2O5 в руде (зона СМС)", f"=IFERROR(SUMPRODUCT(({range_m1}=3)*{p_row_range})/COUNTIF({range_m1},3),0)", fmt_pct),
    )

    for idx_s, (s_name, s_formula, s_fmt) in enumerate(stats_rows):
        curr_s_row = stat_row + 1 + idx_s
        c_lbl = ws_mat.cell(row=curr_s_row, column=stat_col_start, value=s_name)
        c_lbl.font = f_formula
        c_lbl.border = border_data

        c_v = ws_mat.cell(row=curr_s_row, column=stat_col_val, value=s_formula)
        c_v.font = f_total
        c_v.fill = fill_sms
        c_v.border = border_data
        c_v.number_format = s_fmt
        c_v.alignment = Alignment(horizontal="right")

    ws_mat.column_dimensions[get_column_letter(stat_col_start)].width = 38
    ws_mat.column_dimensions[get_column_letter(stat_col_val)].width = 16

    wb.save(filename)
    print(f"Файл ФЭМ (версия 3 — Mill-Constrained) успешно сгенерирован: {filename}")


if __name__ == "__main__":
    create_cut_off_fem()