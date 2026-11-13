import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart import LineChart, Reference


def create_cut_off_fem(filename="Cut_Off_Grade_MAR_FEM.xlsx"):
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
    fmt_pct = "0.00%"
    fmt_rub = '#,##0.00" ₽"'
    fmt_usd = '$#,##0.00'

    # =========================================================================
    # ЛИСТ 1: INPUTS (ИСХОДНЫЕ ДАННЫЕ И ЛЕГЕНДА)
    # =========================================================================
    for sv in ws_in.views.sheetView:
        sv.showGridLines = True

    ws_in["B2"] = "ФЭМ РАСЧЕТА CUT-OFF GRADE (МАР): ИСХОДНЫЕ ДАННЫЕ"
    ws_in["B2"].font = f_title

    ws_in["F2"] = "ЛЕГЕНДА ФОРМАТИРОВАНИЯ (СТАНДАРТ BIG4 / McKinsey)"
    ws_in["F2"].font = f_tbl_hdr

    legend_items = (
        ("F3", "Ручной ввод (Inputs)", f_input, fill_input, "Изменяемые параметры сценария"),
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
        (None, "Базовый объем анализируемого блока (сырая масса)", 1.0, "т", fmt_dec2, "Расчет на 1 т извлекаемой горной массы", "VOL_BASE"),
        (None, "Содержание влаги в исходной руде", 0.035, "%", fmt_pct, "Пересчет сырой массы в сухую", "MOISTURE"),
        (None, "Базовое содержание Fe (общее / из FeO) в руде", 0.030, "%", fmt_pct, "Базовое значение для Calc и подсветки матриц", "FE_BASE"),
        (None, "Коэффициент пересчета FeO в Fe (если ввод в FeO)", 1.000, "доля", fmt_dec2, "1.00 если ввод сразу в Fe; 0.777 если ввод в FeO", "FEO_CONV"),
        (None, "Базовое содержание P2O5 в руде", 0.010, "%", fmt_pct, "Базовое значение для Calc и подсветки матриц", "P2O5_BASE"),
        (None, "Содержание ZrO2 в руде (попутный компонент)", 0.0015, "%", fmt_pct, "Учитывается при переработке на ОФ (БК)", "ZRO2_BASE"),
        (None, "Разубоживание при выемке в карьере (обнулено)", 0.000, "%", fmt_pct, "Установлено 0% по условию задачи", "DILUTION"),

        ("2. МАКРОЭКОНОМИКА И ЦЕНЫ РЕАЛИЗАЦИИ (FCA/FOB)", None, None, None, None, None, None),
        (None, "Курс доллара США", 92.50, "руб./$", fmt_dec2, "Курс пересчета валютной выручки", "FX_RATE"),
        (None, "Цена ЖРК (железорудный концентрат)", 115.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_FE"),
        (None, "Цена АПК (апатитовый концентрат)", 165.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_P"),
        (None, "Цена БК (бадделеитовый концентрат)", 3200.00, "$/т", fmt_usd, "Рыночная цена за 1 т сухого концентрата", "PRICE_ZR"),

        ("3. КАЧЕСТВО КОНЦЕНТРАТОВ И ЗАВИСИМОСТЬ ИЗВЛЕЧЕНИЯ (ОФ)", None, None, None, None, None, None),
        (None, "Содержание Fe в товарном ЖРК", 0.655, "%", fmt_pct, "Базисное качество ЖРК", "GRADE_FE_CONC"),
        (None, "Содержание P2O5 в товарном АПК", 0.390, "%", fmt_pct, "Базисное качество АПК", "GRADE_P_CONC"),
        (None, "Содержание ZrO2 в товарном БК", 0.985, "%", fmt_pct, "Базисное качество БК", "GRADE_ZR_CONC"),
        (None, "Коэффициент 'a' извлечения Fe (наклон прямой)", 1.6961, "коэф.", fmt_dec2, "Формула технологов: (1.6961 * Fe% * 100 + 47.74)/100", "REC_FE_A"),
        (None, "Свободный член 'b' извлечения Fe (в долях)", 0.4774, "доля", fmt_pct, "47.74% (базовый уровень регрессии)", "REC_FE_B"),
        (None, "Коэффициент 'a' извлечения P2O5 (наклон прямой)", 1.4892, "коэф.", fmt_dec2, "Формула технологов: (1.4892 * P2O5% * 100 + 64.232)/100", "REC_P_A"),
        (None, "Свободный член 'b' извлечения P2O5 (в долях)", 0.64232, "доля", fmt_pct, "64.232% (базовый уровень регрессии)", "REC_P_B"),
        (None, "Сквозное извлечение ZrO2 в БК (на ОФ)", 0.550, "%", fmt_pct, "Константа (по условию задачи)", "REC_ZR"),
        (None, "Максимальный предел извлечения Fe на ОФ (потолок)", 0.950, "%", fmt_pct, "Физическое ограничение регрессии для богатых руд", "MAX_REC_FE"),
        (None, "Максимальный предел извлечения P2O5 на ОФ (потолок)", 0.920, "%", fmt_pct, "Физическое ограничение регрессии для богатых руд", "MAX_REC_P"),

        ("4. ПАРАМЕТРЫ РУДОРАЗБОРКИ НА ДСК-СМС (2 СТАДИИ)", None, None, None, None, None, None),
        (None, "Выход предконцентрата после 2-х стадий СМС", 0.7236, "%", fmt_pct, "Доля массы руды, направляемой после СМС на ОФ", "SMS_YIELD"),
        (None, "Извлечение Fe в предконцентрат СМС", 0.940, "%", fmt_pct, "Доля Fe, переходящая в предконцентрат (баланс)", "SMS_REC_FE"),
        (None, "Извлечение P2O5 в предконцентрат СМС", 0.880, "%", fmt_pct, "Доля апатита, переходящая в предконцентрат (баланс)", "SMS_REC_P"),
        (None, "Извлечение ZrO2 в предконцентрат СМС", 0.850, "%", fmt_pct, "Доля бадделеита, уходящая на ОФ", "SMS_REC_ZR"),
        (None, "Опция хвостов СМС (1 = в отвал; 2 = щебень для дорог)", 1, "код", fmt_int, "1 - вывоз в отвал; 2 - штабель щебня + экономия", "SMS_TAIL_OPT"),
        (None, "Экономия / ценность щебня из хвостов СМС", 350.00, "руб./т", fmt_rub, "Альтернативная стоимость покупки щебня для дорог", "GRAVEL_VAL"),

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
        (None, "НДПИ и БВР/Экскавация (справочно, не в Cut-off)", 310.00, "руб./т", fmt_rub, "Sunk costs: одинаковы для всех 3-х сценариев в контуре", "SUNK_COST")
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
    ws_in.column_dimensions["B"].width = 52
    ws_in.column_dimensions["C"].width = 16
    ws_in.column_dimensions["D"].width = 14
    ws_in.column_dimensions["E"].width = 55
    ws_in.column_dimensions["F"].width = 32
    ws_in.column_dimensions["G"].width = 35
    # =========================================================================
    # ЛИСТ 2: UNIT_ECONOMICS_CALC (ДЕТАЛЬНЫЙ РАСЧЕТ С ГРУППИРОВКОЙ СТРОК)
    # =========================================================================
    for sv in ws_calc.views.sheetView:
        sv.showGridLines = True

    ws_calc.sheet_properties.outlinePr.summaryBelow = False

    ws_calc["B2"] = "МОДЕЛЬ UNIT-ЭКОНОМИКИ 1 ТОННЫ ПОРОДЫ ПО ТРЕМ АЛГОРИТМАМ (СЦЕНАРИЯМ)"
    ws_calc["B2"].font = f_title

    ws_calc["B3"] = "Нажмите '+' слева от номеров строк, чтобы развернуть слинкованные исходные параметры с листа Inputs."
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
    add_section("1. МАТЕРИАЛЬНЫЙ БАЛАНС И КАЧЕСТВО ПИТАНИЯ ОФ")
    r_wet = add_row("Сырая горная масса из карьера", "т", f"={cell_map['VOL_BASE']}", f"={cell_map['VOL_BASE']}", f"={cell_map['VOL_BASE']}", "Линк: Базовая 1 тонна из карьера", fmt_dec2)
    r_dry = add_row("Сухая масса горной породы из карьера", "т", 0, 0, 0, "Расчет по строкам текущего листа", fmt_dec2)
    r_moist_l = add_row("   в т.ч. влажность исходной руды (ссылка)", "%", f"={cell_map['MOISTURE']}", f"={cell_map['MOISTURE']}", f"={cell_map['MOISTURE']}", "Импорт из Inputs", fmt_pct, "SUB")
    for col_let in ("D", "E", "F"):
        ws_calc[f"{col_let}{r_dry}"] = f"={col_let}{r_wet}*(1-{col_let}{r_moist_l})"

    r_of_mass = add_row("Масса сухой руды, поступающей на ОФ", "т", 0, f"=E{r_dry}", 0, "В Сц.3 на ОФ идет только предконцентрат СМС", fmt_dec2)
    r_sms_y_l = add_row("   в т.ч. выход предконцентрата ДСК-СМС (ссылка)", "%", 0, 1.0, f"={cell_map['SMS_YIELD']}", "Импорт выхода предконцентрата из Inputs", fmt_pct, "SUB")
    ws_calc[f"F{r_of_mass}"] = f"=F{r_dry}*F{r_sms_y_l}"

    r_tail_mass = add_row("Масса хвостов СМС (выход щебня / отвала)", "т", 0, 0, f"=F{r_dry}-F{r_of_mass}", "Разница между сухой рудой и предконцентратом", fmt_dec2)

    r_fe_feed = add_row("Содержание Fe в питании ОФ", "%", 0, 0, 0, "Балансовый прирост содержания Fe после СМС", fmt_pct)
    r_fe_b_l = add_row("   в т.ч. базовое содержание Fe в руде (ссылка)", "%", 0, f"={cell_map['FE_BASE']}", f"={cell_map['FE_BASE']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_fe_c_l = add_row("   в т.ч. коэффициент пересчета FeO -> Fe (ссылка)", "доля", 0, f"={cell_map['FEO_CONV']}", f"={cell_map['FEO_CONV']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_dil_l = add_row("   в т.ч. разубоживание при добыче (ссылка)", "%", 0, f"={cell_map['DILUTION']}", f"={cell_map['DILUTION']}", "Импорт из Inputs (обнулено = 0%)", fmt_pct, "SUB")
    r_sms_rfe_l = add_row("   в т.ч. извлечение Fe в предконцентрат СМС (ссылка)", "%", 0, 0, f"={cell_map['SMS_REC_FE']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_fe_feed}"] = f"=E{r_fe_b_l}*E{r_fe_c_l}*(1-E{r_dil_l})"
    ws_calc[f"F{r_fe_feed}"] = f"=(E{r_fe_feed}*F{r_sms_rfe_l})/F{r_sms_y_l}"

    r_p_feed = add_row("Содержание P2O5 в питании ОФ", "%", 0, 0, 0, "Балансовый прирост содержания P2O5 после СМС", fmt_pct)
    r_p_b_l = add_row("   в т.ч. базовое содержание P2O5 в руде (ссылка)", "%", 0, f"={cell_map['P2O5_BASE']}", f"={cell_map['P2O5_BASE']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_sms_rp_l = add_row("   в т.ч. извлечение P2O5 в предконцентрат СМС (ссылка)", "%", 0, 0, f"={cell_map['SMS_REC_P']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_p_feed}"] = f"=E{r_p_b_l}*(1-E{r_dil_l})"
    ws_calc[f"F{r_p_feed}"] = f"=(E{r_p_feed}*F{r_sms_rp_l})/F{r_sms_y_l}"

    r_zr_feed = add_row("Содержание ZrO2 в питании ОФ", "%", 0, 0, 0, "Балансовый прирост содержания ZrO2 после СМС", fmt_pct)
    r_zr_b_l = add_row("   в т.ч. базовое содержание ZrO2 в руде (ссылка)", "%", 0, f"={cell_map['ZRO2_BASE']}", f"={cell_map['ZRO2_BASE']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_sms_rzr_l = add_row("   в т.ч. извлечение ZrO2 в предконцентрат СМС (ссылка)", "%", 0, 0, f"={cell_map['SMS_REC_ZR']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_zr_feed}"] = f"=E{r_zr_b_l}*(1-E{r_dil_l})"
    ws_calc[f"F{r_zr_feed}"] = f"=(E{r_zr_feed}*F{r_sms_rzr_l})/F{r_sms_y_l}"

    # --- СЕКЦИЯ 2: ВЫПУСК ТОВАРНЫХ КОНЦЕНТРАТОВ ---
    add_section("2. ВЫПУСК ТОВАРНЫХ КОНЦЕНТРАТОВ (СУХОЙ ВЕС)")
    r_rec_fe_calc = add_row("Расчетное сквозное извлечение Fe на ОФ", "%", 0, 0, 0, "Формула технологов с ограничением потолка", fmt_pct)
    r_rec_fe_a_l = add_row("   в т.ч. коэффициент 'a' регрессии Fe (ссылка)", "коэф.", 0, f"={cell_map['REC_FE_A']}", f"={cell_map['REC_FE_A']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_rec_fe_b_l = add_row("   в т.ч. свободный член 'b' регрессии Fe (ссылка)", "доля", 0, f"={cell_map['REC_FE_B']}", f"={cell_map['REC_FE_B']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_max_fe_l = add_row("   в т.ч. потолок извлечения Fe на ОФ (ссылка)", "%", 0, f"={cell_map['MAX_REC_FE']}", f"={cell_map['MAX_REC_FE']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_rec_fe_calc}"] = f"=MIN(E{r_max_fe_l}, E{r_fe_feed}*E{r_rec_fe_a_l}+E{r_rec_fe_b_l})"
    ws_calc[f"F{r_rec_fe_calc}"] = f"=MIN(F{r_max_fe_l}, F{r_fe_feed}*F{r_rec_fe_a_l}+F{r_rec_fe_b_l})"

    r_prod_fe = add_row("Выпуск ЖРК (железорудного концентрата)", "т", 0, 0, 0, "Масса_ОФ * Fe% * Извлечение_ОФ / Качество_ЖРК", fmt_dec2)
    r_g_fe_l = add_row("   в т.ч. содержание Fe в товарном ЖРК (ссылка)", "%", 0, f"={cell_map['GRADE_FE_CONC']}", f"={cell_map['GRADE_FE_CONC']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_prod_fe}"] = f"=(E{r_of_mass}*E{r_fe_feed}*E{r_rec_fe_calc})/E{r_g_fe_l}"
    ws_calc[f"F{r_prod_fe}"] = f"=(F{r_of_mass}*F{r_fe_feed}*F{r_rec_fe_calc})/F{r_g_fe_l}"

    r_rec_p_calc = add_row("Расчетное сквозное извлечение P2O5 на ОФ", "%", 0, 0, 0, "Формула технологов с ограничением потолка", fmt_pct)
    r_rec_p_a_l = add_row("   в т.ч. коэффициент 'a' регрессии P2O5 (ссылка)", "коэф.", 0, f"={cell_map['REC_P_A']}", f"={cell_map['REC_P_A']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_rec_p_b_l = add_row("   в т.ч. свободный член 'b' регрессии P2O5 (ссылка)", "доля", 0, f"={cell_map['REC_P_B']}", f"={cell_map['REC_P_B']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_max_p_l = add_row("   в т.ч. потолок извлечения P2O5 на ОФ (ссылка)", "%", 0, f"={cell_map['MAX_REC_P']}", f"={cell_map['MAX_REC_P']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_rec_p_calc}"] = f"=MIN(E{r_max_p_l}, E{r_p_feed}*E{r_rec_p_a_l}+E{r_rec_p_b_l})"
    ws_calc[f"F{r_rec_p_calc}"] = f"=MIN(F{r_max_p_l}, F{r_p_feed}*F{r_rec_p_a_l}+F{r_rec_p_b_l})"

    r_prod_p = add_row("Выпуск АПК (апатитового концентрата)", "т", 0, 0, 0, "Масса_ОФ * P2O5% * Извлечение_ОФ / Качество_АПК", fmt_dec2)
    r_g_p_l = add_row("   в т.ч. содержание P2O5 в товарном АПК (ссылка)", "%", 0, f"={cell_map['GRADE_P_CONC']}", f"={cell_map['GRADE_P_CONC']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_prod_p}"] = f"=(E{r_of_mass}*E{r_p_feed}*E{r_rec_p_calc})/E{r_g_p_l}"
    ws_calc[f"F{r_prod_p}"] = f"=(F{r_of_mass}*F{r_p_feed}*F{r_rec_p_calc})/F{r_g_p_l}"
    r_prod_zr = add_row("Выпуск БК (бадделеитового концентрата)", "т", 0, 0, 0, "Масса_ОФ * ZrO2% * Извлечение_БК / Качество_БК", fmt_dec2)
    r_rec_zr_l = add_row("   в т.ч. сквозное извлечение ZrO2 в БК (ссылка)", "%", 0, f"={cell_map['REC_ZR']}", f"={cell_map['REC_ZR']}", "Импорт из Inputs", fmt_pct, "SUB")
    r_g_zr_l = add_row("   в т.ч. содержание ZrO2 в товарном БК (ссылка)", "%", 0, f"={cell_map['GRADE_ZR_CONC']}", f"={cell_map['GRADE_ZR_CONC']}", "Импорт из Inputs", fmt_pct, "SUB")
    ws_calc[f"E{r_prod_zr}"] = f"=(E{r_of_mass}*E{r_zr_feed}*E{r_rec_zr_l})/E{r_g_zr_l}"
    ws_calc[f"F{r_prod_zr}"] = f"=(F{r_of_mass}*F{r_zr_feed}*F{r_rec_zr_l})/F{r_g_zr_l}"

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

    r_rev_gravel = add_row("Экономия на покупке дорожного щебня (хвосты СМС)", "руб.", 0, 0, 0, "Активно при выборе опции 2 в Inputs", fmt_rub)
    r_opt_l = add_row("   в т.ч. опция использования хвостов СМС (ссылка)", "код", 0, 0, f"={cell_map['SMS_TAIL_OPT']}", "Импорт из Inputs (1 = отвал, 2 = щебень)", fmt_int, "SUB")
    r_gval_l = add_row("   в т.ч. ценность/экономия на щебне (ссылка)", "руб./т", 0, 0, f"={cell_map['GRAVEL_VAL']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"F{r_rev_gravel}"] = f"=IF(F{r_opt_l}=2, (F{r_tail_mass}/(1-F{r_moist_l}))*F{r_gval_l}, 0)"

    r_gross_rev = add_row("ИТОГО ДОХОД / ЦЕННОСТЬ БЛОКА (GROSS VALUE)", "руб.", f"=D{r_rev_fe}+D{r_rev_p}+D{r_rev_zr}+D{r_rev_gravel}", f"=E{r_rev_fe}+E{r_rev_p}+E{r_rev_zr}+E{r_rev_gravel}", f"=F{r_rev_fe}+F{r_rev_p}+F{r_rev_zr}+F{r_rev_gravel}", "Сумма выручки от концентратов и экономии", fmt_rub, "TOTAL")

    # --- СЕКЦИЯ 4: ПЕРЕМЕННЫЕ ОПЕРАЦИОННЫЕ ЗАТРАТЫ (OPEX) ---
    add_section("4. ПЕРЕМЕННЫЕ ОПЕРАЦИОННЫЕ ЗАТРАТЫ (OPEX)")
    r_opex_tr = add_row("Транспортировка горной массы и продуктов", "руб.", 0, 0, 0, "Расчет по слинкованным плечам и УДР ниже", fmt_rub)
    r_d_main_l = add_row("   в т.ч. плечо 1-го этапа (Отвал / ОФ / ДСК-СМС) (ссылка)", "км", f"={cell_map['DIST_WASTE']}", f"={cell_map['DIST_OF']}", f"={cell_map['DIST_TO_SMS']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_d_sms_of_l = add_row("   в т.ч. плечо 3б: ДСК-СМС -> ОФ (предконцентрат) (ссылка)", "км", 0, 0, f"={cell_map['DIST_SMS_OF']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_d_sms_tw_l = add_row("   в т.ч. плечо 3в: ДСК-СМС -> Отвал хвостов (ссылка)", "км", 0, 0, f"={cell_map['DIST_SMS_WASTE']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_d_sms_tg_l = add_row("   в т.ч. плечо 3г: ДСК-СМС -> Штабель щебня (ссылка)", "км", 0, 0, f"={cell_map['DIST_SMS_GRAVEL']}", "Импорт из Inputs", fmt_dec2, "SUB")
    r_ctr_l = add_row("   в т.ч. УДР транспортировки горной массы (ссылка)", "руб./т*км", f"={cell_map['OPEX_TR']}", f"={cell_map['OPEX_TR']}", f"={cell_map['OPEX_TR']}", "Импорт из Inputs", fmt_rub, "SUB")

    ws_calc[f"D{r_opex_tr}"] = f"=D{r_wet}*D{r_d_main_l}*D{r_ctr_l}"
    ws_calc[f"E{r_opex_tr}"] = f"=E{r_wet}*E{r_d_main_l}*E{r_ctr_l}"
    ws_calc[f"F{r_opex_tr}"] = f"=F{r_wet}*F{r_d_main_l}*F{r_ctr_l} + (F{r_of_mass}/(1-F{r_moist_l}))*F{r_d_sms_of_l}*F{r_ctr_l} + (F{r_tail_mass}/(1-F{r_moist_l}))*IF(F{r_opt_l}=1,F{r_d_sms_tw_l},F{r_d_sms_tg_l})*F{r_ctr_l}"

    r_opex_crush = add_row("Дробление и сухая магнитная сепарация (ДСК-СМС)", "руб.", 0, 0, 0, "Расчет по слинкованным УДР передела ниже", fmt_rub)
    r_ccr_l = add_row("   в т.ч. УДР дробления / ДСК-СМС на 1 т сухой руды (ссылка)", "руб./т", 0, f"={cell_map['OPEX_CRUSH_OF']}", f"={cell_map['OPEX_SMS']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"E{r_opex_crush}"] = f"=E{r_dry}*E{r_ccr_l}"
    ws_calc[f"F{r_opex_crush}"] = f"=F{r_dry}*F{r_ccr_l}"

    r_opex_of = add_row("Обогащение на ОФ (измельчение, флотация, доводка)", "руб.", 0, 0, 0, "Начисляется на массу, поступившую на ОФ", fmt_rub)
    r_cof_fe_l = add_row("   в т.ч. УДР переработки ЖРК на ОФ (ссылка)", "руб./т", 0, f"={cell_map['OPEX_OF_FE']}", f"={cell_map['OPEX_OF_FE']}", "Импорт из Inputs", fmt_rub, "SUB")
    r_cof_p_l = add_row("   в т.ч. УДР переработки АПК на ОФ (ссылка)", "руб./т", 0, f"={cell_map['OPEX_OF_P']}", f"={cell_map['OPEX_OF_P']}", "Импорт из Inputs", fmt_rub, "SUB")
    r_cof_zr_l = add_row("   в т.ч. УДР переработки БК на ОФ (ссылка)", "руб./т", 0, f"={cell_map['OPEX_OF_ZR']}", f"={cell_map['OPEX_OF_ZR']}", "Импорт из Inputs", fmt_rub, "SUB")
    ws_calc[f"E{r_opex_of}"] = f"=E{r_of_mass}*(E{r_cof_fe_l}+E{r_cof_p_l}+E{r_cof_zr_l})"
    ws_calc[f"F{r_opex_of}"] = f"=F{r_of_mass}*(F{r_cof_fe_l}+F{r_cof_p_l}+F{r_cof_zr_l})"

    r_tot_opex = add_row("ИТОГО ПЕРЕМЕННЫЕ ЗАТРАТЫ (OPEX)", "руб.", f"=D{r_opex_tr}+D{r_opex_crush}+D{r_opex_of}", f"=E{r_opex_tr}+E{r_opex_crush}+E{r_opex_of}", f"=F{r_opex_tr}+F{r_opex_crush}+F{r_opex_of}", "Суммарные релевантные затраты на 1 т", fmt_rub, "TOTAL")

    # --- СЕКЦИЯ 5: ЭКОНОМИЧЕСКИЙ ЭФФЕКТ И ВЫБОР АЛГОРИТМА ---
    add_section("5. ЭКОНОМИЧЕСКИЙ ЭФФЕКТ И ВЫБОР АЛГОРИТМА")
    r_net_profit = add_row("Маржинальная прибыль (Net Value = Доход - OPEX)", "руб./т", f"=D{r_gross_rev}-D{r_tot_opex}", f"=E{r_gross_rev}-E{r_tot_opex}", f"=F{r_gross_rev}-F{r_tot_opex}", "Критерий максимизации стоимости", fmt_rub, "TOTAL")
    r_diff_waste = add_row("Прирост прибыли относительно Отвала (Сц.1)", "руб./т", 0, f"=E{r_net_profit}-$D${r_net_profit}", f"=F{r_net_profit}-$D${r_net_profit}", "Если > 0, выгоднее перерабатывать, чем везти в отвал", fmt_rub)
    r_diff_sms = add_row("Эффект от рудоразборки на ДСК-СМС против прямой ОФ", "руб./т", 0, 0, f"=F{r_net_profit}-E{r_net_profit}", "Выигрыш Сц.3 по сравнению с прямой подачей на ОФ (Сц.2)", fmt_rub)
    r_decision = add_row("ОПТИМАЛЬНОЕ НАПРАВЛЕНИЕ ДЛЯ БАЗОВОГО БЛОКА", "решение", f'=IF(MAX(D{r_net_profit}:F{r_net_profit})=D{r_net_profit},"1. В ОТВАЛ",IF(MAX(D{r_net_profit}:F{r_net_profit})=E{r_net_profit},"2. НАПРЯМУЮ НА ОФ","3. ЧЕРЕЗ ДСК-СМС"))', "", "", "Автоматический выбор максимального Net Value", "@", "TOTAL")

    ws_calc.merge_cells(start_row=r_decision, start_column=4, end_row=r_decision, end_column=6)
    ws_calc.cell(row=r_decision, column=4).fill = fill_sms

    ws_calc.column_dimensions["A"].width = 3
    ws_calc.column_dimensions["B"].width = 56
    ws_calc.column_dimensions["C"].width = 10
    ws_calc.column_dimensions["D"].width = 18
    ws_calc.column_dimensions["E"].width = 20
    ws_calc.column_dimensions["F"].width = 22
    ws_calc.column_dimensions["G"].width = 48

    # =========================================================================
    # ЛИСТ 3: CUTOFF_MATRICES_VISUAL (6 МАТРИЦ В 2 КОЛОНКИ, ГРАФИК, МЕДИАНА)
    # =========================================================================
    from openpyxl.worksheet.formula import ArrayFormula

    for sv in ws_mat.views.sheetView:
        sv.showGridLines = True

    ws_mat["B2"] = "МАТРИЧНЫЙ АНАЛИЗ CUT-OFF GRADE: ЧУВСТВИТЕЛЬНОСТЬ К СОДЕРЖАНИЯМ Fe И P2O5"
    ws_mat["B2"].font = f_title

    ws_mat["B3"] = "Слева: сводные карты решений и эффектов (Матрицы 1-3). Справа: индивидуальная рентабельность сценариев (Матрицы 4_1 - 4_3). Красная рамка — базовый блок из Inputs."
    ws_mat["B3"].font = f_note

    fe_steps = (0.03, 0.05, 0.07, 0.09, 0.10, 0.11, 0.12, 0.13, 0.14, 0.15, 0.16, 0.17, 0.19, 0.21, 0.23, 0.25)
    p_steps = (0.010, 0.015, 0.020, 0.025, 0.030, 0.035, 0.040, 0.045, 0.050, 0.055, 0.060, 0.068, 0.076, 0.084, 0.092)

    v_moist = cell_map["MOISTURE"]
    v_fe_conv = cell_map["FEO_CONV"]
    v_dil = cell_map["DILUTION"]
    rec_fe_a = cell_map["REC_FE_A"]
    rec_fe_b = cell_map["REC_FE_B"]
    max_rec_fe = cell_map["MAX_REC_FE"]
    g_fe_c = cell_map["GRADE_FE_CONC"]
    p_fe = cell_map["PRICE_FE"]
    rec_p_a = cell_map["REC_P_A"]
    rec_p_b = cell_map["REC_P_B"]
    max_rec_p = cell_map["MAX_REC_P"]
    g_p_c = cell_map["GRADE_P_CONC"]
    p_p = cell_map["PRICE_P"]
    fx = cell_map["FX_RATE"]
    sms_y = cell_map["SMS_YIELD"]
    sms_r_fe = cell_map["SMS_REC_FE"]
    sms_r_p = cell_map["SMS_REC_P"]
    fe_base_ref = cell_map["FE_BASE"]
    p_base_ref = cell_map["P2O5_BASE"]

    def build_matrix(start_row, start_col, title, mode):
        # start_col — индекс столбца с заголовками Fe (например, 2 = 'B', 19 = 'S')
        first_data_col = start_col + 1
        last_data_col = start_col + len(p_steps)
        fe_col_let = get_column_letter(start_col)
        first_data_let = get_column_letter(first_data_col)
        last_data_let = get_column_letter(last_data_col)

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

                g_fe_sc2 = f"({fe_ref}*{v_fe_conv}*(1-{v_dil}))"
                g_p_sc2 = f"({p_ref}*(1-{v_dil}))"

                g_fe_sc3 = f"({g_fe_sc2}*{sms_r_fe}/{sms_y})"
                g_p_sc3 = f"({g_p_sc2}*{sms_r_p}/{sms_y})"

                rev_fe_sc2 = f"((1-{v_moist})*{g_fe_sc2}*MIN({max_rec_fe},{g_fe_sc2}*{rec_fe_a}+{rec_fe_b})/{g_fe_c}*{p_fe}*{fx})"
                rev_p_sc2 = f"((1-{v_moist})*{g_p_sc2}*MIN({max_rec_p},{g_p_sc2}*{rec_p_a}+{rec_p_b})/{g_p_c}*{p_p}*{fx})"
                profit_sc2 = f"({rev_fe_sc2} + {rev_p_sc2} + Unit_Economics_Calc!$E${r_rev_zr} - Unit_Economics_Calc!$E${r_tot_opex})"

                rev_fe_sc3 = f"((1-{v_moist})*{sms_y}*{g_fe_sc3}*MIN({max_rec_fe},{g_fe_sc3}*{rec_fe_a}+{rec_fe_b})/{g_fe_c}*{p_fe}*{fx})"
                rev_p_sc3 = f"((1-{v_moist})*{sms_y}*{g_p_sc3}*MIN({max_rec_p},{g_p_sc3}*{rec_p_a}+{rec_p_b})/{g_p_c}*{p_p}*{fx})"
                profit_sc3 = f"({rev_fe_sc3} + {rev_p_sc3} + Unit_Economics_Calc!$F${r_rev_zr} + Unit_Economics_Calc!$F${r_rev_gravel} - Unit_Economics_Calc!$F${r_tot_opex})"

                profit_sc1 = f"Unit_Economics_Calc!$D${r_net_profit}"

                if mode == "DECISION":
                    formula = f'=IF(MAX({profit_sc1},{profit_sc2},{profit_sc3})={profit_sc1},"ОТВАЛ",IF(MAX({profit_sc1},{profit_sc2},{profit_sc3})={profit_sc2},"ОФ","СМС"))'
                    num_f = "@"
                elif mode == "NET_MAX":
                    formula = f"=MAX({profit_sc1},{profit_sc2},{profit_sc3})"
                    num_f = fmt_rub
                elif mode == "SMS_VS_OF":
                    formula = f"={profit_sc3}-{profit_sc2}"
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
        m_range = f"{first_data_let}{r_data_start}:{last_data_let}{r_data_end}"

        # Динамическая подсветка базовой ячейки (красная жирная рамка)
        hl_formula = f"AND(${fe_col_let}{r_data_start}={fe_base_ref},{first_data_let}${r_col_hdr}={p_base_ref})"
        ws_mat.conditional_formatting.add(
            m_range,
            FormulaRule(formula=(hl_formula,), stopIfTrue=False, border=border_highlight, font=Font(bold=True, color="9C0006"))
        )

        return r_data_start, r_data_end, m_range
    # Координаты колонок для левого блока (Матрицы 1, 2, 3) и правого блока (Матрицы 4_1, 4_2, 4_3)
    left_col = 2                               # Колонка B
    gap_col = left_col + len(p_steps) + 1      # Пустой разделительный столбец R
    right_col = gap_col + 1                    # Колонка S

    # --- ЛЕВАЯ КОЛОНКА: Матрицы 1, 2 и 3 ---
    r1_start, r1_end, range_m1 = build_matrix(5, left_col, "МАТРИЦА 1: КАРТА ПРИНЯТИЯ РЕШЕНИЙ (CUT-OFF GRADE: ОТВАЛ / ОФ / ДСК-СМС)", "DECISION")
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=('"ОТВАЛ"',), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=('"ОФ"',), fill=fill_of, font=Font(bold=True, color="1F4E78")))
    ws_mat.conditional_formatting.add(range_m1, CellIsRule(operator="equal", formula=('"СМС"',), fill=fill_sms, font=Font(bold=True, color="006100")))

    row_m2 = r1_end + 3
    r2_start, r2_end, range_m2 = build_matrix(row_m2, left_col, "МАТРИЦА 2: МАКСИМАЛЬНАЯ МАРЖИНАЛЬНАЯ ПРИБЫЛЬ НА 1 Т ПОРОДЫ (РУБ./Т)", "NET_MAX")
    ws_mat.conditional_formatting.add(range_m2, ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="num", mid_value=0, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"
    ))

    row_m3 = r2_end + 3
    r3_start, r3_end, range_m3 = build_matrix(row_m3, left_col, "МАТРИЦА 3: ЭФФЕКТ ОТ РУДОРАЗБОРКИ НА ДСК-СМС ПРОТИВ ПРЯМОЙ ОФ (РУБ./Т)", "SMS_VS_OF")
    ws_mat.conditional_formatting.add(range_m3, ColorScaleRule(
        start_type="min", start_color="D9E1F2",
        mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="max", end_color="70AD47"
    ))

    # --- ПРАВАЯ КОЛОНКА (ровно напротив Матриц 1, 2, 3): Матрицы 4_1, 4_2, 4_3 ---
    _, _, range_m4_1 = build_matrix(5, right_col, "МАТРИЦА 4_1: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ — СЦЕНАРИЙ 1: В ОТВАЛ (РУБ./Т)", "SC1_ONLY")
    ws_mat.conditional_formatting.add(range_m4_1, CellIsRule(operator="lessThan", formula=("0",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m4_1, CellIsRule(operator="greaterThanOrEqual", formula=("0",), fill=fill_sms, font=Font(bold=True, color="006100")))

    _, _, range_m4_2 = build_matrix(row_m2, right_col, "МАТРИЦА 4_2: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ — СЦЕНАРИЙ 2: ПРЯМАЯ ОФ (РУБ./Т)", "SC2_ONLY")
    ws_mat.conditional_formatting.add(range_m4_2, CellIsRule(operator="lessThan", formula=("0",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m4_2, CellIsRule(operator="greaterThanOrEqual", formula=("0",), fill=fill_of, font=Font(bold=True, color="1F4E78")))

    _, _, range_m4_3 = build_matrix(row_m3, right_col, "МАТРИЦА 4_3: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ — СЦЕНАРИЙ 3: ДСК-СМС -> ОФ (РУБ./Т)", "SC3_ONLY")
    ws_mat.conditional_formatting.add(range_m4_3, CellIsRule(operator="lessThan", formula=("0",), fill=fill_waste, font=Font(bold=True, color="9C0006")))
    ws_mat.conditional_formatting.add(range_m4_3, CellIsRule(operator="greaterThanOrEqual", formula=("0",), fill=fill_sms, font=Font(bold=True, color="006100")))


    # --- Ширина столбцов для обеих матричных зон ---
    ws_mat.column_dimensions["A"].width = 3
    ws_mat.column_dimensions[get_column_letter(left_col)].width = 13
    for j in range(len(p_steps)):
        ws_mat.column_dimensions[get_column_letter(left_col + 1 + j)].width = 11.5

    ws_mat.column_dimensions[get_column_letter(gap_col)].width = 4
    ws_mat.column_dimensions[get_column_letter(right_col)].width = 13
    for j in range(len(p_steps)):
        ws_mat.column_dimensions[get_column_letter(right_col + 1 + j)].width = 11.5

    # --- График и статистика в самом правом краю (после Матриц 4_1 - 4_3) ---
    far_right_col = right_col + len(p_steps) + 2
    ws_mat.column_dimensions[get_column_letter(far_right_col - 1)].width = 4

    chart = LineChart()
    chart.title = "Зависимость максимальной удельной прибыли (руб./т) от Fe при разном P2O5"
    chart.style = 10
    chart.y_axis.title = "Прибыль, руб./т"
    chart.x_axis.title = "Содержание Fe в исходной руде, %"
    chart.width = 22
    chart.height = 12

    data_ref = Reference(ws_mat, min_col=left_col + 1, min_row=r2_start - 1, max_col=left_col + len(p_steps), max_row=r2_end)
    cats_ref = Reference(ws_mat, min_col=left_col, min_row=r2_start, max_row=r2_end)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    ws_mat.add_chart(chart, f"{get_column_letter(far_right_col)}5")

    # --- Блок статистики и медианных содержаний для зоны СМС ---
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
        ("Количество сочетаний (ячеек) для СМС", f'=COUNTIF({range_m1},"СМС")', fmt_int, False),
        ("Доля матрицы, эффективная для СМС", f'=COUNTIF({range_m1},"СМС")/COUNTA({range_m1})', fmt_pct, False),
        ("Медиана содержания Fe в руде (зона СМС)", f'=IFERROR(MEDIAN(IF({range_m1}="СМС",{fe_col_range})),0)', fmt_pct, True),
        ("Медиана содержания P2O5 в руде (зона СМС)", f'=IFERROR(MEDIAN(IF({range_m1}="СМС",{p_row_range})),0)', fmt_pct, True),
        ("Среднее содержание Fe в руде (зона СМС)", f'=IFERROR(AVERAGE(IF({range_m1}="СМС",{fe_col_range})),0)', fmt_pct, True),
        ("Среднее содержание P2O5 в руде (зона СМС)", f'=IFERROR(AVERAGE(IF({range_m1}="СМС",{p_row_range})),0)', fmt_pct, True),
    )

    for idx_s, (s_name, s_formula, s_fmt, is_arr) in enumerate(stats_rows):
        curr_s_row = stat_row + 1 + idx_s
        c_lbl = ws_mat.cell(row=curr_s_row, column=stat_col_start, value=s_name)
        c_lbl.font = f_formula
        c_lbl.border = border_data

        cell_ref = f"{get_column_letter(stat_col_val)}{curr_s_row}"
        val_to_write = ArrayFormula(cell_ref, s_formula) if is_arr else s_formula

        c_v = ws_mat.cell(row=curr_s_row, column=stat_col_val, value=val_to_write)
        c_v.font = f_total
        c_v.fill = fill_sms
        c_v.border = border_data
        c_v.number_format = s_fmt
        c_v.alignment = Alignment(horizontal="right")

    ws_mat.column_dimensions[get_column_letter(stat_col_start)].width = 38
    ws_mat.column_dimensions[get_column_letter(stat_col_val)].width = 16

    wb.save(filename)
    print(f"Файл ФЭМ успешно сгенерирован: {filename}")


if __name__ == "__main__":
    create_cut_off_fem()