from flask import Flask, request, send_file
from flask_cors import CORS
import openpyxl
import openpyxl.utils
from openpyxl.styles import numbers, Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.views import Pane
from copy import copy
import io
from datetime import datetime
import json
import traceback

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max file size
CORS(app)

# Список целевых контрагентов (используется только для фильтрации документов)
TARGET_CONTRAGENTS = ['ВАНКОРНЕФТЬ АО', 'РН-Ванкор ООО']

# Колонки (1‑индексация Excel)
COLUMNS = {
    'DOCUMENT_NAME': 1,      # A
    'DEBT_AMOUNT': 12,       # L
    'UNSIGNED_DEBT': 13,     # M
    'OVERDUE': 15,           # O
    'UNSIGNED_OVERDUE': 16,  # P
    'DAYS': 18,              # R
    'OUR_DEBT': 19,          # S - "Наш долг" (не трогаем)
    'NOT_OVERDUE': 20,       # T - "Не просрочено"
    'INTERVAL_1_15': 21,     # U - "От 1 до 15 дней"
    'INTERVAL_16_29': 22,    # V - "От 16 до 29 дней"
    'INTERVAL_30_89': 23,    # W - "От 30 до 89 дней"
    'INTERVAL_90_179': 24,   # X - "От 90 до 179 дней"
    'INTERVAL_180_PLUS': 25, # Y - "Свыше 180 дней"
}

# Все колонки с суммами (которые нужно суммировать)
SUM_COLUMNS = [
    COLUMNS['DEBT_AMOUNT'],        # L
    COLUMNS['UNSIGNED_DEBT'],      # M
    COLUMNS['OVERDUE'],            # O
    COLUMNS['UNSIGNED_OVERDUE'],   # P
    COLUMNS['NOT_OVERDUE'],        # T
    COLUMNS['INTERVAL_1_15'],      # U
    COLUMNS['INTERVAL_16_29'],     # V
    COLUMNS['INTERVAL_30_89'],     # W
    COLUMNS['INTERVAL_90_179'],    # X
    COLUMNS['INTERVAL_180_PLUS'],  # Y
]

# Все числовые колонки (включая дни)
NUMERIC_COLUMNS = SUM_COLUMNS + [COLUMNS['DAYS']]

def is_cell_merged(ws, row, col):
    """Проверяет, является ли ячейка частью объединённого диапазона"""
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
            merged_range.min_col <= col <= merged_range.max_col):
            return True
    return False

def get_cell_to_write(ws, row, col):
    """Возвращает ячейку для записи (если объединена - возвращает главную ячейку)"""
    if not is_cell_merged(ws, row, col):
        return ws.cell(row=row, column=col)

    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
            merged_range.min_col <= col <= merged_range.max_col):
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)

    return ws.cell(row=row, column=col)

def safe_set_number_format(ws, row, col, value):
    """Безопасно устанавливает значение с проверкой на объединённые ячейки и заголовки"""
    if row <= 13:  # не трогаем заголовки
        return
    cell = get_cell_to_write(ws, row, col)
    
    # Сохраняем существующие стили перед обновлением
    existing_font = copy(cell.font) if cell.has_style else None
    existing_fill = copy(cell.fill) if cell.has_style else None
    existing_border = copy(cell.border) if cell.has_style else None
    existing_protection = copy(cell.protection) if cell.has_style else None
    
    cell.value = value
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal='right')
    
    # Восстанавливаем стили (если они были)
    if existing_font:
        cell.font = existing_font
    if existing_fill:
        cell.fill = existing_fill
    if existing_border:
        cell.border = existing_border
    if existing_protection:
        cell.protection = existing_protection

def safe_set_value(ws, row, col, value):
    """Безопасно устанавливает значение (без форматирования)"""
    if row <= 13:
        return
    cell = get_cell_to_write(ws, row, col)
    
    # Сохраняем существующие стили перед обновлением
    existing_font = copy(cell.font) if cell.has_style else None
    existing_fill = copy(cell.fill) if cell.has_style else None
    existing_border = copy(cell.border) if cell.has_style else None
    existing_protection = copy(cell.protection) if cell.has_style else None
    
    cell.value = value
    # Для дней тоже применяем выравнивание вправо
    cell.alignment = Alignment(horizontal='right')
    
    # Восстанавливаем стили (если они были)
    if existing_font:
        cell.font = existing_font
    if existing_fill:
        cell.fill = existing_fill
    if existing_border:
        cell.border = existing_border
    if existing_protection:
        cell.protection = existing_protection

def align_numeric_cells(ws):
    """Выравнивает все числовые ячейки по правому краю"""
    print("Выравнивание числовых ячеек по правому краю...")

    for row in range(14, ws.max_row + 1):  # начиная с 14 строки (после заголовков)
        for col in NUMERIC_COLUMNS:
            cell = ws.cell(row=row, column=col)
            if cell.value is not None and not is_cell_merged(ws, row, col):
                # Сохраняем существующие стили
                existing_font = copy(cell.font) if cell.has_style else None
                existing_fill = copy(cell.fill) if cell.has_style else None
                existing_border = copy(cell.border) if cell.has_style else None
                existing_protection = copy(cell.protection) if cell.has_style else None
                
                cell.alignment = Alignment(horizontal='right')
                
                # Восстанавливаем стили
                if existing_font:
                    cell.font = existing_font
                if existing_fill:
                    cell.fill = existing_fill
                if existing_border:
                    cell.border = existing_border
                if existing_protection:
                    cell.protection = existing_protection

def get_cell_value(ws, row, col):
    """Безопасно получает значение ячейки с учётом объединённых"""
    if is_cell_merged(ws, row, col):
        cell = get_cell_to_write(ws, row, col)
        return cell.value
    return ws.cell(row=row, column=col).value

def safe_set_top_table_value(ws, row, col, value):
    """Безопасно устанавливает значение в верхней таблице (строки 1-8) с учётом объединённых ячеек"""
    cell = get_cell_to_write(ws, row, col)
    cell.value = value
    cell.number_format = '#,##0.00' if col == 5 else '0.00%'  # E - суммы, F - проценты
    cell.alignment = Alignment(horizontal='right')

def get_interval_col(days):
    """Возвращает колонку для указанного количества дней просрочки"""
    if days <= 0:
        return COLUMNS['NOT_OVERDUE']      # T
    elif 1 <= days <= 15:
        return COLUMNS['INTERVAL_1_15']    # U
    elif 16 <= days <= 29:
        return COLUMNS['INTERVAL_16_29']   # V
    elif 30 <= days <= 89:
        return COLUMNS['INTERVAL_30_89']   # W
    elif 90 <= days <= 179:
        return COLUMNS['INTERVAL_90_179']  # X
    else:
        return COLUMNS['INTERVAL_180_PLUS'] # Y

def clear_all_intervals(ws, row):
    """Очищает все интервальные колонки для строки"""
    if row <= 13:
        return
    interval_cols = [
        COLUMNS['NOT_OVERDUE'],      # T
        COLUMNS['INTERVAL_1_15'],    # U
        COLUMNS['INTERVAL_16_29'],   # V
        COLUMNS['INTERVAL_30_89'],   # W
        COLUMNS['INTERVAL_90_179'],  # X
        COLUMNS['INTERVAL_180_PLUS'], # Y
    ]
    for col in interval_cols:
        safe_set_number_format(ws, row, col, 0)

# Расширенный список типов документов
DOCUMENT_KEYWORDS = [
    'Акт', 'Реализация', 'Корректировка', 'Поступление',
    'Взаимозачет', 'Взаимозачёт', 'Списание', 'УПД', 'Счет-фактура',
    'Товарная накладная', 'ТОРГ-12', 'Универсальный передаточный'
]

def find_structure(ws):
    """Находит все строки филиалов, контрагентов, договоров и документов"""
    filials = []      # строки с "ДТ "
    kontragents = []  # любые строки, которые являются контрагентами
    dogovors = []     # строки с "Договор"
    documents = []    # строки с документами
    total_row = None  # строка с "Итого"

    for row in range(14, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        if not cell_value:
            continue

        str_val = str(cell_value).strip()

        # 1. Филиалы
        if str_val.startswith('ДТ '):
            filials.append(row)

        # 2. Итог
        elif 'Итого' in str_val or 'ИТОГО' in str_val:
            total_row = row

        # 3. Договоры
        elif str_val.startswith('Договор') or 'договор' in str_val.lower():
            dogovors.append(row)

        # 4. Документы - расширенная проверка
        elif any(keyword in str_val for keyword in DOCUMENT_KEYWORDS):
            documents.append(row)

        # 5. Контрагенты (все остальное, что не попало в другие категории)
        else:
            # Проверяем, что это не пустая строка и не служебная
            if len(str_val) > 2 and not str_val[0].isdigit():
                kontragents.append(row)

    return filials, kontragents, dogovors, documents, total_row

def recalc_totals(ws):
    """Пересчитывает все итоговые строки в файле с учётом иерархии
    
    Иерархия суммирования (строгая, без дублирования):
    Документы → Договоры → Контрагенты → Филиалы → Итого
    
    ВАЖНО: Если у контрагента нет договоров, НЕ суммируем документы напрямую.
    Контрагент уже содержит сумму в исходном файле.
    """
    print("\n=== ПЕРЕСЧЁТ ИТОГОВ ===")

    filials, kontragents, dogovors, documents, total_row = find_structure(ws)

    print(f"Найдено: филиалов={len(filials)}, контрагентов={len(kontragents)}, "
          f"договоров={len(dogovors)}, документов={len(documents)}")

    # Создаём множества для быстрого поиска
    kontragent_set = set(kontragents)
    filial_set = set(filials)
    dogovor_set = set(dogovors)
    document_set = set(documents)

    # Функция для суммирования значений ТОЛЬКО для указанных строк
    def sum_rows(rows, col):
        total = 0
        for r in rows:
            val = get_cell_value(ws, r, col)
            if isinstance(val, (int, float)):
                total += val
        return total

    # Функция для нахождения максимального значения дней среди указанных строк
    def max_days_in_rows(rows):
        max_val = 0
        for r in rows:
            val = get_cell_value(ws, r, COLUMNS['DAYS'])
            if isinstance(val, (int, float)) and val > max_val:
                max_val = val
        return max_val

    # 1. Пересчитываем договоры (суммируем документы под ними)
    for i, dog_row in enumerate(dogovors):
        # Находим документы, принадлежащие этому договору
        doc_rows = []
        for r in range(dog_row + 1, ws.max_row + 1):
            if r in dogovor_set or r in kontragent_set or r in filial_set or r == total_row:
                break
            if r in document_set:
                doc_rows.append(r)

        if doc_rows:
            print(f"Договор стр.{dog_row}: документы {doc_rows}")

            # Суммируем все денежные колонки по документам
            for col in SUM_COLUMNS:
                total = sum_rows(doc_rows, col)
                safe_set_number_format(ws, dog_row, col, total)

            # Для дней берём максимальное значение среди документов
            max_day = max_days_in_rows(doc_rows)
            safe_set_value(ws, dog_row, COLUMNS['DAYS'], max_day)

    # 2. Пересчитываем контрагентов (суммируем ТОЛЬКО договоры под ними)
    # ВАЖНО: Если нет договоров - НЕ трогаем контрагента!
    # Контрагент уже содержит правильную сумму в исходном файле.
    # Это предотвращает двойное суммирование.
    for i, kontr_row in enumerate(kontragents):
        # Находим договоры, принадлежащие этому контрагенту
        dog_rows = []
        for r in range(kontr_row + 1, ws.max_row + 1):
            if r in kontragent_set or r in filial_set or r == total_row:
                break
            if r in dogovor_set:
                dog_rows.append(r)

        if dog_rows:
            print(f"Контрагент стр.{kontr_row}: договоры {dog_rows}")

            # Суммируем ВСЕ денежные колонки по договорам
            for col in SUM_COLUMNS:
                total = sum_rows(dog_rows, col)
                safe_set_number_format(ws, kontr_row, col, total)

            max_day = max_days_in_rows(dog_rows)
            safe_set_value(ws, kontr_row, COLUMNS['DAYS'], max_day)
        else:
            # Нет договоров - НЕ трогаем контрагента
            # Он уже содержит правильную сумму из исходного файла
            print(f"Контрагент стр.{kontr_row}: нет договоров, пропускаем (используем значение из файла)")

    # 3. Пересчитываем филиалы (суммируем контрагентов под ними)
    for i, fil_row in enumerate(filials):
        # Находим контрагентов, принадлежащие этому филиалу
        kontr_rows = []
        for r in range(fil_row + 1, ws.max_row + 1):
            if r in filial_set or r == total_row:
                break
            if r in kontragent_set:
                kontr_rows.append(r)

        if kontr_rows:
            print(f"Филиал стр.{fil_row}: контрагенты {kontr_rows}")

            for col in SUM_COLUMNS:
                total = sum_rows(kontr_rows, col)
                safe_set_number_format(ws, fil_row, col, total)

            max_day = max_days_in_rows(kontr_rows)
            safe_set_value(ws, fil_row, COLUMNS['DAYS'], max_day)

    # 4. Пересчитываем общий итог (суммируем филиалы)
    if total_row and filials:
        print(f"Общий итог стр.{total_row}: филиалы {filials}")

        for col in SUM_COLUMNS:
            total = sum_rows(filials, col)
            safe_set_number_format(ws, total_row, col, total)

        max_day = max_days_in_rows(filials)
        safe_set_value(ws, total_row, COLUMNS['DAYS'], max_day)

    print("=== ПЕРЕСЧЁТ ИТОГОВ (НИЖНЯЯ ТАБЛИЦА) ЗАВЕРШЕН ===\n")

def update_top_table(ws, total_row):
    """Обновляет верхнюю сводную таблицу (строки 1-8) значениями из итоговой строки"""
    print("\n=== ОБНОВЛЕНИЕ ВЕРХНЕЙ ТАБЛИЦЫ ===")

    if not total_row:
        print("Не найдена итоговая строка")
        return

    # Получаем значения из итоговой строки
    t_value = get_cell_value(ws, total_row, COLUMNS['NOT_OVERDUE']) or 0        # T37
    u_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_1_15']) or 0     # U37
    v_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_16_29']) or 0    # V37
    w_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_30_89']) or 0    # W37
    x_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_90_179']) or 0   # X37
    y_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_180_PLUS']) or 0 # Y37
    l_value = get_cell_value(ws, total_row, COLUMNS['DEBT_AMOUNT']) or 0       # L37

    print(f"Значения из строки {total_row}:")
    print(f"  T (не просрочено): {t_value}")
    print(f"  U (1-15 дней): {u_value}")
    print(f"  V (16-29 дней): {v_value}")
    print(f"  W (30-89 дней): {w_value}")
    print(f"  X (90-179 дней): {x_value}")
    print(f"  Y (180+ дней): {y_value}")
    print(f"  L (итого): {l_value}")

    # Обновляем строки верхней таблицы с учётом объединённых ячеек
    # Строка 2: Не просрочено
    safe_set_top_table_value(ws, 2, 5, t_value)  # E2
    safe_set_top_table_value(ws, 2, 6, t_value / l_value if l_value else 0)  # F2

    # Строка 3: От 1 до 15 дней
    safe_set_top_table_value(ws, 3, 5, u_value)  # E3
    safe_set_top_table_value(ws, 3, 6, u_value / l_value if l_value else 0)  # F3

    # Строка 4: От 16 до 29 дней
    safe_set_top_table_value(ws, 4, 5, v_value)  # E4
    safe_set_top_table_value(ws, 4, 6, v_value / l_value if l_value else 0)  # F4

    # Строка 5: От 30 до 89 дней
    safe_set_top_table_value(ws, 5, 5, w_value)  # E5
    safe_set_top_table_value(ws, 5, 6, w_value / l_value if l_value else 0)  # F5

    # Строка 6: От 90 до 179 дней
    safe_set_top_table_value(ws, 6, 5, x_value)  # E6
    safe_set_top_table_value(ws, 6, 6, x_value / l_value if l_value else 0)  # F6

    # Строка 7: Свыше 180 дней
    safe_set_top_table_value(ws, 7, 5, y_value)  # E7
    safe_set_top_table_value(ws, 7, 6, y_value / l_value if l_value else 0)  # F7

    # Строка 8: Итого
    safe_set_top_table_value(ws, 8, 5, l_value)  # E8
    safe_set_top_table_value(ws, 8, 6, 1.0)      # F8

    print("Верхняя таблица обновлена")
    print("=== ОБНОВЛЕНИЕ ВЕРХНЕЙ ТАБЛИЦЫ ЗАВЕРШЕНО ===\n")

def copy_worksheet_full(ws, wb):
    """Полное копирование листа с сохранением ВСЕГО форматирования"""
    new_ws = wb.create_sheet(ws.title)
    
    # 1. Копируем все ячейки с полным форматированием
    print(f"Копирование ячеек листа '{ws.title}'...")
    for row in ws.iter_rows():
        for cell in row:
            new_cell = new_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.alignment = copy(cell.alignment)
                new_cell.protection = copy(cell.protection)
            # Копируем тип данных (формулы)
            if cell.data_type == 'f':
                new_cell.data_type = 'f'
    
    # 2. Копируем объединённые ячейки
    for merged_range in ws.merged_cells.ranges:
        new_ws.merge_cells(str(merged_range))
    
    # 3. Копируем размеры колонок и группировку
    for col_letter, col_dim in ws.column_dimensions.items():
        new_ws.column_dimensions[col_letter].width = col_dim.width
        if col_dim.outline_level:
            new_ws.column_dimensions[col_letter].outline_level = col_dim.outline_level
        if col_dim.hidden:
            new_ws.column_dimensions[col_letter].hidden = col_dim.hidden
    
    # 4. Копируем размеры строк, группировку и скрытие
    for row_num, row_dim in ws.row_dimensions.items():
        new_ws.row_dimensions[row_num].height = row_dim.height
        if row_dim.outline_level:
            new_ws.row_dimensions[row_num].outline_level = row_dim.outline_level
        if row_dim.hidden:
            new_ws.row_dimensions[row_num].hidden = row_dim.hidden
    
    # 5. Копируем настройки группировки (summary below/right)
    if hasattr(ws.sheet_properties, 'outlinePr') and ws.sheet_properties.outlinePr:
        new_ws.sheet_properties.outlinePr.summaryBelow = ws.sheet_properties.outlinePr.summaryBelow
        new_ws.sheet_properties.outlinePr.summaryRight = ws.sheet_properties.outlinePr.summaryRight
    
    # 6. Копируем настройки печати и страницы
    if ws.print_options:
        for attr in ['grid_lines', 'grid_lines_set', 'horizontal_centered', 'vertical_centered']:
            if hasattr(ws.print_options, attr):
                setattr(new_ws.print_options, attr, getattr(ws.print_options, attr))
    
    if ws.page_setup:
        for attr in ['orientation', 'paperSize', 'scale', 'fitToHeight', 'fitToWidth',
                     'pageOrder', 'blackAndWhite', 'draft', 'cellComments', 'errors']:
            if hasattr(ws.page_setup, attr) and getattr(ws.page_setup, attr) is not None:
                setattr(new_ws.page_setup, attr, getattr(ws.page_setup, attr))
    
    if ws.page_margins:
        for attr in ['left', 'right', 'top', 'bottom', 'header', 'footer']:
            if hasattr(ws.page_margins, attr):
                setattr(new_ws.page_margins, attr, getattr(ws.page_margins, attr))
    
    # 7. Копируем freeze panes
    if ws.sheet_view and ws.sheet_view.pane:
        pane = ws.sheet_view.pane
        new_ws.sheet_view.pane = Pane(
            active_pane=pane.activePane,
            state=pane.state,
            top_left_cell=pane.topLeftCell,
            x_split=pane.xSplit,
            y_split=pane.ySplit
        )
    
    # 8. Копируем автофильтр
    if ws.auto_filter.ref:
        new_ws.auto_filter.ref = ws.auto_filter.ref
    
    # 9. Копируем conditional formatting
    for cf_rule in ws.conditional_formatting._cf_rules:
        new_ws.conditional_formatting._cf_rules.append(cf_rule)
    
    # 10. Копируем таблицы (table)
    if hasattr(ws, 'tables') and ws.tables:
        for table in ws.tables:
            new_ws.tables.add(table)
    
    print(f"Лист '{ws.title}' скопирован с полным форматированием")
    return new_ws

@app.route('/save-excel', methods=['POST'])
def save_excel():
    try:
        file = request.files['file']
        data = json.loads(request.form['data'])

        print(f"\n=== ПОЛУЧЕН ЗАПРОС ===")
        print(f"Файл: {file.filename}")
        print(f"Документов для обновления: {len(data['updatedDocuments'])}")

        wb = openpyxl.load_workbook(io.BytesIO(file.read()))

        # ===== ШАГ 0: ПЕРЕИМЕНОВЫВАЕМ ВСЕ ЛИСТЫ СРАЗУ =====
        # Это самая первая операция, чтобы избежать любых конфликтов имён
        print("\n=== ШАГ 0: ПЕРЕИМЕНОВАНИЕ ЛИСТОВ ===")

        # Собираем все листы
        all_sheets = list(wb.sheetnames)
        print(f"  Исходные листы: {all_sheets}")

        # === Определяем листы по содержимому ===
        siuat_sheets = []     # листы, начинающиеся с "Свод ДЗ СИ УАТ" или "Свод ДЗ СИ"
        summary_sheets = []   # листы, содержащие "Сводные"
        main_sheets = []      # листы с данными ДЗ (содержат "ДТ ")

        for sheet_name in all_sheets:
            if sheet_name.startswith('Свод ДЗ СИ УАТ') or sheet_name.startswith('Свод ДЗ СИ'):
                siuat_sheets.append(sheet_name)
            elif 'Сводные' in sheet_name or 'Сводная' in sheet_name:
                summary_sheets.append(sheet_name)
            else:
                # Проверяем, содержит ли данные ДЗ
                test_ws = wb[sheet_name]
                is_main = False
                for r in range(1, min(30, test_ws.max_row + 1)):
                    cell_val = test_ws.cell(row=r, column=1).value
                    if cell_val and str(cell_val).strip().startswith('ДТ '):
                        is_main = True
                        break

                if is_main:
                    main_sheets.append(sheet_name)

        print(f"  main_sheets={main_sheets}, siuat_sheets={siuat_sheets}, summary_sheets={summary_sheets}")

        # === СТРАТЕГИЯ ПЕРЕИМЕНОВАНИЯ ===
        # Основной файл содержит 2 листа с данными ДТ:
        #   - Первый лист (main_sheets[0]) = главный → "Свод ДЗ"
        #   - Второй лист (main_sheets[1], если есть) = СИ УАТ → "Свод ДЗ СИ УАТ"
        #
        # Порядок переименования: сначала переименовываем ВСЁ во временные имена,
        # чтобы избежать конфликтов, затем в целевые.

        sheets_to_process = []  # (old_name, new_name)

        # 1. Сначала переименовываем все листы в уникальные временные имена
        temp_names = {}
        temp_counter = 0

        for sheet_name in all_sheets:
            if sheet_name.startswith('Свод ДЗ СИ УАТ') or sheet_name.startswith('Свод ДЗ СИ'):
                temp_counter += 1
                temp_names[sheet_name] = f'__temp_siuat_{temp_counter}__'
            elif 'Сводные' in sheet_name or 'Сводная' in sheet_name:
                temp_counter += 1
                temp_names[sheet_name] = f'__temp_summary_{temp_counter}__'

        # Для листов с данными ДЗ
        for idx, sheet_name in enumerate(main_sheets):
            if idx == 0:
                temp_names[sheet_name] = '__temp_main__'
            elif idx == 1:
                temp_names[sheet_name] = '__temp_siuat_from_main__'
            else:
                temp_names[sheet_name] = f'__temp_extra_{idx}__'

        # Переименовываем во временные имена
        for old_name, temp_name in temp_names.items():
            if old_name != temp_name and old_name in wb.sheetnames:
                wb[old_name].title = temp_name
                print(f"  Шаг 1: '{old_name}' → '{temp_name}'")

        # 2. Теперь переименовываем в целевые имена
        # Находим листы по временным именам
        for sn in list(wb.sheetnames):
            if sn == '__temp_main__':
                wb[sn].title = 'Свод ДЗ'
                print(f"  Шаг 2: '__temp_main__' → 'Свод ДЗ'")
            elif sn == '__temp_siuat_from_main__':
                wb[sn].title = 'Свод ДЗ СИ УАТ'
                print(f"  Шаг 2: '__temp_siuat_from_main__' → 'Свод ДЗ СИ УАТ'")
            elif sn.startswith('__temp_siuat_'):
                wb[sn].title = 'Свод ДЗ СИ УАТ'
                print(f"  Шаг 2: '{sn}' → 'Свод ДЗ СИ УАТ'")
            elif sn.startswith('__temp_summary_'):
                wb[sn].title = 'Сводные таблицы'
                print(f"  Шаг 2: '{sn}' → 'Сводные таблицы'")
            elif sn.startswith('__temp_extra_'):
                del wb[sn]
                print(f"  Удалён лишний лист: {sn}")

        print(f"  Листы после переименования: {wb.sheetnames}")
        print("=== ПЕРЕИМЕНОВАНИЕ ЗАВЕРШЕНО ===\n")

        # Теперь ws должен быть переименован в "Свод ДЗ"
        # Получаем правильный ws
        ws = wb['Свод ДЗ']
        wb._active_sheet_index = wb.sheetnames.index('Свод ДЗ')

        today = datetime.now().date()
        print(f"Текущая дата: {today}")

        # Находим итоговую строку
        _, _, _, _, total_row = find_structure(ws)

        # Обновляем документы
        updated_rows = set()
        for idx, item in enumerate(data['updatedDocuments']):
            row_number = item['rowNumber']
            debt_amount = float(item['amount'])
            expected_date_str = item['date']

            expected_date = None
            if expected_date_str and expected_date_str != 'null' and expected_date_str != 'None' and expected_date_str != '':
                try:
                    clean_str = expected_date_str.strip().strip('"').strip("'")
                    if clean_str:
                        expected_date = datetime.fromisoformat(clean_str.replace('Z', '+00:00')).date()
                except Exception as e:
                    print(f"  Ошибка парсинга даты: {e}")
                    expected_date = None

            clear_all_intervals(ws, row_number)

            if expected_date is None or expected_date >= today:
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], 0)
                safe_set_value(ws, row_number, COLUMNS['DAYS'], 0)
                safe_set_number_format(ws, row_number, COLUMNS['NOT_OVERDUE'], debt_amount)
                updated_rows.add(row_number)
            elif expected_date < today:
                days_overdue = (today - expected_date).days
                interval_col = get_interval_col(days_overdue)
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], debt_amount)
                safe_set_value(ws, row_number, COLUMNS['DAYS'], days_overdue)
                safe_set_number_format(ws, row_number, interval_col, debt_amount)
                updated_rows.add(row_number)

        if updated_rows:
            print(f"\nОбновлено строк: {len(updated_rows)}")
            recalc_totals(ws)
            if total_row:
                update_top_table(ws, total_row)
        else:
            print("\nНет обновлённых строк")

        align_numeric_cells(ws)

        # ===== ИЗВЛЕКАЕМ ДАННЫЕ ИЗ ГЛАВНОГО ЛИСТА (до переименования) =====
        current_day_data_from_file = extract_filial_overdue(ws)
        data['currentDayData'] = current_day_data_from_file
        total_debt_data = extract_total_row_debt(ws, total_row)
        print(f"Из итоговой строки ДТ: общая ДЗ={total_debt_data['totalDebt']}, ПДЗ={total_debt_data['totalOverdue']}")

        # ===== 3. ДОБАВЛЯЕМ ЛИСТ «Свод ДЗ СИ УАТ» =====
        siuat_file = request.files.get('siUatFile')
        siuat_total_debt = 0
        siuat_total_overdue = 0
        siuat_sheet_created = False

        print(f"\n=== ОТЛАДКА СИ УАТ ===")
        print(f"  siuat_file: {siuat_file}")
        print(f"  siuat_file.filename: {siuat_file.filename if siuat_file else 'None'}")
        print(f"  request.files keys: {list(request.files.keys())}")

        # Проверяем, переданы ли значения с фронтенда
        summary_siuat = data.get('summarySIUAT', {})
        print(f"  summarySIUAT с фронтенда: {summary_siuat}")
        if summary_siuat:
            siuat_total_debt = summary_siuat.get('totalDebt', 0) or 0
            siuat_total_overdue = summary_siuat.get('totalOverdue', 0) or 0
            print(f"  Получены данные СИ УАТ с фронтенда: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")

        if siuat_file and siuat_file.filename:
            print(f"\n=== ДОБАВЛЯЕМ ЛИСТ 'Свод ДЗ СИ УАТ' из файла {siuat_file.filename} ===")
            try:
                # Удаляем все существующие листы «Свод ДЗ СИ УАТ*»
                sheets_to_delete = [sn for sn in wb.sheetnames if sn.startswith('Свод ДЗ СИ УАТ')]
                for sheet_name in sheets_to_delete:
                    del wb[sheet_name]
                    print(f"  Удалён лист: {sheet_name}")

                # Читаем файл СИ УАТ
                siuat_file_content = siuat_file.read()
                print(f"  Размер файла СИ УАТ: {len(siuat_file_content)} байт")

                siuat_wb = openpyxl.load_workbook(io.BytesIO(siuat_file_content))
                print(f"  Листы в файле СИ УАТ: {siuat_wb.sheetnames}")

                # Берём первый лист как источник
                siuat_ws = siuat_wb.worksheets[0]
                print(f"  Используем лист: '{siuat_ws.title}'")
                print(f"  Размер листа: {siuat_ws.max_row} строк, {siuat_ws.max_column} колонок")

                # Копируем лист
                new_siuat_ws = copy_worksheet_full(siuat_ws, wb)
                new_siuat_ws.title = 'Свод ДЗ СИ УАТ'
                siuat_sheet_created = True
                print(f"  Лист скопирован и переименован в 'Свод ДЗ СИ УАТ'")
                print(f"  Все листы в wb: {wb.sheetnames}")

                # Отладка: проверяем данные в столбцах 12 и 15
                print(f"  ОТЛАДКА: проверяем столбцы 12 и 15 в новом листе...")
                for r in range(1, min(10, new_siuat_ws.max_row + 1)):
                    v12 = new_siuat_ws.cell(row=r, column=12).value
                    v15 = new_siuat_ws.cell(row=r, column=15).value
                    name = new_siuat_ws.cell(row=r, column=1).value
                    print(f"    Строка {r} ({name}): col12={v12}, col15={v15}")

                # Извлекаем суммы из листа СИ УАТ динамически (максимум в столбцах 12 и 15)
                # Используем только если с фронтенда не пришли данные (равны 0)
                if siuat_total_debt == 0:
                    print(f"  Вызываем extract_siuat_totals_by_max...")
                    file_debt, file_overdue = extract_siuat_totals_by_max(new_siuat_ws)
                    if file_debt > 0:
                        siuat_total_debt = file_debt
                        siuat_total_overdue = file_overdue
                        print(f"  Используем данные из файла: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")
                    else:
                        print(f"  ВНИМАНИЕ: extract_siuat_totals_by_max вернул 0!")
                else:
                    print(f"  Используем данные с фронтенда, файл не сканируем")
            except Exception as e:
                print(f"!!! Ошибка при добавлении листа СИ УАТ: {e}")
                traceback.print_exc()
        else:
            print("  Файл СИ УАТ не загружен (siuat_file=None или filename пустой)")
            print(f"  siuat_total_debt = {siuat_total_debt}, siuat_total_overdue = {siuat_total_overdue}")

        # ===== 4. СОЗДАЁМ ЛИСТ «Сводные таблицы» =====
        print("\n=== СОЗДАЁМ ЛИСТ 'Сводные таблицы' ===")
        try:
            # Удаляем существующий лист, если он есть
            if 'Сводные таблицы' in wb.sheetnames:
                del wb['Сводные таблицы']
                print("  Удалён существующий лист 'Сводные таблицы'")

            summary_ws = wb.create_sheet('Сводные таблицы')
            create_summary_sheet(
                summary_ws, data,
                total_debt=total_debt_data['totalDebt'],
                total_overdue=total_debt_data['totalOverdue'],
                siuat_total_debt=siuat_total_debt,
                siuat_total_overdue=siuat_total_overdue,
            )
            print("Лист 'Сводные таблицы' создан")
        except Exception as e:
            print(f"!!! Ошибка: {e}")
            traceback.print_exc()

        # Сохраняем результат
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        print("\n=== ФАЙЛ УСПЕШНО ОБРАБОТАН, ОТПРАВЛЯЕМ ===\n")

        return send_file(
            output,
            as_attachment=True,
            download_name=f'ДЗ_обновленный_{datetime.now().strftime("%Y-%m-%d")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print("\n!!! ОШИБКА !!!")
        print(str(e))
        traceback.print_exc()
        return {'error': str(e)}, 500



def create_summary_sheet(ws, data, total_debt=0, total_overdue=0, siuat_total_debt=0, siuat_total_overdue=0):
    """Создаёт лист 'Сводные таблицы' с тремя блоками."""
    print("Создание листа 'Сводные таблицы'...")

    current_date = data.get('currentDate', datetime.now().strftime('%Y-%m-%d'))
    previous_date = data.get('previousDate', '')
    current_day_data = data.get('currentDayData', {})
    previous_day_data = data.get('previousDayData', {})

    title_font = Font(bold=True, size=14)
    header_font_white = Font(bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    number_format = '#,##0.00'
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    total_font = Font(bold=True, size=11)
    total_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    # ТАБЛИЦА 1: Динамика по подразделениям
    row = 1
    ws.cell(row=row, column=1, value='Динамика по подразделениям').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 2

    headers = ['Подразделение', current_date, previous_date, 'Динамика']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    row += 1

    all_filials = sorted(set(list(current_day_data.keys()) + list(previous_day_data.keys())))
    total_current = total_previous = total_delta = 0

    for filial in all_filials:
        cv = current_day_data.get(filial, 0)
        pv = previous_day_data.get(filial, 0)
        d = cv - pv
        total_current += cv; total_previous += pv; total_delta += d
        ws.cell(row=row, column=1, value=filial).border = thin_border
        cc = ws.cell(row=row, column=2, value=cv); cc.number_format = number_format; cc.border = thin_border; cc.alignment = Alignment(horizontal='right')
        cp = ws.cell(row=row, column=3, value=pv); cp.number_format = number_format; cp.border = thin_border; cp.alignment = Alignment(horizontal='right')
        cd = ws.cell(row=row, column=4, value=d); cd.number_format = number_format; cd.border = thin_border; cd.alignment = Alignment(horizontal='right')
        if d > 0: cd.fill = red_fill
        elif d < 0: cd.fill = green_fill
        row += 1

    ws.cell(row=row, column=1, value='Общий итог').font = total_font; ws.cell(row=row, column=1).fill = total_fill; ws.cell(row=row, column=1).border = thin_border
    for col, val in enumerate([total_current, total_previous, total_delta], 2):
        c = ws.cell(row=row, column=col, value=val); c.number_format = number_format; c.font = total_font; c.fill = total_fill; c.border = thin_border; c.alignment = Alignment(horizontal='right')
    row += 3

    # ТАБЛИЦА 2: Свод задолженности ДТ
    ws.cell(row=row, column=1, value='Свод задолженности ДТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2); row += 2
    summary_dt = data.get('summaryDT', {})
    rows_dt = [
        ('общая ДЗ', total_debt), ('из них ПДЗ', total_overdue),
        ('в т.ч. Судебная', summary_dt.get('legal', 0)),
        ('не подлежащая к взысканию', summary_dt.get('notRecoverable', 0)),
        ('подлежащая к взысканию', summary_dt.get('recoverable', 0)),
    ]
    for label, value in rows_dt:
        cl = ws.cell(row=row, column=1, value=label); cl.border = thin_border
        if 'ПДЗ' in label: cl.font = Font(bold=True, color='FF0000')
        elif 'Судебная' in label: cl.font = Font(bold=True, color='0000FF')
        else: cl.font = Font(bold=True)
        cv = ws.cell(row=row, column=2, value=value); cv.number_format = number_format; cv.border = thin_border; cv.alignment = Alignment(horizontal='right')
        if 'ПДЗ' in label: cv.font = Font(bold=True, color='FF0000')
        elif 'Судебная' in label: cv.font = Font(bold=True, color='0000FF')
        else: cv.font = Font(bold=True)
        row += 1
    row += 3

    # ТАБЛИЦА 3: Свод задолженности СИ УАТ
    ws.cell(row=row, column=1, value='Свод задолженности СИ УАТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2); row += 2
    print(f"Свод СИ УАТ: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")
    for label, value in [('Общая ДЗ', siuat_total_debt), ('Из них ПДЗ', siuat_total_overdue)]:
        cl = ws.cell(row=row, column=1, value=label); cl.border = thin_border; cl.font = Font(bold=True)
        cv = ws.cell(row=row, column=2, value=value); cv.number_format = number_format; cv.border = thin_border; cv.alignment = Alignment(horizontal='right')
        if 'ПДЗ' in label: cv.font = Font(bold=True, color='FF0000')
        else: cv.font = Font(bold=True)
        row += 1
    print(f"Лист 'Сводные таблицы' создан")


def extract_filial_overdue(ws):
    """Извлекает суммы просрочки из строк филиалов (колонка O)."""
    data = {}
    for r in range(14, ws.max_row + 1):
        v = get_cell_value(ws, r, 1)
        if v and str(v).strip().startswith('ДТ '):
            ov = get_cell_value(ws, r, COLUMNS['OVERDUE'])
            data[str(v).strip()] = ov if isinstance(ov, (int, float)) else 0
    for k in data: data[k] = round(data[k], 2)
    print(f"\n=== ДАННЫЕ ФИЛИАЛОВ ===")
    for f, a in sorted(data.items()): print(f"  {f}: {a:,.2f}")
    return data


def find_siuat_columns(ws):
    """Находит колонки 'всего' и 'просроченно' в файле СИ УАТ по заголовкам."""
    total_col = None
    overdue_col = None

    # Собираем ВСЕ кандидатов с приоритетами, затем выбираем лучших
    # Приоритет 0 = самый высокий (точное/лучшее совпадение)
    total_candidates = []
    overdue_candidates = []

    # Ищем заголовки в первых 20 строках
    for r in range(1, min(21, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=r, column=c).value
            if not cell_val:
                continue
            cell_str = str(cell_val).lower().strip()
            cell_str = ' '.join(cell_str.split())

            # --- Колонка "всего" ---
            # Приоритет: "всего" (0) > "общая дз" (1) > "общая задолженность" (2) > "общая" (3)
            if 'всего' in cell_str:
                total_candidates.append((c, 0, cell_str))
            elif 'общая дз' in cell_str:
                total_candidates.append((c, 1, cell_str))
            elif 'общая задолженность' in cell_str:
                total_candidates.append((c, 2, cell_str))
            elif 'общая' in cell_str or 'total' in cell_str:
                total_candidates.append((c, 3, cell_str))

            # --- Колонка "просроченно" ---
            # Приоритет: "просроченно" (0) > "просрочка" (1) > "пдз" (2) > "просроченная дз" (3)
            if 'просроченно' in cell_str:
                overdue_candidates.append((c, 0, cell_str))
            elif 'просрочка' in cell_str:
                overdue_candidates.append((c, 1, cell_str))
            elif 'пдз' in cell_str:
                overdue_candidates.append((c, 2, cell_str))
            elif 'просроченная дз' in cell_str or 'просроченная задолженность' in cell_str:
                overdue_candidates.append((c, 3, cell_str))
            elif 'overdue' in cell_str:
                overdue_candidates.append((c, 3, cell_str))

    # Выбираем лучшего кандидата (наименьший приоритет, наименьший номер колонки)
    if total_candidates:
        total_candidates.sort(key=lambda x: (x[1], x[0]))
        best = total_candidates[0]
        total_col = best[0]
        print(f"  Найдена колонка 'всего': колонка {total_col} (приоритет {best[1]}, '{best[2]}')")
        # Отладка: показать всех кандидатов
        for c, p, s in total_candidates:
            print(f"    кандидат: колонка {c}, приоритет {p}, '{s}'")

    if overdue_candidates:
        overdue_candidates.sort(key=lambda x: (x[1], x[0]))
        best = overdue_candidates[0]
        overdue_col = best[0]
        print(f"  Найдена колонка 'просроченно': колонка {overdue_col} (приоритет {best[1]}, '{best[2]}')")
        for c, p, s in overdue_candidates:
            print(f"    кандидат: колонка {c}, приоритет {p}, '{s}'")

    # Fallback: используем стандартные колонки 12 и 15
    if total_col is None:
        total_col = 12
        print(f"  Fallback: колонка 'всего' = {total_col}")
    if overdue_col is None:
        overdue_col = 15
        print(f"  Fallback: колонка 'просроченно' = {overdue_col}")

    return total_col, overdue_col


def _parse_cell_number(value):
    """Безопасно преобразует значение ячейки в число."""
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        # Удаляем пробелы, заменяем запятую на точку
        cleaned = value.replace(' ', '').replace(',', '.')
        try:
            return float(cleaned)
        except ValueError:
            return 0
    return 0


def extract_siuat_totals_by_max(ws):
    """Извлекает totalDebt и totalOverdue из листа СИ УАТ.

    Стратегия:
    1. Ищем колонки по заголовкам с приоритетами
    2. Ищем строку заголовков (где есть "Всего" и "Просрочено") — это граница основной таблицы
    3. Ищем "ИТОГО" ПОСЛЕ строки заголовков (в основной таблице, не в сводной сверху)
    4. Если не нашли — берём максимальные значения по строкам основной таблицы
    """
    total_col, overdue_col = find_siuat_columns(ws)
    print(f"  Используем колонки: всего={total_col}, просроченно={overdue_col}")

    # Находим строку заголовков основной таблицы (где есть "Контрагент" или "Договор" в столбце 1)
    # Это граница между верхней сводной таблицей и основной таблицей
    header_row = None
    for r in range(1, min(30, ws.max_row + 1)):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val:
            str_val = str(cell_val).strip().lower()
            if 'контрагент' in str_val or 'договор' in str_val or 'филиал' in str_val:
                header_row = r
                print(f"  Найдена строка заголовков основной таблицы: строка {r} ('{cell_val}')")
                break

    # Ищем "ИТОГО" только ПОСЛЕ строки заголовков
    itogo_row = None
    start_search = header_row if header_row else 1
    for r in range(start_search, ws.max_row + 1):
        cell_val = get_cell_value(ws, r, 1)
        if cell_val:
            str_val = str(cell_val).strip().lower()
            if 'итог' in str_val:
                itogo_row = r
                print(f"  Найдена строка 'ИТОГО' на строке {r}")
                break

    # Если нашли "ИТОГО" — берём данные из неё
    if itogo_row:
        v_total = _parse_cell_number(get_cell_value(ws, itogo_row, total_col))
        v_overdue = _parse_cell_number(get_cell_value(ws, itogo_row, overdue_col))
        total_debt = round(v_total, 2)
        total_overdue = round(v_overdue, 2)
        print(f"  Данные из строки 'ИТОГО': общая ДЗ={total_debt}, ПДЗ={total_overdue}")

        if total_debt > 0 and total_overdue > 0:
            return total_debt, total_overdue

    # Если не нашли "ИТОГО" или данные = 0, ищем максимальные значения по строкам основной таблицы
    print(f"  Ищем максимальные значения по строкам основной таблицы (начиная со строки {start_search})...")
    max_debt = 0
    max_overdue = 0

    for r in range(start_search, ws.max_row + 1):
        v_total = _parse_cell_number(get_cell_value(ws, r, total_col))
        if v_total > max_debt:
            max_debt = v_total

        v_overdue = _parse_cell_number(get_cell_value(ws, r, overdue_col))
        if v_overdue > max_overdue:
            max_overdue = v_overdue

    total_debt = round(max_debt, 2)
    total_overdue = round(max_overdue, 2)

    print(f"  Результат: общая ДЗ={total_debt} (col {total_col}), ПДЗ={total_overdue} (col {overdue_col})")

    return total_debt, total_overdue


def extract_siuat_totals(ws):
    """Извлекает totalDebt и totalOverdue из листа СИ УАТ динамически."""
    total_col, overdue_col = find_siuat_columns(ws)

    total_debt = 0
    total_overdue = 0

    print(f"  Поиск строки 'ИТОГО' (колонки: всего={total_col}, просроченно={overdue_col})...")

    # Ищем строку "Итого"/"ИТОГО" (гибкий поиск)
    for r in range(1, ws.max_row + 1):
        cell_val = get_cell_value(ws, r, 1)
        if cell_val:
            str_val = str(cell_val).strip().lower()
            # Проверяем различные варианты написания
            if 'итог' in str_val or 'total' in str_val or 'всего' in str_val:
                v_total = get_cell_value(ws, r, total_col)
                v_overdue = get_cell_value(ws, r, overdue_col)
                total_debt = round(v_total, 2) if isinstance(v_total, (int, float)) else 0
                total_overdue = round(v_overdue, 2) if isinstance(v_overdue, (int, float)) else 0
                print(f"  СИ УАТ из строки '{cell_val}' (row {r}): общая ДЗ={total_debt} (col {total_col}), ПДЗ={total_overdue} (col {overdue_col})")
                break

    # Fallback: если не нашли "Итого", ищем последнюю строку с числовыми данными
    if total_debt == 0:
        print(f"  Строка 'ИТОГО' не найдена, используем fallback...")
        for r in range(ws.max_row, 0, -1):
            val = get_cell_value(ws, r, total_col)
            if isinstance(val, (int, float)) and val > 0:
                total_debt = round(val, 2)
                v_overdue = get_cell_value(ws, r, overdue_col)
                total_overdue = round(v_overdue, 2) if isinstance(v_overdue, (int, float)) else 0
                print(f"  СИ УАТ fallback строка {r}: общая ДЗ={total_debt}, ПДЗ={total_overdue}")
                break

    return total_debt, total_overdue


def extract_total_row_debt(ws, total_row):
    """Извлекает общую ДЗ и ПДЗ из итоговой строки."""
    result = {'totalDebt': 0, 'totalOverdue': 0}
    if not total_row: return result
    td = get_cell_value(ws, total_row, COLUMNS['DEBT_AMOUNT'])
    to = get_cell_value(ws, total_row, COLUMNS['OVERDUE'])
    result['totalDebt'] = round(td, 2) if isinstance(td, (int, float)) else 0
    result['totalOverdue'] = round(to, 2) if isinstance(to, (int, float)) else 0
    return result


@app.route('/save-suppliers', methods=['POST'])
def save_suppliers():
    """Обработка и сохранение сводных таблиц оплат поставщикам"""
    try:
        file = request.files['file']
        data = json.loads(request.form['data'])

        print(f"\n=== ПОЛУЧЕН ЗАПРОС НА ОБРАБОТКУ ОПЛАТ ПОСТАВЩИКАМ ===")
        print(f"Файл: {file.filename}")
        print(f"Сводных таблиц: {len(data.get('pivotTables', []))}")

        wb = openpyxl.load_workbook(io.BytesIO(file.read()))

        for pivot_table in data.get('pivotTables', []):
            sheet_name = pivot_table['sheetName']
            headers = pivot_table['headers']
            rows_data = pivot_table['data']

            print(f"\nОбработка сводной таблицы: {sheet_name}")
            print(f"  Подразделений: {len(headers)}")
            print(f"  Контрагентов: {len(rows_data)}")

            if sheet_name not in wb.sheetnames:
                print(f"  Предупреждение: лист '{sheet_name}' не найден, создаём новый")
                ws = wb.create_sheet(sheet_name)
                last_row = 0
            else:
                ws = wb[sheet_name]
                last_row = ws.max_row
                print(f"  Реестр заканчивается на строке: {last_row}")

            pivot_start_row = last_row + 4
            create_pivot_sheet_at_row(ws, headers, rows_data, 'Сводная таблица', pivot_start_row)
            print(f"  Сводная таблица добавлена со строки: {pivot_start_row}")

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        print("\n=== ФАЙЛ ОПЛАТ УСПЕШНО ОБРАБОТАН, ОТПРАВЛЯЕМ ===\n")

        return send_file(
            output,
            as_attachment=True,
            download_name=f'Оплаты_поставщикам_{datetime.now().strftime("%Y-%m-%d")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print("\n!!! ОШИБКА ПРИ ОБРАБОТКЕ ОПЛАТ !!!")
        print(str(e))
        traceback.print_exc()
        return {'error': str(e)}, 500


def create_pivot_sheet_at_row(ws, headers, rows_data, title, start_row):
    """Создаёт сводную таблицу оплат поставщикам начиная с указанной строки"""
    print(f"Создание сводной таблицы '{title}', начиная со строки {start_row}...")

    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    explanation_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    total_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    number_format = '#,##0.00'
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    row = start_row
    ws.cell(row=row, column=1, value='Сводная таблица оплат по подразделениям').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers) + 3)
    row += 2

    header_cells = ['Контрагент'] + headers + ['Итого', 'Пояснение']
    for col, header in enumerate(header_cells, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    row += 1

    total_all = 0
    for item in rows_data:
        ws.cell(row=row, column=1, value=item['contractor']).border = thin_border
        total_sum = 0
        for col_idx, h in enumerate(headers, 2):
            value = item.get(h, 0)
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.number_format = number_format
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right')
            total_sum += value
        total_all += total_sum
        cell_total = ws.cell(row=row, column=len(headers) + 2, value=total_sum)
        cell_total.number_format = number_format
        cell_total.border = thin_border
        cell_total.alignment = Alignment(horizontal='right')
        cell_explanation = ws.cell(row=row, column=len(headers) + 3, value=item.get('explanation', ''))
        cell_explanation.border = thin_border
        if item.get('explanation'):
            for col_idx in range(1, len(headers) + 4):
                ws.cell(row=row, column=col_idx).fill = explanation_fill
        row += 1

    ws.cell(row=row, column=1, value='ИТОГО').font = Font(bold=True, size=11)
    ws.cell(row=row, column=1).fill = total_fill
    ws.cell(row=row, column=1).border = thin_border

    for col_idx, h in enumerate(headers, 2):
        subtotal = sum(item.get(h, 0) for item in rows_data)
        cell = ws.cell(row=row, column=col_idx, value=subtotal)
        cell.number_format = number_format
        cell.font = Font(bold=True, size=11)
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')

    cell_grand_total = ws.cell(row=row, column=len(headers) + 2, value=total_all)
    cell_grand_total.number_format = number_format
    cell_grand_total.font = Font(bold=True, size=11)
    cell_grand_total.fill = total_fill
    cell_grand_total.border = thin_border
    cell_grand_total.alignment = Alignment(horizontal='right')

    ws.cell(row=row, column=len(headers) + 3).fill = total_fill
    ws.cell(row=row, column=len(headers) + 3).border = thin_border

    print(f"Сводная таблица '{title}' создана, строк: {row - start_row + 1}")


if __name__ == '__main__':
    print("Сервер запущен. Для остановки нажми Ctrl+C\n")
    app.run(debug=False, port=5000, host='0.0.0.0')
