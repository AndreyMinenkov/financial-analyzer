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

# Расширенный список типов документов (листья дерева)
DOCUMENT_KEYWORDS = [
    'Акт', 'Реализация', 'Корректировка', 'Поступление',
    'Взаимозачет', 'Взаимозачёт', 'Списание', 'УПД', 'Счет-фактура',
    'Товарная накладная', 'ТОРГ-12', 'Универсальный передаточный',
    'Определение'
]

# Промежуточные уровни (не листья, но и не итоги)
INTERMEDIATE_KEYWORDS = [
    'Счет на оплату', 'Счёт на оплату', 'Счет №', 'Счёт №',
    'Договор', 'договор', 'Соглашение', 'Дополнительное соглашение',
    'Передача права', 'Компенсации', 'Штрафы', 'Претензия'
]

def find_structure(ws):
    """Находит все строки филиалов, контрагентов, договоров, промежуточных и документов

    Алгоритм определения уровней:
    1. Филиал (ур. 0): начинается с "ДТ "
    2. Документ (ур. 3): содержит DOCUMENT_KEYWORDS (листья дерева)
    3. Промежуточный (ур. 2): НЕ документ, НЕ филиал, НЕ итого,
       но имеет потомков (документы или другие промежуточные)
    4. Контрагент (ур. 1): всё остальное — либо лист, либо родитель документов
    5. Итого: содержит "Итого"

    Промежуточные уровни определяются по наличию потомков:
    - Если строка имеет ниже себя документы/промежуточные до следующего филиала/контрагента — это intermediate
    - Если строка не имеет потомков — это kontragent (лист)
    """
    filials = []
    kontragent_candidates = []  # потенциальные контрагенты/промежуточные
    documents = []
    total_row = None

    # Первый проход: классифицируем по keywords
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

        # 3. Документы по keywords (листья — Акт, Реализация, УПД и т.д.)
        elif any(keyword in str_val for keyword in DOCUMENT_KEYWORDS):
            documents.append(row)

        # 4. Остальное — кандидаты в контрагенты или промежуточные
        else:
            if len(str_val) > 2 and not str_val[0].isdigit():
                kontragent_candidates.append(row)

    # Второй проход: определяем, какие kontragent_candidates имеют потомков
    # и должны быть переклассифицированы как intermediates
    kontragents = []
    intermediates = []

    # Создаём sorted список всех специальных строк для проверки потомков
    all_special_sorted = sorted(filials + kontragent_candidates + documents)
    filial_set = set(filials)
    document_set = set(documents)
    candidate_set = set(kontragent_candidates)

    # Промежуточные по keywords: строки, содержащие INTERMEDIATE_KEYWORDS
    # Эти строки — заголовки групп (Договор, Счёт, Соглашение и т.д.)
    intermediate_by_keyword = set()
    for candidate_row in kontragent_candidates:
        cell_value = ws.cell(row=candidate_row, column=1).value
        if cell_value:
            str_val = str(cell_value).strip()
            if any(keyword in str_val for keyword in INTERMEDIATE_KEYWORDS):
                intermediate_by_keyword.add(candidate_row)

    # Разделяем candidates:
    # 1. intermediate_by_keyword → сразу intermediates
    # 2. Остальные: проверяем наличие ПРЯМЫХ потомков-документов
    #    (документ между candidate и следующим intermediate/контрагентом)
    for candidate_row in kontragent_candidates:
        if candidate_row in intermediate_by_keyword:
            intermediates.append(candidate_row)
            continue

        # Проверяем, есть ли у candidate прямые потомки-документы
        # (т.е. документ до следующего intermediate/контрагента/филиала)
        candidate_idx = all_special_sorted.index(candidate_row)
        has_direct_document_child = False

        for j in range(candidate_idx + 1, len(all_special_sorted)):
            next_row = all_special_sorted[j]

            # Нашли следующий филиал или контрагент-кандидат — потомков нет
            if next_row in filial_set or next_row in candidate_set:
                break

            # Нашли intermediate_by_keyword — document не прямой потомок candidate,
            # а потомок этого intermediate. candidate = kontragent.
            if next_row in intermediate_by_keyword:
                break

            # Нашли document — прямой потомок candidate → это intermediate
            if next_row in document_set:
                has_direct_document_child = True
                break

        if has_direct_document_child:
            intermediates.append(candidate_row)
        else:
            kontragents.append(candidate_row)

    return filials, kontragents, intermediates, documents, total_row

def build_tree_hierarchy(filials, kontragents, intermediates, documents):
    """Строит дерево иерархии: филиал → контрагент → промежуточные → документы
    
    Возвращает dict: parent_row -> [child_rows]
    """
    # Объединяем все строки с сортировкой
    all_rows = sorted(filials + kontragents + intermediates + documents)
    
    # Определяем уровень каждой строки
    row_levels = {}
    for r in filials:
        row_levels[r] = 0  # Филиал - уровень 0
    for r in kontragents:
        row_levels[r] = 1  # Контрагент - уровень 1
    for r in intermediates:
        row_levels[r] = 2  # Промежуточный - уровень 2
    for r in documents:
        row_levels[r] = 3  # Документ - уровень 3 (лист)
    
    # Строим дерево: для каждой строки находим непосредственных детей
    children = {}
    
    for i, row in enumerate(all_rows):
        current_level = row_levels[row]
        children[row] = []
        
        # Ищем детей: следующие строки с уровнем current_level + 1
        for j in range(i + 1, len(all_rows)):
            next_row = all_rows[j]
            next_level = row_levels[next_row]
            
            if next_level <= current_level:
                # Нашли строку того же или высшего уровня - дети закончились
                break
            
            if next_level == current_level + 1:
                # Непосредственный ребёнок
                children[row].append(next_row)
    
    return children

def recalc_totals(ws):
    """Пересчитывает все итоговые строки в файле используя ДЕРЕВО иерархии
    
    Алгоритм:
    1. Строим дерево: филиал → контрагент → промежуточные → документы
    2. Считаем снизу вверх: каждый родитель = сумма его непосредственных детей
    3. Это исключает двойное суммирование
    """
    print("\n=== ПЕРЕСЧЁТ ИТОГОВ (ДЕРЕВО) ===")

    filials, kontragents, intermediates, documents, total_row = find_structure(ws)

    print(f"Найдено: филиалов={len(filials)}, контрагентов={len(kontragents)}, "
          f"промежуточных={len(intermediates)}, документов={len(documents)}")

    # Создаём множества для быстрого поиска
    all_special_rows = set(filials + kontragents + intermediates + documents)
    if total_row:
        all_special_rows.add(total_row)
    
    kontragent_set = set(kontragents)
    filial_set = set(filials)
    intermediate_set = set(intermediates)
    document_set = set(documents)

    # Строим дерево иерархии
    children = build_tree_hierarchy(filials, kontragents, intermediates, documents)
    
    print(f"\nДерево иерархии:")
    for parent, childs in children.items():
        if childs:
            print(f"  Строка {parent} → дети: {childs}")

    # Функция для получения значения ячейки
    def get_val(row, col):
        val = get_cell_value(ws, row, col)
        return val if isinstance(val, (int, float)) else 0

    # Считаем снизу вверх: для каждого родителя суммируем детей
    # Обрабатываем в обратном порядке (от документов к филиалам)
    all_rows_sorted = sorted(filials + kontragents + intermediates, reverse=True)
    
    calculated_rows = set()  # Отслеживаем уже пересчитанные строки

    for parent_row in all_rows_sorted:
        childs = children.get(parent_row, [])
        
        if not childs:
            # Нет детей - пропускаем (оставляем значение из файла)
            print(f"Строка {parent_row}: нет детей, пропускаем")
            continue
        
        # Проверяем, есть ли среди детей уже пересчитанные строки
        # Если да - используем их значения (они уже содержат сумму своих детей)
        # Если нет - суммируем документы напрямую
        
        has_calculated_childs = any(c in calculated_rows for c in childs)

        if has_calculated_childs:
            # Среди детей есть уже пересчитанные - суммируем их значения
            # Непересчитанные дети-документы уже учтены в значениях пересчитанных промежуточных,
            # поэтому суммировать их отдельно нельзя - это приведёт к задвоению.
            # НО: непересчитанные контрагенты-листы и документы-листы (без детей) НЕ учтены нигде,
            # их нужно добавить к сумме.
            calculated_childs = [c for c in childs if c in calculated_rows]
            # Листовые контрагенты: не пересчитаны И kontragent
            leaf_kontragents = [c for c in childs if c not in calculated_rows and c in kontragent_set]
            # Листовые документы: не пересчитаны И document (т.е. документы без потомков)
            leaf_docs = [c for c in childs if c not in calculated_rows and c in document_set]

            all_summed_childs = calculated_childs + leaf_kontragents + leaf_docs

            print(f"Строка {parent_row}: суммируем пересчитанных {calculated_childs} + листовые контрагенты {leaf_kontragents} + листовые документы {leaf_docs} (из {childs})")

            for col in SUM_COLUMNS:
                total = sum(get_val(c, col) for c in all_summed_childs)
                safe_set_number_format(ws, parent_row, col, total)

            # Для дней берём максимум
            max_day = max((get_val(c, COLUMNS['DAYS']) for c in all_summed_childs), default=0)
            safe_set_value(ws, parent_row, COLUMNS['DAYS'], max_day)
        else:
            # Нет пересчитанных детей - суммируем документы и контрагенты-листы
            doc_childs = [c for c in childs if c in document_set]
            leaf_kontragents = [c for c in childs if c in kontragent_set]

            all_summed_childs = doc_childs + leaf_kontragents

            if all_summed_childs:
                print(f"Строка {parent_row}: суммируем документы {doc_childs} + листовые контрагенты {leaf_kontragents}")

                for col in SUM_COLUMNS:
                    total = sum(get_val(c, col) for c in all_summed_childs)
                    safe_set_number_format(ws, parent_row, col, total)

                max_day = max((get_val(c, COLUMNS['DAYS']) for c in all_summed_childs), default=0)
                safe_set_value(ws, parent_row, COLUMNS['DAYS'], max_day)
            else:
                # Дети - не документы и не контрагенты (например, промежуточные без детей)
                # Оставляем как есть
                print(f"Строка {parent_row}: дети не документы и не контрагенты, пропускаем")
                continue
        
        calculated_rows.add(parent_row)

    # 4. Пересчитываем общий итог (суммируем филиалы)
    if total_row and filials:
        print(f"\nОбщий итог стр.{total_row}: филиалы {filials}")

        for col in SUM_COLUMNS:
            total = sum(get_val(f, col) for f in filials)
            safe_set_number_format(ws, total_row, col, total)

        max_day = max((get_val(f, COLUMNS['DAYS']) for f in filials), default=0)
        safe_set_value(ws, total_row, COLUMNS['DAYS'], max_day)

    print("=== ПЕРЕСЧЁТ ИТОГОВ (ДЕРЕВО) ЗАВЕРШЕН ===\n")

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
        ws = wb.active

        today = datetime.now().date()
        print(f"Текущая дата: {today}")

        # Находим итоговую строку для последующего обновления верхней таблицы
        _, _, _, _, total_row = find_structure(ws)

        # Обновляем документы
        updated_rows = set()
        for idx, item in enumerate(data['updatedDocuments']):
            row_number = item['rowNumber']
            debt_amount = float(item['amount'])
            expected_date_str = item['date']

            # Парсим дату
            expected_date = None
            if expected_date_str and expected_date_str != 'null' and expected_date_str != 'None' and expected_date_str != '':
                try:
                    clean_str = expected_date_str.strip().strip('"').strip("'")
                    if clean_str:
                        expected_date = datetime.fromisoformat(clean_str.replace('Z', '+00:00')).date()
                except Exception as e:
                    print(f"  Ошибка парсинга даты: {e}")
                    expected_date = None

            # Очищаем все интервалы перед обновлением
            clear_all_intervals(ws, row_number)

            # Если даты нет или она в будущем/сегодня - НЕ ПРОСРОЧЕНО
            if expected_date is None or expected_date >= today:
                # O (просрочено) - очищаем
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], 0)
                # R (дни) - очищаем
                safe_set_value(ws, row_number, COLUMNS['DAYS'], 0)
                # Запись в T (не просрочено)
                safe_set_number_format(ws, row_number, COLUMNS['NOT_OVERDUE'], debt_amount)
                updated_rows.add(row_number)

            elif expected_date < today:
                # ПРОСРОЧЕНО
                days_overdue = (today - expected_date).days
                interval_col = get_interval_col(days_overdue)

                # O (просрочено) - сумма долга
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], debt_amount)
                # R (дни просрочки)
                safe_set_value(ws, row_number, COLUMNS['DAYS'], days_overdue)
                # Запись в нужный интервал
                safe_set_number_format(ws, row_number, interval_col, debt_amount)
                updated_rows.add(row_number)

        if updated_rows:
            print(f"\nОбновлено строк: {len(updated_rows)}")
            # ПЕРЕСЧИТЫВАЕМ ВСЕ ИТОГИ с учётом иерархии
            recalc_totals(ws)

            # ОБНОВЛЯЕМ ВЕРХНЮЮ ТАБЛИЦУ
            if total_row:
                update_top_table(ws, total_row)
        else:
            print("\nНет обновлённых строк")

        # Выравниваем все числовые ячейки по правому краю
        align_numeric_cells(ws)

        # ===== 1. ПЕРЕИМЕНОВЫВАЕМ ГЛАВНЫЙ ЛИСТ =====
        ws.title = 'Свод ДЗ'
        print(f"\nГлавный лист переименован в 'Свод ДЗ'")

        # ===== 2. ИЗВЛЕКАЕМ ДАННЫЕ ИЗ ИТОГОВОГО ФАЙЛА =====

        # 2a. Просрочка по филиалам (для таблицы динамики)
        current_day_data_from_file = extract_filial_overdue(ws)
        data['currentDayData'] = current_day_data_from_file

        # 2b. Общая ДЗ и ПДЗ из итоговой строки (для "Свод задолженности ДТ")
        total_debt_data = extract_total_row_debt(ws, total_row)
        print(f"Из итоговой строки: общая ДЗ={total_debt_data['totalDebt']}, ПДЗ={total_debt_data['totalOverdue']}")

        # ===== 3. ДОБАВЛЯЕМ ЛИСТ «Свод ДЗ СИ УАТ» =====
        siuat_file = request.files.get('siUatFile')
        siuat_total_debt = 0
        siuat_total_overdue = 0
        if siuat_file and siuat_file.filename:
            print(f"\n=== ДОБАВЛЯЕМ ЛИСТ 'Свод ДЗ СИ УАТ' из файла {siuat_file.filename} ===")
            try:
                siuat_wb = openpyxl.load_workbook(io.BytesIO(siuat_file.read()))
                siuat_ws = siuat_wb.active

                # Полное копирование с сохранением всего форматирования
                copy_worksheet_full(siuat_ws, wb)

                # Переименовываем скопированный лист
                new_siuat_ws = wb[wb.sheetnames[-1]]
                new_siuat_ws.title = 'Свод ДЗ СИ УАТ'

                # Извлекаем общую ДЗ и ПДЗ из СИ УАТ
                siuat_totals = extract_siuat_totals(new_siuat_ws)
                siuat_total_debt = siuat_totals['totalDebt']
                siuat_total_overdue = siuat_totals['totalOverdue']
                print(f"СИ УАТ: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")

            except Exception as e:
                print(f"!!! Ошибка при добавлении листа СИ УАТ: {e}")
                traceback.print_exc()
        else:
            print("Файл СИ УАТ не загружен, пропускаем лист 'Свод ДЗ СИ УАТ'")

        # ===== 4. СОЗДАЁМ ЛИСТ «Сводные таблицы» =====
        print("\n=== СОЗДАЁМ ЛИСТ 'Сводные таблицы' ===")
        try:
            summary_ws = wb.create_sheet('Сводные таблицы')
            create_summary_sheet(
                summary_ws,
                data,
                total_debt=total_debt_data['totalDebt'],
                total_overdue=total_debt_data['totalOverdue'],
                siuat_total_debt=siuat_total_debt,
                siuat_total_overdue=siuat_total_overdue,
            )
            print("Лист 'Сводные таблицы' создан")
        except Exception as e:
            print(f"!!! Ошибка при создании листа сводных таблиц: {e}")
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
    """Создаёт лист 'Сводные таблицы' с тремя блоками:
    1. Динамика по подразделениям
    2. Свод задолженности ДТ (данные из итогового файла)
    3. Свод задолженности СИ УАТ (2 строки: Общая ДЗ, Из них ПДЗ)
    """
    print("Создание листа 'Сводные таблицы'...")

    current_date = data.get('currentDate', datetime.now().strftime('%Y-%m-%d'))
    previous_date = data.get('previousDate', '')
    current_day_data = data.get('currentDayData', {})
    previous_day_data = data.get('previousDayData', {})

    # Стили
    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font_white = Font(bold=True, size=11, color='FFFFFF')
    number_format = '#,##0.00'
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    total_font = Font(bold=True, size=11)
    total_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    # ===== ТАБЛИЦА 1: Динамика по подразделениям =====
    row = 1
    ws.cell(row=row, column=1, value='Динамика по подразделениям').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 2

    # Заголовки
    headers = ['Подразделение', current_date, previous_date, 'Динамика']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    row += 1

    # Собираем все подразделения
    all_filials = sorted(set(list(current_day_data.keys()) + list(previous_day_data.keys())))

    total_current = 0
    total_previous = 0
    total_delta = 0

    for filial in all_filials:
        current_val = current_day_data.get(filial, 0)
        previous_val = previous_day_data.get(filial, 0)
        delta = current_val - previous_val

        total_current += current_val
        total_previous += previous_val
        total_delta += delta

        ws.cell(row=row, column=1, value=filial).border = thin_border
        cell_curr = ws.cell(row=row, column=2, value=current_val)
        cell_curr.number_format = number_format
        cell_curr.border = thin_border
        cell_curr.alignment = Alignment(horizontal='right')

        cell_prev = ws.cell(row=row, column=3, value=previous_val)
        cell_prev.number_format = number_format
        cell_prev.border = thin_border
        cell_prev.alignment = Alignment(horizontal='right')

        cell_delta = ws.cell(row=row, column=4, value=delta)
        cell_delta.number_format = number_format
        cell_delta.border = thin_border
        cell_delta.alignment = Alignment(horizontal='right')

        # Цветовая индикация
        if delta > 0:
            cell_delta.fill = red_fill
        elif delta < 0:
            cell_delta.fill = green_fill

        row += 1

    # Итоговая строка
    ws.cell(row=row, column=1, value='Общий итог').font = total_font
    ws.cell(row=row, column=1).fill = total_fill
    ws.cell(row=row, column=1).border = thin_border

    cell = ws.cell(row=row, column=2, value=total_current)
    cell.number_format = number_format
    cell.font = total_font
    cell.fill = total_fill
    cell.border = thin_border
    cell.alignment = Alignment(horizontal='right')

    cell = ws.cell(row=row, column=3, value=total_previous)
    cell.number_format = number_format
    cell.font = total_font
    cell.fill = total_fill
    cell.border = thin_border
    cell.alignment = Alignment(horizontal='right')

    cell = ws.cell(row=row, column=4, value=total_delta)
    cell.number_format = number_format
    cell.font = total_font
    cell.fill = total_fill
    cell.border = thin_border
    cell.alignment = Alignment(horizontal='right')
    if total_delta > 0:
        cell.fill = PatternFill(start_color='FFB4B4', end_color='FFB4B4', fill_type='solid')
    elif total_delta < 0:
        cell.fill = PatternFill(start_color='A5D6A7', end_color='A5D6A7', fill_type='solid')

    row += 3

    # ===== ТАБЛИЦА 2: Свод задолженности ДТ =====
    ws.cell(row=row, column=1, value='Свод задолженности ДТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 2

    # Используем данные из итогового файла (не client-side расчёт)
    print(f"Свод ДТ: общая ДЗ={total_debt}, ПДЗ={total_overdue}")

    summary_dt_data = [
        ('общая ДЗ', total_debt),
        ('из них ПДЗ', total_overdue),
    ]

    for label, value in summary_dt_data:
        cell_label = ws.cell(row=row, column=1, value=label)
        cell_label.border = thin_border
        if 'ПДЗ' in label:
            cell_label.font = Font(bold=True, color='FF0000')
        else:
            cell_label.font = Font(bold=True)

        cell_value = ws.cell(row=row, column=2, value=value)
        cell_value.number_format = number_format
        cell_value.border = thin_border
        cell_value.alignment = Alignment(horizontal='right')
        if 'ПДЗ' in label:
            cell_value.font = Font(bold=True, color='FF0000')
        else:
            cell_value.font = Font(bold=True)
        row += 1

    row += 3

    # ===== ТАБЛИЦА 3: Свод задолженности СИ УАТ =====
    ws.cell(row=row, column=1, value='Свод задолженности СИ УАТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 2

    # Только 2 строки: Общая ДЗ и Из них ПДЗ (из файла СИ УАТ)
    print(f"Свод СИ УАТ: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")

    summary_siuat_data = [
        ('Общая ДЗ', siuat_total_debt),
        ('Из них ПДЗ', siuat_total_overdue),
    ]

    for label, value in summary_siuat_data:
        cell_label = ws.cell(row=row, column=1, value=label)
        cell_label.border = thin_border
        if 'ПДЗ' in label:
            cell_label.font = Font(bold=True, color='FF0000')
        else:
            cell_label.font = Font(bold=True)

        cell_value = ws.cell(row=row, column=2, value=value)
        cell_value.number_format = number_format
        cell_value.border = thin_border
        cell_value.alignment = Alignment(horizontal='right')
        if 'ПДЗ' in label:
            cell_value.font = Font(bold=True, color='FF0000')
        else:
            cell_value.font = Font(bold=True)
        row += 1

    print(f"Лист 'Сводные таблицы' создан, последняя строка: {row}")

def save_suppliers():
    """Обработка и сохранение сводных таблиц оплат поставщикам

    Структура файла на выходе:
    - Строки 1..M: Оригинальный реестр платежей
    - Строки M+3..: Сводная таблица (ниже реестра)
    """
    try:
        file = request.files['file']
        data = json.loads(request.form['data'])

        print(f"\n=== ПОЛУЧЕН ЗАПРОС НА ОБРАБОТКУ ОПЛАТ ПОСТАВЩИКАМ ===")
        print(f"Файл: {file.filename}")
        print(f"Сводных таблиц: {len(data.get('pivotTables', []))}")

        wb = openpyxl.load_workbook(io.BytesIO(file.read()))

        # Обрабатываем каждую сводную таблицу
        for pivot_table in data.get('pivotTables', []):
            sheet_name = pivot_table['sheetName']
            headers = pivot_table['headers']
            rows_data = pivot_table['data']

            print(f"\nОбработка сводной таблицы: {sheet_name}")
            print(f"  Подразделений: {len(headers)}")
            print(f"  Контрагентов: {len(rows_data)}")

            # Находим оригинальный лист
            if sheet_name not in wb.sheetnames:
                print(f"  Предупреждение: лист '{sheet_name}' не найден, создаём новый")
                ws = wb.create_sheet(sheet_name)
                last_row = 0
            else:
                ws = wb[sheet_name]
                last_row = ws.max_row
                print(f"  Реестр заканчивается на строке: {last_row}")

            # Сводные таблицы формируем НИЖЕ реестра, через 3 пустые строки
            pivot_start_row = last_row + 4

            # Создаём сводную таблицу ниже реестра
            create_pivot_sheet_at_row(ws, headers, rows_data, 'Сводная таблица', pivot_start_row)

            print(f"  Сводная таблица добавлена со строки: {pivot_start_row}")

        # Сохраняем результат
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

    # Стили
    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    explanation_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    total_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    number_format = '#,##0.00'
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    row = start_row
    # Заголовок
    ws.cell(row=row, column=1, value='Сводная таблица оплат по подразделениям').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers) + 3)
    row += 2

    # Шапка таблицы
    header_cells = ['Контрагент'] + headers + ['Итого', 'Пояснение']
    for col, header in enumerate(header_cells, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    row += 1

    # Данные
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

        # Итого
        cell_total = ws.cell(row=row, column=len(headers) + 2, value=total_sum)
        cell_total.number_format = number_format
        cell_total.border = thin_border
        cell_total.alignment = Alignment(horizontal='right')

        # Пояснение
        cell_explanation = ws.cell(row=row, column=len(headers) + 3, value=item.get('explanation', ''))
        cell_explanation.border = thin_border

        # Жёлтый фон для строк с пояснениями
        if item.get('explanation'):
            for col_idx in range(1, len(headers) + 4):
                cell = ws.cell(row=row, column=col_idx)
                cell.fill = explanation_fill

        row += 1

    # Итоговая строка
    ws.cell(row=row, column=1, value='ИТОГО').font = Font(bold=True, size=11)
    ws.cell(row=row, column=1).fill = total_fill
    ws.cell(row=row, column=1).border = thin_border

    # Суммы по подразделениям
    for col_idx, h in enumerate(headers, 2):
        subtotal = sum(item.get(h, 0) for item in rows_data)
        cell = ws.cell(row=row, column=col_idx, value=subtotal)
        cell.number_format = number_format
        cell.font = Font(bold=True, size=11)
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')

    # Общий итог
    cell_grand_total = ws.cell(row=row, column=len(headers) + 2, value=total_all)
    cell_grand_total.number_format = number_format
    cell_grand_total.font = Font(bold=True, size=11)
    cell_grand_total.fill = total_fill
    cell_grand_total.border = thin_border
    cell_grand_total.alignment = Alignment(horizontal='right')

    ws.cell(row=row, column=len(headers) + 3).fill = total_fill
    ws.cell(row=row, column=len(headers) + 3).border = thin_border

    print(f"Сводная таблица '{title}' создана, строк: {row - start_row + 1}")


def extract_filial_overdue(ws):
    """Извлекает суммы просрочки из итоговых строк филиалов (колонка O = OVERDUE)

    Проходит по всем строкам, находит строки начинающиеся с 'ДТ ',
    берёт значение из колонки COLUMNS['OVERDUE'] (просрочено).
    Используется после recalc_totals, когда итоговые строки уже пересчитаны.
    """
    filial_data = {}
    for row in range(14, ws.max_row + 1):
        # Используем get_cell_value для корректной работы с объединёнными ячейками
        cell_value = get_cell_value(ws, row, 1)
        if not cell_value:
            continue
        str_val = str(cell_value).strip()
        # Ищем строки филиалов
        if str_val.startswith('ДТ '):
            overdue = get_cell_value(ws, row, COLUMNS['OVERDUE'])
            filial_data[str_val] = overdue if isinstance(overdue, (int, float)) else 0
    # Округляем до 2 знаков
    for key in filial_data:
        filial_data[key] = round(filial_data[key], 2)
    print(f"\n=== ИЗВЛЕЧЕНЫ ДАННЫЕ ФИЛИАЛОВ (из итогового файла) ===")
    for filial, amount in sorted(filial_data.items()):
        print(f"  {filial}: {amount:,.2f}")
    print(f"  Всего филиалов: {len(filial_data)}")
    return filial_data


def extract_total_row_debt(ws, total_row):
    """Извлекает общую ДЗ (колонка L) и ПДЗ (колонка O) из итоговой строки"""
    result = {'totalDebt': 0, 'totalOverdue': 0}
    if not total_row:
        print("  Итоговая строка не найдена, используем нули")
        return result
    total_debt = get_cell_value(ws, total_row, COLUMNS['DEBT_AMOUNT'])
    total_overdue = get_cell_value(ws, total_row, COLUMNS['OVERDUE'])
    result['totalDebt'] = round(total_debt, 2) if isinstance(total_debt, (int, float)) else 0
    result['totalOverdue'] = round(total_overdue, 2) if isinstance(total_overdue, (int, float)) else 0
    return result


def extract_siuat_totals(ws):
    """Извлекает общую ДЗ (колонка L) и ПДЗ (колонка O) из итоговой строки СИ УАТ

    Ищет строку 'Итого' и берёт значения из колонок L и O.
    """
    result = {'totalDebt': 0, 'totalOverdue': 0}
    for row in range(1, ws.max_row + 1):
        cell_value = get_cell_value(ws, row, 1)
        if not cell_value:
            continue
        str_val = str(cell_value).strip()
        if 'Итого' in str_val or 'ИТОГО' in str_val:
            total_debt = get_cell_value(ws, row, COLUMNS['DEBT_AMOUNT'])
            total_overdue = get_cell_value(ws, row, COLUMNS['OVERDUE'])
            result['totalDebt'] = round(total_debt, 2) if isinstance(total_debt, (int, float)) else 0
            result['totalOverdue'] = round(total_overdue, 2) if isinstance(total_overdue, (int, float)) else 0
            print(f"  Найдена итоговая строка СИ УАТ: строка {row}")
            break
    return result


if __name__ == '__main__':
    print("Сервер запущен. Для остановки нажми Ctrl+C\n")
    app.run(debug=True, port=5000, host='0.0.0.0')
