from flask import Flask, request, send_file
from flask_cors import CORS
import openpyxl
import openpyxl.utils
from openpyxl.styles import numbers, Alignment, Font, PatternFill, Border, Side
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
    cell.value = value
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal='right')

def safe_set_value(ws, row, col, value):
    """Безопасно устанавливает значение (без форматирования)"""
    if row <= 13:
        return
    cell = get_cell_to_write(ws, row, col)
    cell.value = value
    # Для дней тоже применяем выравнивание вправо
    cell.alignment = Alignment(horizontal='right')

def align_numeric_cells(ws):
    """Выравнивает все числовые ячейки по правому краю"""
    print("Выравнивание числовых ячеек по правому краю...")

    for row in range(14, ws.max_row + 1):  # начиная с 14 строки (после заголовков)
        for col in NUMERIC_COLUMNS:
            cell = ws.cell(row=row, column=col)
            if cell.value is not None and not is_cell_merged(ws, row, col):
                cell.alignment = Alignment(horizontal='right')

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

def find_structure(ws):
    """Находит все строки филиалов, контрагентов, договоров и документов"""
    filials = []      # строки с "ДТ "
    kontragents = []  # любые строки, которые являются контрагентами
    dogovors = []     # строки с "Договор"
    documents = []    # строки с актами/реализациями
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

        # 4. Документы
        elif any(x in str_val for x in ['Акт', 'Реализация', 'Корректировка', 'Поступление']):
            documents.append(row)

        # 5. Контрагенты (все остальное, что не попало в другие категории)
        else:
            # Проверяем, что это не пустая строка и не служебная
            if len(str_val) > 2 and not str_val[0].isdigit():
                kontragents.append(row)

    return filials, kontragents, dogovors, documents, total_row

def recalc_totals(ws):
    """Пересчитывает все итоговые строки в файле с учётом иерархии"""
    print("\n=== ПЕРЕСЧЁТ ИТОГОВ ===")

    filials, kontragents, dogovors, documents, total_row = find_structure(ws)

    print(f"Найдено: филиалов={len(filials)}, контрагентов={len(kontragents)}, "
          f"договоров={len(dogovors)}, документов={len(documents)}")

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

    # 1. Пересчитываем договоры (суммируем ТОЛЬКО документы под ними)
    for i, dog_row in enumerate(dogovors):
        # Находим документы, принадлежащие этому договору
        doc_rows = []
        for r in range(dog_row + 1, ws.max_row + 1):
            if r in dogovors or r in kontragents or r in filials or r == total_row:
                break
            if r in documents:
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
    for i, kontr_row in enumerate(kontragents):
        # Находим договоры, принадлежащие этому контрагенту
        dog_rows = []
        for r in range(kontr_row + 1, ws.max_row + 1):
            if r in kontragents or r in filials or r == total_row:
                break
            if r in dogovors:
                dog_rows.append(r)

        if dog_rows:
            print(f"Контрагент стр.{kontr_row}: договоры {dog_rows}")

            # Пересчитываем ВСЕХ контрагентов (и целевых, и нецелевых)
            for col in SUM_COLUMNS:
                total = sum_rows(dog_rows, col)
                safe_set_number_format(ws, kontr_row, col, total)

            max_day = max_days_in_rows(dog_rows)
            safe_set_value(ws, kontr_row, COLUMNS['DAYS'], max_day)

    # 3. Пересчитываем филиалы (суммируем ТОЛЬКО контрагентов под ними)
    for i, fil_row in enumerate(filials):
        # Находим контрагентов, принадлежащих этому филиалу
        kontr_rows = []
        for r in range(fil_row + 1, ws.max_row + 1):
            if r in filials or r == total_row:
                break
            if r in kontragents:
                kontr_rows.append(r)

        if kontr_rows:
            print(f"Филиал стр.{fil_row}: контрагенты {kontr_rows}")

            for col in SUM_COLUMNS:
                total = sum_rows(kontr_rows, col)
                safe_set_number_format(ws, fil_row, col, total)

            max_day = max_days_in_rows(kontr_rows)
            safe_set_value(ws, fil_row, COLUMNS['DAYS'], max_day)

    # 4. Пересчитываем общий итог (суммируем ТОЛЬКО филиалы)
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

        # ===== СОЗДАЁМ ДОПОЛНИТЕЛЬНЫЕ ЛИСТЫ =====

        # data уже содержит все поля summaryData (updatedDocuments, currentDayData, summaryDT и т.д.)
        print(f"\n=== Ключи в data: {list(data.keys())}")

        # 1. Лист «Свод ДЗ СИ УАТ» — копируем из загруженного файла
        siuat_file = request.files.get('siUatFile')
        if siuat_file and siuat_file.filename:
            print(f"\n=== ДОБАВЛЯЕМ ЛИСТ 'Свод ДЗ СИ УАТ' из файла {siuat_file.filename} ===")
            try:
                siuat_wb = openpyxl.load_workbook(io.BytesIO(siuat_file.read()))
                siuat_ws = siuat_wb.active

                # Копируем ячейки, стили, объединённые диапазоны вручную
                new_ws = wb.create_sheet('Свод ДЗ СИ УАТ')

                # Копируем все ячейки с полным форматированием
                for row in siuat_ws.iter_rows(min_row=1, max_row=siuat_ws.max_row,
                                               min_col=1, max_col=siuat_ws.max_column):
                    for cell in row:
                        new_cell = new_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            new_cell.font = copy(cell.font)
                            new_cell.border = copy(cell.border)
                            new_cell.fill = copy(cell.fill)
                            new_cell.number_format = cell.number_format
                            new_cell.alignment = copy(cell.alignment)
                            new_cell.protection = copy(cell.protection)

                # Копируем объединённые ячейки
                for merged_range in siuat_ws.merged_cells.ranges:
                    new_ws.merge_cells(str(merged_range))

                # Копируем размеры колонок и группировку
                for col_letter, col_dim in siuat_ws.column_dimensions.items():
                    new_ws.column_dimensions[col_letter].width = col_dim.width
                    if col_dim.outline_level:
                        new_ws.column_dimensions[col_letter].outline_level = col_dim.outline_level
                    if col_dim.hidden:
                        new_ws.column_dimensions[col_letter].hidden = col_dim.hidden

                # Копируем размеры строк, группировку и скрытие
                for row_num, row_dim in siuat_ws.row_dimensions.items():
                    new_ws.row_dimensions[row_num].height = row_dim.height
                    if row_dim.outline_level:
                        new_ws.row_dimensions[row_num].outline_level = row_dim.outline_level
                    if row_dim.hidden:
                        new_ws.row_dimensions[row_num].hidden = row_dim.hidden

                # Копируем настройки группировки (summary below/right)
                new_ws.sheet_properties.outlinePr.summaryBelow = siuat_ws.sheet_properties.outlinePr.summaryBelow
                new_ws.sheet_properties.outlinePr.summaryRight = siuat_ws.sheet_properties.outlinePr.summaryRight

                # Копируем настройки печати и страницы
                if siuat_ws.print_options:
                    for attr in ['grid_lines', 'grid_lines_set', 'horizontal_centered', 'vertical_centered']:
                        if hasattr(siuat_ws.print_options, attr):
                            setattr(new_ws.print_options, attr, getattr(siuat_ws.print_options, attr))

                if siuat_ws.page_setup:
                    for attr in ['orientation', 'paperSize', 'scale', 'fitToHeight', 'fitToWidth',
                                 'pageOrder', 'blackAndWhite', 'draft', 'cellComments', 'errors']:
                        if hasattr(siuat_ws.page_setup, attr) and getattr(siuat_ws.page_setup, attr) is not None:
                            setattr(new_ws.page_setup, attr, getattr(siuat_ws.page_setup, attr))

                if siuat_ws.page_margins:
                    for attr in ['left', 'right', 'top', 'bottom', 'header', 'footer']:
                        if hasattr(siuat_ws.page_margins, attr):
                            setattr(new_ws.page_margins, attr, getattr(siuat_ws.page_margins, attr))

                print("Лист 'Свод ДЗ СИ УАТ' добавлен с полным форматированием")
            except Exception as e:
                print(f"!!! Ошибка при добавлении листа СИ УАТ: {e}")
                traceback.print_exc()
        else:
            print("Файл СИ УАТ не загружен, пропускаем лист 'Свод ДЗ СИ УАТ'")

        # 2. Лист «Сводные таблицы»
        print("\n=== СОЗДАЁМ ЛИСТ 'Сводные таблицы' ===")
        print(f"currentDayData: {json.dumps(data.get('currentDayData', {}), ensure_ascii=False)[:300]}")
        print(f"previousDayData: {json.dumps(data.get('previousDayData', {}), ensure_ascii=False)[:300]}")
        print(f"summaryDT: {data.get('summaryDT', {})}")
        try:
            summary_ws = wb.create_sheet('Сводные таблицы')
            create_summary_sheet(summary_ws, data)
            print("Лист 'Сводные таблицы' создан успешно")
        except Exception as e:
            print(f"!!! Ошибка при создании листа 'Сводные таблицы': {e}")
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

if __name__ == '__main__':
    print("Сервер запущен. Для остановки нажми Ctrl+C\n")
    app.run(debug=True, port=5000)


# ===== ОБРАБОТКА ОПЛАТ ПОСТАВЩИКАМ =====
@app.route('/save-suppliers', methods=['POST'])
def save_suppliers():
    try:
        file = request.files['file']
        data = json.loads(request.form['data'])

        print(f"\n=== ПОЛУЧЕН ЗАПРОС ОПЛАТ ПОСТАВЩИКАМ ===")
        print(f"Файл: {file.filename}")
        print(f"Сводных таблиц: {len(data.get('pivotTables', []))}")

        wb = openpyxl.load_workbook(io.BytesIO(file.read()))

        # Для каждой сводной таблицы добавляем на соответствующий лист
        for pivot_info in data.get('pivotTables', []):
            sheet_name = pivot_info['sheetName']
            pivot_data = pivot_info['data']
            pivot_headers = pivot_info['headers']

            if sheet_name not in wb.sheetnames:
                print(f"Лист '{sheet_name}' не найден, пропускаем")
                continue

            ws = wb[sheet_name]
            append_pivot_table(ws, pivot_data, pivot_headers)

        # Сохраняем результат
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        print("\n=== ФАЙЛ УСПЕШНО ОБРАБОТАН, ОТПРАВЛЯЕМ ===\n")

        return send_file(
            output,
            as_attachment=True,
            download_name=f'Оплаты_поставщикам_{datetime.now().strftime("%Y-%m-%d")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print("\n!!! ОШИБКА !!!")
        print(str(e))
        traceback.print_exc()
        return {'error': str(e)}, 500


def create_summary_sheet(ws, data):
    """Создаёт лист 'Сводные таблицы' с тремя блоками:
    1. Динамика по подразделениям
    2. Свод задолженности ДТ
    3. Свод задолженности СИ УАТ
    """
    print("Создание листа 'Сводные таблицы'...")

    current_date = data.get('currentDate', datetime.now().strftime('%Y-%m-%d'))
    previous_date = data.get('previousDate', '')
    current_day_data = data.get('currentDayData', {})
    previous_day_data = data.get('previousDayData', {})
    summary_dt = data.get('summaryDT', {})
    summary_siuat = data.get('summarySIUAT', {})

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

    summary_dt_data = [
        ('общая ДЗ', summary_dt.get('totalDebt', 0)),
        ('из них ПДЗ', summary_dt.get('totalOverdue', 0)),
        ('в т.ч. Судебная', summary_dt.get('legal', 0)),
        ('не подлежащая к взысканию', summary_dt.get('notRecoverable', 0)),
        ('подлежащая к взысканию', summary_dt.get('recoverable', 0)),
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
        row += 1

    row += 3

    # ===== ТАБЛИЦА 3: Свод задолженности СИ УАТ =====
    ws.cell(row=row, column=1, value='Свод задолженности СИ УАТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 2

    summary_siuat_data = [
        ('в т.ч. Судебная', summary_siuat.get('legal', 0)),
        ('не подлежащая к взысканию', summary_siuat.get('notRecoverable', 0)),
        ('подлежащая к взысканию', summary_siuat.get('recoverable', 0)),
    ]

    for label, value in summary_siuat_data:
        cell_label = ws.cell(row=row, column=1, value=label)
        cell_label.border = thin_border
        cell_label.font = Font(bold=True)

        cell_value = ws.cell(row=row, column=2, value=value)
        cell_value.number_format = number_format
        cell_value.border = thin_border
        cell_value.alignment = Alignment(horizontal='right')
        row += 1

    # Ширина колонок
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 18

    print("Лист 'Сводные таблицы' создан")


def append_pivot_table(ws, pivot_data, pivot_headers):
    """Добавляет сводную таблицу на существующий лист через 5 строк после данных"""

    # Находим последнюю строку с данными
    last_row = ws.max_row

    # Стартовая строка для сводной (5 пустых строк + 1 для заголовка)
    title_row = last_row + 6  # +1 переход + 5 пустых
    header_row = last_row + 7
    data_start_row = last_row + 8

    # Стили
    title_font = Font(bold=True, size=14)
    title_alignment = Alignment(horizontal='center')

    header_fill = PatternFill(start_color='FF8C00', end_color='FF8C00', fill_type='solid')
    header_font = Font(bold=True, size=11, color='FFFFFF')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    total_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    total_font = Font(bold=True, size=11)
    total_alignment = Alignment(horizontal='right')

    explanation_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    number_format = '#,##0.00'

    # Заголовок сводной таблицы
    title_cell = ws.cell(row=title_row, column=1, value='Сводная таблица оплат по подразделениям')
    title_cell.font = title_font
    title_cell.alignment = title_alignment

    # Объединение для заголовка
    last_col = len(pivot_headers) + 3  # Контрагент + подразделения + Итого + Пояснение
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=last_col)

    # Заголовки таблицы
    headers = ['Контрагент'] + pivot_headers + ['Итого', 'Пояснение']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

    # Записываем данные
    col_totals = {h: 0 for h in pivot_headers}
    grand_total = 0

    for row_idx, row_data in enumerate(pivot_data):
        current_row = data_start_row + row_idx

        # Контрагент
        ws.cell(row=current_row, column=1, value=row_data['contractor']).border = thin_border

        # Подразделения
        for col_idx, h in enumerate(pivot_headers, 2):
            value = row_data.get(h, 0)
            # Если 0 — ставим "-", иначе число
            if value == 0:
                cell = ws.cell(row=current_row, column=col_idx, value='-')
            else:
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.number_format = number_format
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right')
            col_totals[h] = col_totals.get(h, 0) + value

        # Итого по строке
        row_total = row_data.get('total', 0)
        if row_total == 0:
            total_cell = ws.cell(row=current_row, column=len(pivot_headers) + 2, value='-')
        else:
            total_cell = ws.cell(row=current_row, column=len(pivot_headers) + 2, value=row_total)
            total_cell.number_format = number_format
        total_cell.border = thin_border
        total_cell.alignment = Alignment(horizontal='right')
        total_cell.font = Font(bold=True)
        grand_total += row_total

        # Пояснение
        explanation = row_data.get('explanation', '')
        expl_cell = ws.cell(row=current_row, column=len(pivot_headers) + 3, value=explanation)
        expl_cell.border = thin_border

        # Жёлтый фон для строк с пояснениями
        if explanation:
            for col_idx in range(1, last_col + 1):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.fill = explanation_fill

    # Итоговая строка
    total_row = data_start_row + len(pivot_data)

    ws.cell(row=total_row, column=1, value='ИТОГО').font = total_font
    ws.cell(row=total_row, column=1).fill = total_fill
    ws.cell(row=total_row, column=1).alignment = total_alignment
    ws.cell(row=total_row, column=1).border = thin_border

    for col_idx, h in enumerate(pivot_headers, 2):
        col_total = col_totals[h]
        if col_total == 0:
            cell = ws.cell(row=total_row, column=col_idx, value='-')
        else:
            cell = ws.cell(row=total_row, column=col_idx, value=col_total)
            cell.number_format = number_format
        cell.font = total_font
        cell.fill = total_fill
        cell.alignment = total_alignment
        cell.border = thin_border

    grand_total_cell = ws.cell(row=total_row, column=len(pivot_headers) + 2, value=grand_total if grand_total > 0 else '-')
    if grand_total > 0:
        grand_total_cell.number_format = number_format
    grand_total_cell.font = total_font
    grand_total_cell.fill = total_fill
    grand_total_cell.alignment = total_alignment
    grand_total_cell.border = thin_border

    ws.cell(row=total_row, column=len(pivot_headers) + 3, value='').fill = total_fill
    ws.cell(row=total_row, column=len(pivot_headers) + 3).border = thin_border

    # Ширина колонок
    ws.column_dimensions['A'].width = 35
    for i in range(2, len(pivot_headers) + 2):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = 18

    last_col_letter = openpyxl.utils.get_column_letter(len(pivot_headers) + 2)
    ws.column_dimensions[last_col_letter].width = 18

    expl_col_letter = openpyxl.utils.get_column_letter(len(pivot_headers) + 3)
    ws.column_dimensions[expl_col_letter].width = 50
