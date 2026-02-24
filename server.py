from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import numbers, Alignment
import io
from datetime import datetime
import json
import traceback
import os
import gc
import psutil

app = Flask(__name__)

# Настройка CORS для продакшена
CORS(app, origins=["null", "file://", "http://localhost:*", "https://*.onrender.com"])

# Оптимизация для Render - один воркер
if 'RENDER' in os.environ:
    os.environ['WEB_CONCURRENCY'] = '1'
    print("✅ Render detected: WEB_CONCURRENCY set to 1")

# Список целевых контрагентов
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
    cell.alignment = Alignment(horizontal='right')

def align_numeric_cells(ws):
    """Выравнивает все числовые ячейки по правому краю"""
    print("Выравнивание числовых ячеек по правому краю...")
    
    for row in range(14, ws.max_row + 1):
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
    cell.number_format = '#,##0.00' if col == 5 else '0.00%'
    cell.alignment = Alignment(horizontal='right')

def get_interval_col(days):
    """Возвращает колонку для указанного количества дней просрочки"""
    if days <= 0:
        return COLUMNS['NOT_OVERDUE']
    elif 1 <= days <= 15:
        return COLUMNS['INTERVAL_1_15']
    elif 16 <= days <= 29:
        return COLUMNS['INTERVAL_16_29']
    elif 30 <= days <= 89:
        return COLUMNS['INTERVAL_30_89']
    elif 90 <= days <= 179:
        return COLUMNS['INTERVAL_90_179']
    else:
        return COLUMNS['INTERVAL_180_PLUS']

def clear_all_intervals(ws, row):
    """Очищает все интервальные колонки для строки"""
    if row <= 13:
        return
    interval_cols = [
        COLUMNS['NOT_OVERDUE'],
        COLUMNS['INTERVAL_1_15'],
        COLUMNS['INTERVAL_16_29'],
        COLUMNS['INTERVAL_30_89'],
        COLUMNS['INTERVAL_90_179'],
        COLUMNS['INTERVAL_180_PLUS'],
    ]
    for col in interval_cols:
        safe_set_number_format(ws, row, col, 0)

def find_structure(ws):
    """Находит все строки филиалов, контрагентов, договоров и документов"""
    filials = []
    kontragents = []
    dogovors = []
    documents = []
    total_row = None
    
    for row in range(14, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        if not cell_value:
            continue
        
        str_val = str(cell_value).strip()
        
        if str_val.startswith('ДТ '):
            filials.append(row)
        elif ('ООО' in str_val or 'АО' in str_val) and not str_val.startswith('ДТ '):
            kontragents.append(row)
        elif str_val.startswith('Договор') or 'договор' in str_val.lower():
            dogovors.append(row)
        elif any(x in str_val for x in ['Акт', 'Реализация', 'Корректировка', 'Поступление']):
            documents.append(row)
        elif 'Итого' in str_val or 'ИТОГО' in str_val:
            total_row = row
    
    return filials, kontragents, dogovors, documents, total_row

def recalc_totals(ws):
    """Пересчитывает все итоговые строки в файле с учётом иерархии"""
    print("\n=== ПЕРЕСЧЁТ ИТОГОВ ===")
    
    filials, kontragents, dogovors, documents, total_row = find_structure(ws)
    
    print(f"Найдено: филиалов={len(filials)}, контрагентов={len(kontragents)}, "
          f"договоров={len(dogovors)}, документов={len(documents)}")
    
    def sum_rows(rows, col):
        total = 0
        for r in rows:
            val = get_cell_value(ws, r, col)
            if isinstance(val, (int, float)):
                total += val
        return total
    
    def max_days_in_rows(rows):
        max_val = 0
        for r in rows:
            val = get_cell_value(ws, r, COLUMNS['DAYS'])
            if isinstance(val, (int, float)) and val > max_val:
                max_val = val
        return max_val
    
    # Пересчет договоров
    for i, dog_row in enumerate(dogovors):
        doc_rows = []
        for r in range(dog_row + 1, ws.max_row + 1):
            if r in dogovors or r in kontragents or r in filials or r == total_row:
                break
            if r in documents:
                doc_rows.append(r)
        
        if doc_rows:
            print(f"Договор стр.{dog_row}: документы {doc_rows}")
            for col in SUM_COLUMNS:
                total = sum_rows(doc_rows, col)
                safe_set_number_format(ws, dog_row, col, total)
            max_day = max_days_in_rows(doc_rows)
            safe_set_value(ws, dog_row, COLUMNS['DAYS'], max_day)
    
    # Пересчет контрагентов
    for i, kontr_row in enumerate(kontragents):
        dog_rows = []
        for r in range(kontr_row + 1, ws.max_row + 1):
            if r in kontragents or r in filials or r == total_row:
                break
            if r in dogovors:
                dog_rows.append(r)
        
        if dog_rows:
            print(f"Контрагент стр.{kontr_row}: договоры {dog_rows}")
            for col in SUM_COLUMNS:
                total = sum_rows(dog_rows, col)
                safe_set_number_format(ws, kontr_row, col, total)
            max_day = max_days_in_rows(dog_rows)
            safe_set_value(ws, kontr_row, COLUMNS['DAYS'], max_day)
    
    # Пересчет филиалов
    for i, fil_row in enumerate(filials):
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
    
    # Пересчет общего итога
    if total_row and filials:
        print(f"Общий итог стр.{total_row}: филиалы {filials}")
        for col in SUM_COLUMNS:
            total = sum_rows(filials, col)
            safe_set_number_format(ws, total_row, col, total)
        max_day = max_days_in_rows(filials)
        safe_set_value(ws, total_row, COLUMNS['DAYS'], max_day)
    
    print("=== ПЕРЕСЧЁТ ИТОГОВ ЗАВЕРШЕН ===\n")

def update_top_table(ws, total_row):
    """Обновляет верхнюю сводную таблицу (строки 1-8)"""
    print("\n=== ОБНОВЛЕНИЕ ВЕРХНЕЙ ТАБЛИЦЫ ===")
    
    if not total_row:
        print("Не найдена итоговая строка")
        return
    
    t_value = get_cell_value(ws, total_row, COLUMNS['NOT_OVERDUE']) or 0
    u_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_1_15']) or 0
    v_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_16_29']) or 0
    w_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_30_89']) or 0
    x_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_90_179']) or 0
    y_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_180_PLUS']) or 0
    l_value = get_cell_value(ws, total_row, COLUMNS['DEBT_AMOUNT']) or 0
    
    print(f"Значения из строки {total_row}: T={t_value}, U={u_value}, V={v_value}, W={w_value}, X={x_value}, Y={y_value}, L={l_value}")
    
    safe_set_top_table_value(ws, 2, 5, t_value)
    safe_set_top_table_value(ws, 2, 6, t_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 3, 5, u_value)
    safe_set_top_table_value(ws, 3, 6, u_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 4, 5, v_value)
    safe_set_top_table_value(ws, 4, 6, v_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 5, 5, w_value)
    safe_set_top_table_value(ws, 5, 6, w_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 6, 5, x_value)
    safe_set_top_table_value(ws, 6, 6, x_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 7, 5, y_value)
    safe_set_top_table_value(ws, 7, 6, y_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 8, 5, l_value)
    safe_set_top_table_value(ws, 8, 6, 1.0)
    
    print("Верхняя таблица обновлена")
    print("=== ОБНОВЛЕНИЕ ВЕРХНЕЙ ТАБЛИЦЫ ЗАВЕРШЕНО ===\n")

@app.route('/', methods=['GET'])
def index():
    """Корневой маршрут для проверки"""
    memory_usage = psutil.Process().memory_info().rss / 1024 / 1024
    return jsonify({
        'status': 'ok',
        'message': 'Сервер сверки долгов работает (оптимизированная версия)',
        'memory_mb': round(memory_usage, 2),
        'endpoints': {
            '/save-excel': 'POST - обновление Excel файла',
            '/': 'GET - проверка сервера'
        }
    }), 200

@app.route('/save-excel', methods=['POST', 'OPTIONS'])
def save_excel():
    """Основной эндпоинт для обработки Excel файлов"""
    if request.method == 'OPTIONS':
        response = app.make_default_options_response()
        return response
    
    try:
        # Принудительная сборка мусора перед началом
        gc.collect()
        
        # Логируем использование памяти
        memory_before = psutil.Process().memory_info().rss / 1024 / 1024
        print(f"\n=== ПАМЯТЬ ДО ОБРАБОТКИ: {memory_before:.2f} MB ===")
        
        file = request.files['file']
        data = json.loads(request.form['data'])
        
        print(f"\n=== ПОЛУЧЕН ЗАПРОС ===")
        print(f"Файл: {file.filename}")
        print(f"Документов для обновления: {len(data['updatedDocuments'])}")
        
        # Загружаем Excel с минимальными настройками
        file_bytes = file.read()
        wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes),
            read_only=False,
            keep_vba=False,
            data_only=True
        )
        ws = wb.active
        
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
            if expected_date_str and expected_date_str not in ('null', 'None', ''):
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
        
        # Сохраняем результат
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Явно удаляем workbook для освобождения памяти
        del wb
        del ws
        
        # Принудительная сборка мусора после обработки
        gc.collect()
        
        memory_after = psutil.Process().memory_info().rss / 1024 / 1024
        print(f"=== ПАМЯТЬ ПОСЛЕ ОБРАБОТКИ: {memory_after:.2f} MB ===\n")
        
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
    # Добавляем зависимость psutil в requirements.txt если её нет
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 Сервер запущен на порту {port}")
    print(f"📊 Доступная память: {psutil.virtual_memory().available / 1024 / 1024:.2f} MB")
    app.run(host='0.0.0.0', port=port, debug=False)
