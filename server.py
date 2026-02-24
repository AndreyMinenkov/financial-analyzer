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

app = Flask(__name__)

# Максимально разрешаем CORS
CORS(app, origins="*", supports_credentials=True)

# Оптимизация для Render - один воркер и отключаем ненужное
if 'RENDER' in os.environ:
    os.environ['WEB_CONCURRENCY'] = '1'
    # Отключаем доступ к метрикам psutil
    os.environ['GUNICORN_CMD_ARGS'] = '--max-requests 1 --max-requests-jitter 10'

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

# Остальные функции без изменений (is_cell_merged, get_cell_to_write, и т.д.)
# ... (весь код функций остается таким же, как в предыдущей версии, 
#      просто убираем импорт psutil и все обращения к нему)

@app.route('/', methods=['GET'])
def index():
    """Корневой маршрут для проверки"""
    return jsonify({
        'status': 'ok',
        'message': 'Сервер сверки долгов работает (light version)',
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
        
        file = request.files['file']
        data = json.loads(request.form['data'])
        
        print(f"\n=== ПОЛУЧЕН ЗАПРОС ===")
        print(f"Файл: {file.filename}")
        print(f"Документов для обновления: {len(data['updatedDocuments'])}")
        
        # Загружаем Excel с минимальными настройками
        file_bytes = file.read()
        
        # Освобождаем file из памяти
        del file
        
        wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes),
            read_only=False,
            keep_vba=False,
            data_only=True
        )
        
        # Освобождаем file_bytes
        del file_bytes
        
        ws = wb.active
        today = datetime.now().date()
        
        # Находим итоговую строку
        _, _, _, _, total_row = find_structure(ws)
        
        # Обновляем документы
        updated_rows = set()
        for item in data['updatedDocuments']:
            row_number = item['rowNumber']
            debt_amount = float(item['amount'])
            expected_date_str = item['date']
            
            expected_date = None
            if expected_date_str and expected_date_str not in ('null', 'None', ''):
                try:
                    clean_str = expected_date_str.strip().strip('"').strip("'")
                    if clean_str:
                        expected_date = datetime.fromisoformat(clean_str.replace('Z', '+00:00')).date()
                except Exception:
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
            recalc_totals(ws)
            if total_row:
                update_top_table(ws, total_row)
        
        align_numeric_cells(ws)
        
        # Сохраняем результат
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Явно удаляем всё для освобождения памяти
        del wb
        del ws
        del data
        gc.collect()
        
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
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 Сервер запущен на порту {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
