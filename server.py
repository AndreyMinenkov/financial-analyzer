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
CORS(app, origins="*", supports_credentials=True)

if 'RENDER' in os.environ:
    os.environ['WEB_CONCURRENCY'] = '1'

# ... (COLUMNS, SUM_COLUMNS, NUMERIC_COLUMNS остаются теми же)

# Все функции (is_cell_merged, get_cell_to_write и т.д.) остаются без изменений
# Но recalc_totals нужно переписать для оптимизации памяти

def recalc_totals_optimized(ws, updated_rows):
    """Пересчитывает итоги только для затронутых строк"""
    print("\n=== ПЕРЕСЧЁТ ИТОГОВ (ОПТИМИЗИРОВАННЫЙ) ===")
    
    filials, kontragents, dogovors, documents, total_row = find_structure(ws)
    
    # Создаем множества для быстрого поиска
    filials_set = set(filials)
    kontragents_set = set(kontragents)
    dogovors_set = set(dogovors)
    documents_set = set(documents)
    
    # Находим все родительские строки, которые нужно пересчитать
    rows_to_recalc = set()
    
    for doc_row in updated_rows:
        if doc_row in documents_set:
            rows_to_recalc.add(doc_row)
            # Ищем родительский договор
            for r in range(doc_row - 1, 13, -1):
                if r in dogovors_set:
                    rows_to_recalc.add(r)
                    # Ищем родительского контрагента
                    for k in range(r - 1, 13, -1):
                        if k in kontragents_set:
                            rows_to_recalc.add(k)
                            # Ищем родительский филиал
                            for f in range(k - 1, 13, -1):
                                if f in filials_set:
                                    rows_to_recalc.add(f)
                                    break
                            break
                    break
    
    print(f"Нужно пересчитать строки: {sorted(rows_to_recalc)}")
    
    # Функция для суммирования значений
    def sum_children(parent_row, child_type, child_set):
        total = {col: 0 for col in SUM_COLUMNS}
        max_day = 0
        
        for r in range(parent_row + 1, ws.max_row + 1):
            if r in filials_set or r in kontragents_set or r in dogovors_set or r == total_row:
                break
            if r in child_set:
                for col in SUM_COLUMNS:
                    val = get_cell_value(ws, r, col)
                    if isinstance(val, (int, float)):
                        total[col] += val
                day_val = get_cell_value(ws, r, COLUMNS['DAYS'])
                if isinstance(day_val, (int, float)) and day_val > max_day:
                    max_day = day_val
        
        return total, max_day
    
    # Пересчитываем только нужные строки
    for row in sorted(rows_to_recalc, reverse=True):
        if row in dogovors_set:
            children = [d for d in documents if d > row and d < (next((r for r in sorted(dogovors_set) if r > row), ws.max_row + 1))]
            totals, max_day = sum_children(row, documents, documents_set)
            for col, val in totals.items():
                safe_set_number_format(ws, row, col, val)
            safe_set_value(ws, row, COLUMNS['DAYS'], max_day)
            print(f"Договор {row}: пересчитан")
            
        elif row in kontragents_set:
            children = [d for d in dogovors if d > row and d < (next((r for r in sorted(kontragents_set) if r > row), ws.max_row + 1))]
            totals, max_day = sum_children(row, dogovors, dogovors_set)
            for col, val in totals.items():
                safe_set_number_format(ws, row, col, val)
            safe_set_value(ws, row, COLUMNS['DAYS'], max_day)
            print(f"Контрагент {row}: пересчитан")
            
        elif row in filials_set:
            children = [k for k in kontragents if k > row and k < (next((r for r in sorted(filials_set) if r > row), ws.max_row + 1))]
            totals, max_day = sum_children(row, kontragents, kontragents_set)
            for col, val in totals.items():
                safe_set_number_format(ws, row, col, val)
            safe_set_value(ws, row, COLUMNS['DAYS'], max_day)
            print(f"Филиал {row}: пересчитан")
    
    # Пересчитываем общий итог если нужно
    if total_row and filials:
        totals, max_day = sum_children(total_row, filials, filials_set)
        for col, val in totals.items():
            safe_set_number_format(ws, total_row, col, val)
        safe_set_value(ws, total_row, COLUMNS['DAYS'], max_day)
        print(f"Общий итог {total_row}: пересчитан")

@app.route('/', methods=['GET'])
def index():
    return jsonify({
        'status': 'ok',
        'message': 'Сервер сверки долгов (ultra light)',
        'endpoints': {'/save-excel': 'POST'}
    })

@app.route('/save-excel', methods=['POST', 'OPTIONS'])
def save_excel():
    if request.method == 'OPTIONS':
        return app.make_default_options_response()
    
    try:
        gc.collect()
        
        file = request.files['file']
        data = json.loads(request.form['data'])
        
        print(f"\n=== ПОЛУЧЕН ЗАПРОС ===")
        print(f"Файл: {file.filename}")
        print(f"Документов для обновления: {len(data['updatedDocuments'])}")
        
        # Загружаем Excel
        wb = openpyxl.load_workbook(io.BytesIO(file.read()), data_only=True)
        ws = wb.active
        
        today = datetime.now().date()
        _, _, _, _, total_row = find_structure(ws)
        
        updated_rows = set()
        
        # Обновляем документы
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
            else:
                days_overdue = (today - expected_date).days
                interval_col = get_interval_col(days_overdue)
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], debt_amount)
                safe_set_value(ws, row_number, COLUMNS['DAYS'], days_overdue)
                safe_set_number_format(ws, row_number, interval_col, debt_amount)
                updated_rows.add(row_number)
        
        if updated_rows:
            recalc_totals_optimized(ws, updated_rows)
            if total_row:
                update_top_table(ws, total_row)
        
        align_numeric_cells(ws)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Чистим память
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
