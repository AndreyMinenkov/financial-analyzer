# db.py — Модуль работы с PostgreSQL для хранения истории сверок
import psycopg2
import psycopg2.extras
from datetime import date, datetime
from decimal import Decimal
import json

# Настройки подключения
DB_CONFIG = {
    'host': '127.0.0.1',
    'port': 5432,
    'database': 'financial_analyzer',
    'user': 'postgres',
    'password': 'Kapapa661109'
}


def get_connection():
    """Создаёт новое подключение к БД"""
    return psycopg2.connect(**DB_CONFIG)


def save_swipe_data(swipe_date, filial_data, counterparty_data, total_debt=0, total_overdue=0):
    """
    Сохраняет данные сверки в БД.
    
    Args:
        swipe_date: str — дата сверки (YYYY-MM-DD)
        filial_data: dict — {filial_name: overdue_amount}
        counterparty_data: dict — {(filial, counterparty): debt_amount}
        total_debt: float — общая ДЗ
        total_overdue: float — общая ПДЗ
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        # 1. Сохраняем или обновляем сводку
        cur.execute("""
            INSERT INTO swipe_history (swipe_date, total_overdue, total_debt, filial_count, counterparty_count, updated_at)
            VALUES (%s, %s, %s, %s, %s, NOW())
            ON CONFLICT (swipe_date) DO UPDATE SET
                total_overdue = EXCLUDED.total_overdue,
                total_debt = EXCLUDED.total_debt,
                filial_count = EXCLUDED.filial_count,
                counterparty_count = EXCLUDED.counterparty_count,
                updated_at = NOW()
            RETURNING id
        """, (swipe_date, total_overdue, total_debt, len(filial_data), len(counterparty_data)))

        swipe_row = cur.fetchone()
        swipe_id = swipe_row[0] if swipe_row else None
        
        if not swipe_id:
            raise Exception("Не удалось получить ID сверки")
        
        conn.commit()

        # 2. Сохраняем данные по филиалам
        filial_rows = []
        for filial_name, overdue in filial_data.items():
            filial_rows.append((swipe_id, swipe_date, filial_name, float(overdue), 0))

        if filial_rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO filial_snapshots (swipe_id, swipe_date, filial_name, overdue_amount, total_debt_amount)
                VALUES %s
                ON CONFLICT (swipe_date, filial_name) DO UPDATE SET
                    swipe_id = EXCLUDED.swipe_id,
                    overdue_amount = EXCLUDED.overdue_amount,
                    total_debt_amount = EXCLUDED.total_debt_amount
                """,
                filial_rows
            )
        conn.commit()

        # 3. Сохраняем данные по контрагентам
        cp_rows = []
        for (filial_name, cp_name), debt in counterparty_data.items():
            cp_rows.append((swipe_id, swipe_date, filial_name, cp_name, float(debt)))

        if cp_rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO counterparty_snapshots (swipe_id, swipe_date, filial_name, counterparty_name, debt_amount)
                VALUES %s
                ON CONFLICT (swipe_date, filial_name, counterparty_name) DO UPDATE SET
                    swipe_id = EXCLUDED.swipe_id,
                    debt_amount = EXCLUDED.debt_amount
                """,
                cp_rows
            )
        conn.commit()

        cur.close()
        print(f"✅ Данные сверки за {swipe_date} сохранены: {len(filial_data)} филиалов, {len(counterparty_data)} контрагентов")

        return {'success': True, 'swipe_id': swipe_id}

    except Exception as e:
        print(f"❌ Ошибка сохранения данных сверки: {e}")
        if conn:
            conn.rollback()
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            conn.close()


def get_swipe_dates(from_date=None, to_date=None):
    """
    Возвращает список дат сверок.
    
    Returns:
        list of dict — [{'date': '2026-04-15', 'total_overdue': 586395480.60, 'total_debt': ..., 'filial_count': 29}, ...]
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = "SELECT swipe_date, total_overdue, total_debt, filial_count, counterparty_count, created_at FROM swipe_history"
        params = []
        conditions = []

        if from_date:
            conditions.append("swipe_date >= %s")
            params.append(from_date)
        if to_date:
            conditions.append("swipe_date <= %s")
            params.append(to_date)

        if conditions:
            query += " WHERE " + " AND ".join(conditions)

        query += " ORDER BY swipe_date ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        result = []
        for row in rows:
            result.append({
                'date': row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date']),
                'total_overdue': float(row['total_overdue']) if row['total_overdue'] else 0,
                'total_debt': float(row['total_debt']) if row['total_debt'] else 0,
                'filial_count': row['filial_count'],
                'counterparty_count': row['counterparty_count'],
            })

        cur.close()
        return result

    except Exception as e:
        print(f"❌ Ошибка чтения дат сверок: {e}")
        return []
    finally:
        if conn:
            conn.close()


def get_filial_trend(from_date, to_date, filial_name=None):
    """
    Возвращает данные для графика динамики ПДЗ по филиалам.
    
    Args:
        from_date: str — дата начала периода
        to_date: str — дата конца периода
        filial_name: str или None — если указан, только этот филиал; если None — все филиалы
    
    Returns:
        {
            'dates': ['2026-04-10', '2026-04-11', ...],
            'series': [
                {'name': 'ДТ ТУРУХАНСК', 'data': [258127335.10, 260000000, ...]},
                {'name': 'ДТ УК', 'data': [...]}
            ]
        }
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = """
            SELECT swipe_date, filial_name, overdue_amount 
            FROM filial_snapshots 
            WHERE swipe_date >= %s AND swipe_date <= %s
        """
        params = [from_date, to_date]

        if filial_name:
            query += " AND filial_name = %s"
            params.append(filial_name)

        query += " ORDER BY swipe_date ASC, filial_name ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        # Группируем по датам и филиалам
        dates_set = set()
        filial_data = {}  # filial_name -> {date -> amount}

        for row in rows:
            d = row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date'])
            fn = row['filial_name']
            amount = float(row['overdue_amount']) if row['overdue_amount'] else 0

            dates_set.add(d)
            if fn not in filial_data:
                filial_data[fn] = {}
            filial_data[fn][d] = amount

        # Сортируем даты
        dates = sorted(dates_set)

        # Формируем series
        series = []
        for fn, data in sorted(filial_data.items()):
            series.append({
                'name': fn,
                'data': [data.get(d, 0) for d in dates]
            })

        cur.close()
        return {
            'dates': dates,
            'series': series
        }

    except Exception as e:
        print(f"❌ Ошибка чтения тренда филиалов: {e}")
        return {'dates': [], 'series': []}
    finally:
        if conn:
            conn.close()


def get_counterparty_trend(from_date, to_date, filial_name=None, counterparty_name=None):
    """
    Возвращает данные для графика динамики по контрагентам.
    
    Returns:
        {
            'dates': [...],
            'series': [
                {'name': 'РН-Ванкор ООО (ДТ ТУРУХАНСК)', 'data': [...]},
                ...
            ]
        }
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = """
            SELECT swipe_date, filial_name, counterparty_name, debt_amount 
            FROM counterparty_snapshots 
            WHERE swipe_date >= %s AND swipe_date <= %s
        """
        params = [from_date, to_date]

        if filial_name:
            query += " AND filial_name = %s"
            params.append(filial_name)
        if counterparty_name:
            query += " AND counterparty_name = %s"
            params.append(counterparty_name)

        query += " ORDER BY swipe_date ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        dates_set = set()
        cp_data = {}  # "контрагент (филиал)" -> {date -> amount}

        for row in rows:
            d = row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date'])
            fn = row['filial_name']
            cp = row['counterparty_name']
            amount = float(row['debt_amount']) if row['debt_amount'] else 0

            label = f"{cp} ({fn})"
            dates_set.add(d)
            if label not in cp_data:
                cp_data[label] = {}
            cp_data[label][d] = amount

        dates = sorted(dates_set)

        series = []
        for label, data in sorted(cp_data.items()):
            series.append({
                'name': label,
                'data': [data.get(d, 0) for d in dates]
            })

        cur.close()
        return {
            'dates': dates,
            'series': series
        }

    except Exception as e:
        print(f"❌ Ошибка чтения тренда контрагентов: {e}")
        return {'dates': [], 'series': []}
    finally:
        if conn:
            conn.close()


def get_filial_list(from_date=None, to_date=None):
    """Возвращает список уникальных филиалов за период"""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        query = "SELECT DISTINCT filial_name FROM filial_snapshots"
        params = []
        conditions = []

        if from_date:
            conditions.append("swipe_date >= %s")
            params.append(from_date)
        if to_date:
            conditions.append("swipe_date <= %s")
            params.append(to_date)

        if conditions:
            query += " WHERE " + " AND ".join(conditions)

        query += " ORDER BY filial_name ASC"

        cur.execute(query, params)
        result = [row[0] for row in cur.fetchall()]
        cur.close()
        return result

    except Exception as e:
        print(f"❌ Ошибка чтения списка филиалов: {e}")
        return []
    finally:
        if conn:
            conn.close()


def get_counterparty_list(filial_name=None):
    """Возвращает список уникальных контрагентов (опционально фильтруя по филиалу)"""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        query = "SELECT DISTINCT counterparty_name FROM counterparty_snapshots"
        params = []

        if filial_name:
            query += " WHERE filial_name = %s"
            params.append(filial_name)

        query += " ORDER BY counterparty_name ASC"

        cur.execute(query, params)
        result = [row[0] for row in cur.fetchall()]
        cur.close()
        return result

    except Exception as e:
        print(f"❌ Ошибка чтения списка контрагентов: {e}")
        return []
    finally:
        if conn:
            conn.close()


def get_summary(from_date, to_date):
    """
    Возвращает сводную статистику за период.
    
    Returns:
        {
            'min_overdue': ...,
            'max_overdue': ...,
            'avg_overdue': ...,
            'latest_overdue': ...,
            'trend': 'up' | 'down' | 'stable',
            'swipe_count': ...
        }
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT 
                COUNT(*) as swipe_count,
                MIN(total_overdue) as min_overdue,
                MAX(total_overdue) as max_overdue,
                AVG(total_overdue) as avg_overdue,
                (SELECT total_overdue FROM swipe_history WHERE swipe_date <= %s ORDER BY swipe_date DESC LIMIT 1) as first_overdue,
                (SELECT total_overdue FROM swipe_history WHERE swipe_date <= %s ORDER BY swipe_date DESC LIMIT 1) as last_overdue
            FROM swipe_history 
            WHERE swipe_date >= %s AND swipe_date <= %s
        """, (from_date, to_date, from_date, to_date))

        row = cur.fetchone()
        cur.close()

        if not row or row['swipe_count'] == 0:
            return None

        first = float(row['first_overdue']) if row['first_overdue'] else 0
        last = float(row['last_overdue']) if row['last_overdue'] else 0

        if first > 0 and last > 0:
            change = ((last - first) / first) * 100
            trend = 'up' if change > 5 else ('down' if change < -5 else 'stable')
        else:
            trend = 'stable'

        return {
            'min_overdue': float(row['min_overdue']) if row['min_overdue'] else 0,
            'max_overdue': float(row['max_overdue']) if row['max_overdue'] else 0,
            'avg_overdue': float(row['avg_overdue']) if row['avg_overdue'] else 0,
            'latest_overdue': last,
            'trend': trend,
            'swipe_count': row['swipe_count'],
        }

    except Exception as e:
        print(f"❌ Ошибка чтения сводки: {e}")
        return None
    finally:
        if conn:
            conn.close()


def delete_swipe_data(swipe_date):
    """Удаляет данные сверки за указанную дату (CASCADE удалит связанные записи)"""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("DELETE FROM swipe_history WHERE swipe_date = %s RETURNING id", (swipe_date,))
        deleted = cur.fetchone()
        conn.commit()
        cur.close()

        if deleted:
            print(f"✅ Удалены данные за {swipe_date}")
            return {'success': True, 'message': f'Данные за {swipe_date} удалены'}
        else:
            return {'success': False, 'message': f'Данные за {swipe_date} не найдены'}

    except Exception as e:
        print(f"❌ Ошибка удаления: {e}")
        if conn:
            conn.rollback()
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            conn.close()
