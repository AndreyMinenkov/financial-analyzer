-- Схема базы данных для хранения истории сверок дебиторской задолженности
-- Финансовый анализатор - Динамика просроченной задолженности

-- Таблица 1: Сводка по дате сверки
CREATE TABLE IF NOT EXISTS swipe_history (
    id SERIAL PRIMARY KEY,
    swipe_date DATE NOT NULL UNIQUE,           -- Дата сверки (уникальная — одна сверка в день)
    total_overdue NUMERIC(20, 2) DEFAULT 0,    -- Общая просроченная задолженность
    total_debt NUMERIC(20, 2) DEFAULT 0,       -- Общая дебиторская задолженность
    filial_count INT DEFAULT 0,                -- Количество филиалов
    counterparty_count INT DEFAULT 0,          -- Количество контрагентов
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Таблица 2: Данные по филиалам (ДТ) на дату сверки
CREATE TABLE IF NOT EXISTS filial_snapshots (
    id SERIAL PRIMARY KEY,
    swipe_id INT NOT NULL REFERENCES swipe_history(id) ON DELETE CASCADE,
    swipe_date DATE NOT NULL,                   -- Денормализация для удобства запросов
    filial_name VARCHAR(255) NOT NULL,         -- Название филиала (например, "ДТ ТУРУХАНСК")
    overdue_amount NUMERIC(20, 2) DEFAULT 0,   -- Просроченная задолженность
    total_debt_amount NUMERIC(20, 2) DEFAULT 0, -- Общая задолженность
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(swipe_date, filial_name)             -- Один филиал — одна запись на дату
);

-- Таблица 3: Данные по контрагентам на дату сверки (2-й уровень детализации)
CREATE TABLE IF NOT EXISTS counterparty_snapshots (
    id SERIAL PRIMARY KEY,
    swipe_id INT NOT NULL REFERENCES swipe_history(id) ON DELETE CASCADE,
    swipe_date DATE NOT NULL,                   -- Денормализация для удобства запросов
    filial_name VARCHAR(255) NOT NULL,         -- Название филиала (родитель)
    counterparty_name VARCHAR(255) NOT NULL,   -- Название контрагента (например, "РН-Ванкор ООО")
    debt_amount NUMERIC(20, 2) DEFAULT 0,      -- Задолженность контрагента
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(swipe_date, filial_name, counterparty_name) -- Одна запись на комбинацию
);

-- Индексы для ускорения запросов
CREATE INDEX IF NOT EXISTS idx_filial_swipe_date ON filial_snapshots(swipe_date);
CREATE INDEX IF NOT EXISTS idx_filial_name ON filial_snapshots(filial_name);
CREATE INDEX IF NOT EXISTS idx_counterparty_swipe_date ON counterparty_snapshots(swipe_date);
CREATE INDEX IF NOT EXISTS idx_counterparty_filial ON counterparty_snapshots(filial_name);
CREATE INDEX IF NOT EXISTS idx_counterparty_name ON counterparty_snapshots(counterparty_name);
CREATE INDEX IF NOT EXISTS idx_swipe_history_date ON swipe_history(swipe_date);

-- Комментарий
COMMENT ON TABLE swipe_history IS 'История сверок дебиторской задолженности — сводка по дате';
COMMENT ON TABLE filial_snapshots IS 'Снимки данных по филиалам (ДТ) на дату сверки';
COMMENT ON TABLE counterparty_snapshots IS 'Снимки данных по контрагентам на дату сверки (2-й уровень)';
