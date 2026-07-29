// reports-manager.js — Управление страницей отчётов и дашбордов (с вкладками)
class ReportsManager {
    constructor() {
        console.log('📊 ReportsManager: конструктор вызван');
        this.charts = {};
        this.apiBase = 'http://31.130.155.16:5000';
        this.currentSection = null;
        this.currentTab = 'overview';
        this.selectedFilials = new Set();
        this.allFilialData = [];
        
        // Параметры детализации
        this.detailMainPeriod = 'month';
        this.detailBreakdown = 'decade';
        
        // Допустимые комбинации период/детализация
        this.validCombinations = {
            decade: ['day'],
            month: ['decade', 'day'],
            quarter: ['month', 'decade'],
            year: ['quarter', 'month']
        };
        
        // Кэш данных
        this.rawSwipeData = null;
        this.filialsDetailData = null;
    }

    // ===== ИНИЦИАЛИЗАЦИЯ =====
    init() {
        console.log('📊 ReportsManager: init() вызван');
        this.setupListeners();
        this.setDefaultDates();
        this.setupDetailFilters();
        console.log('📊 ReportsManager: init() завершён');
    }

    setupListeners() {
        console.log('📊 ReportsManager: setupListeners() вызван');
        
        // Открытие раздела "Дебиторка"
        const card = document.getElementById('reportCardDebt');
        if (card) {
            card.addEventListener('click', () => this.openDashboard('debt'));
        }

        // Кнопка "Назад"
        const backBtn = document.getElementById('backToReportsBtn');
        if (backBtn) {
            backBtn.addEventListener('click', () => this.closeDashboard());
        }

        // Кнопка "Сформировать"
        const buildBtn = document.getElementById('buildDashboardBtn');
        if (buildBtn) {
            buildBtn.addEventListener('click', () => this.buildDashboard());
        }

        // При смене филиала — обновить список контрагентов
        const filialSelect = document.getElementById('dashFilialSelect');
        if (filialSelect) {
            filialSelect.addEventListener('change', () => this.loadCounterpartyList());
        }

        // ===== ВКЛАДКИ ДАШБОРДА =====
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const tab = e.currentTarget.dataset.tab;
                this.switchTab(tab);
            });
        });

        // ===== ВКЛАДКА "СРАВНЕНИЕ": УПРАВЛЕНИЕ ЧЕКБОКСАМИ =====
        const selectTop5Btn = document.getElementById('selectTop5Btn');
        if (selectTop5Btn) {
            selectTop5Btn.addEventListener('click', () => this.selectTop5Filials());
        }

        const selectAllBtn = document.getElementById('selectAllFilialsBtn');
        if (selectAllBtn) {
            selectAllBtn.addEventListener('click', () => this.selectAllFilials());
        }

        const clearSelectionBtn = document.getElementById('clearSelectionBtn');
        if (clearSelectionBtn) {
            clearSelectionBtn.addEventListener('click', () => this.clearFilialSelection());
        }

        // Экспорт таблицы детализации
        const exportFilialsBtn = document.getElementById('exportFilialsBtn');
        if (exportFilialsBtn) {
            exportFilialsBtn.addEventListener('click', () => this.exportFilialsTable());
        }

        console.log('📊 ReportsManager: setupListeners() завершён');
    }

    // ===== НАСТРОЙКА ФИЛЬТРОВ ДЕТАЛИЗАЦИИ =====
    setupDetailFilters() {
        const mainPeriodSelect = document.getElementById('detailMainPeriod');
        const breakdownSelect = document.getElementById('detailBreakdown');
        const applyBtn = document.getElementById('applyDetailFiltersBtn');

        if (mainPeriodSelect) {
            mainPeriodSelect.addEventListener('change', (e) => {
                this.detailMainPeriod = e.target.value;
                this.updateBreakdownOptions();
            });
        }

        if (breakdownSelect) {
            breakdownSelect.addEventListener('change', (e) => {
                this.detailBreakdown = e.target.value;
            });
        }

        if (applyBtn) {
            applyBtn.addEventListener('click', () => {
                this.buildDashboard();
            });
        }

        // Инициализация опций детализации
        this.updateBreakdownOptions();
    }

    updateBreakdownOptions() {
        const breakdownSelect = document.getElementById('detailBreakdown');
        if (!breakdownSelect) return;

        const validOptions = this.validCombinations[this.detailMainPeriod] || ['decade'];
        const currentValue = breakdownSelect.value;

        breakdownSelect.innerHTML = '';
        validOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = this.getBreakdownLabel(opt);
            breakdownSelect.appendChild(option);
        });

        // Сохраняем текущее значение если оно валидно
        if (validOptions.includes(currentValue)) {
            breakdownSelect.value = currentValue;
        }
    }

    getBreakdownLabel(type) {
        const labels = {
            day: 'День',
            decade: 'Декада',
            month: 'Месяц',
            quarter: 'Квартал'
        };
        return labels[type] || type;
    }

    // ===== ПЕРЕКЛЮЧЕНИЕ ВКЛАДОК =====
    switchTab(tabId) {
        console.log('📊 ReportsManager: переключение вкладки:', tabId);
        
        // Обновляем кнопки
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabId);
        });

        // Обновляем контент
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.toggle('active', content.id === `tab-${tabId}`);
        });

        this.currentTab = tabId;

        // Загружаем данные для вкладки при необходимости
        if (tabId === 'details' && this.rawSwipeData) {
            this.renderFilialsDetailTable();
        } else if (tabId === 'comparison') {
            this.renderFilialCheckboxes();
        }
    }

    setDefaultDates() {
        const today = new Date();
        const weekAgo = new Date(today);
        weekAgo.setDate(today.getDate() - 30);

        const toDateEl = document.getElementById('dashToDate');
        const fromDateEl = document.getElementById('dashFromDate');
        
        if (toDateEl) toDateEl.value = this.formatDateISO(today);
        if (fromDateEl) fromDateEl.value = this.formatDateISO(weekAgo);
    }

    formatDateISO(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    // ===== НАВИГАЦИЯ =====
    openDashboard(section) {
        console.log('📊 ReportsManager: openDashboard() вызван, section:', section);
        this.currentSection = section;
        
        const reportCard = document.getElementById('reportCardDebt');
        const dashboardPanel = document.getElementById('dashboardPanel');
        
        if (reportCard) reportCard.style.display = 'none';
        if (dashboardPanel) dashboardPanel.style.display = 'block';
        
        this.loadFilialList();
    }

    closeDashboard() {
        console.log('📊 ReportsManager: closeDashboard() вызван');
        this.currentSection = null;
        const reportCard = document.getElementById('reportCardDebt');
        const dashboardPanel = document.getElementById('dashboardPanel');
        if (reportCard) reportCard.style.display = 'flex';
        if (dashboardPanel) dashboardPanel.style.display = 'none';
    }

    // ===== ЗАГРУЗКА СПИСКОВ ФИЛЬТРОВ =====
    async loadFilialList() {
        console.log('📊 ReportsManager: loadFilialList() вызван');
        try {
            const from = document.getElementById('dashFromDate')?.value;
            const to = document.getElementById('dashToDate')?.value;
            const url = `${this.apiBase}/api/filial-list?from=${from}&to=${to}`;
            
            const resp = await fetch(url);
            const json = await resp.json();

            const select = document.getElementById('dashFilialSelect');
            if (!select) return;
            
            const currentValue = select.value;
            select.innerHTML = '<option value="">— Все филиалы —</option>';

            if (json.success && json.data) {
                json.data.forEach(name => {
                    const opt = document.createElement('option');
                    opt.value = name;
                    opt.textContent = name;
                    select.appendChild(opt);
                });
            }

            select.value = currentValue;
            await this.loadCounterpartyList();
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки списка филиалов:', e);
        }
    }

    async loadCounterpartyList() {
        console.log('📊 ReportsManager: loadCounterpartyList() вызван');
        try {
            const filial = document.getElementById('dashFilialSelect')?.value;
            let url = `${this.apiBase}/api/counterparty-list`;
            if (filial) url += `?filial=${encodeURIComponent(filial)}`;
            
            const resp = await fetch(url);
            const json = await resp.json();

            const select = document.getElementById('dashCounterpartySelect');
            if (!select) return;
            
            const currentValue = select.value;
            select.innerHTML = '<option value="">— Все контрагенты —</option>';

            if (json.success && json.data) {
                json.data.forEach(name => {
                    const opt = document.createElement('option');
                    opt.value = name;
                    opt.textContent = name;
                    select.appendChild(opt);
                });
            }

            select.value = currentValue;
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки списка контрагентов:', e);
        }
    }

    // ===== ПОСТРОЕНИЕ ДАШБОРДА =====
    async buildDashboard() {
        console.log('📊 ReportsManager: buildDashboard() вызван');
        const from = document.getElementById('dashFromDate')?.value;
        const to = document.getElementById('dashToDate')?.value;
        const filial = document.getElementById('dashFilialSelect')?.value;
        const counterparty = document.getElementById('dashCounterpartySelect')?.value;

        if (!from || !to) {
            alert('Укажите период');
            return;
        }

        console.log('📊 Построение дашборда:', { from, to, filial, counterparty });

        // Загружаем все данные для всех вкладок
        await Promise.all([
            this.loadSummary(from, to),
            this.loadSwipeHistory(from, to),
            this.loadFilialTrendData(from, to, filial),
            this.loadCounterpartyTrend(from, to, filial, counterparty),
            this.loadRawSwipeData(from, to, filial)
        ]);

        // Рендерим графики для активной вкладки
        if (this.currentTab === 'overview') {
            this.renderOverviewCharts();
        } else if (this.currentTab === 'details') {
            this.renderFilialsDetailTable();
        } else if (this.currentTab === 'comparison') {
            this.renderComparisonChart();
        }

        console.log('📊 ReportsManager: buildDashboard() завершён');
    }

    // ===== ЗАГРУЗКА СВОДКИ =====
    async loadSummary(from, to) {
        try {
            const resp = await fetch(`${this.apiBase}/api/summary?from=${from}&to=${to}`);
            const json = await resp.json();

            if (json.success && json.data) {
                const d = json.data;
                document.getElementById('dashSwipeCount').textContent = d.swipe_count || 0;
                document.getElementById('dashMinOverdue').textContent = this.formatCurrency(d.min_overdue || 0);
                document.getElementById('dashMaxOverdue').textContent = this.formatCurrency(d.max_overdue || 0);
                document.getElementById('dashAvgOverdue').textContent = this.formatCurrency(d.avg_overdue || 0);
            } else {
                this.clearSummaryCards();
            }
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки сводки:', e);
            this.clearSummaryCards();
        }
    }

    clearSummaryCards() {
        document.getElementById('dashSwipeCount').textContent = '—';
        document.getElementById('dashMinOverdue').textContent = '—';
        document.getElementById('dashMaxOverdue').textContent = '—';
        document.getElementById('dashAvgOverdue').textContent = '—';
    }

    // ===== ЗАГРУЗКА ИСТОРИИ СВЕРОК =====
    async loadSwipeHistory(from, to) {
        try {
            const resp = await fetch(`${this.apiBase}/api/swipe-dates?from=${from}&to=${to}`);
            const json = await resp.json();

            const tbody = document.getElementById('swipeHistoryBody');
            if (!tbody) return;

            if (json.success && json.data && json.data.length > 0) {
                let html = '';
                [...json.data].reverse().forEach(row => {
                    const dateFormatted = this.formatDateDisplay(row.date);
                    html += `<tr>
                        <td>${dateFormatted}</td>
                        <td class="number-cell">${this.formatCurrency(row.total_overdue)}</td>
                        <td class="number-cell">${this.formatCurrency(row.total_debt)}</td>
                        <td>${row.filial_count}</td>
                        <td>${row.counterparty_count}</td>
                    </tr>`;
                });
                tbody.innerHTML = html;
            } else {
                tbody.innerHTML = '<tr class="empty-row"><td colspan="5">Нет данных за указанный период</td></tr>';
            }
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки истории:', e);
        }
    }

    // ===== ЗАГРУЗКА СЫРЫХ ДАННЫХ ДЛЯ ДЕТАЛИЗАЦИИ =====
    async loadRawSwipeData(from, to, filial) {
        try {
            let url = `${this.apiBase}/api/swipe-raw?from=${from}&to=${to}`;
            if (filial) url += `&filial=${encodeURIComponent(filial)}`;

            const resp = await fetch(url);
            const json = await resp.json();

            if (json.success && json.data) {
                this.rawSwipeData = json.data;
                console.log(`📊 Загружено ${this.rawSwipeData.length} записей сырых данных`);
            } else {
                this.rawSwipeData = [];
            }
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки сырых данных:', e);
            this.rawSwipeData = [];
        }
    }

    // ===== ЗАГРУЗКА ДАННЫХ ДЛЯ ГРАФИКА ФИЛИАЛОВ =====
    async loadFilialTrendData(from, to, filial) {
        try {
            let url = `${this.apiBase}/api/filial-trend?from=${from}&to=${to}`;
            if (filial) url += `&filial=${encodeURIComponent(filial)}`;

            const resp = await fetch(url);
            const json = await resp.json();

            if (!json.success || !json.data || json.data.dates.length === 0) {
                this.destroyChart('overdueTrend');
                return;
            }

            this.filialTrendData = json.data;
            this.allFilialData = json.data.series.map(s => ({
                name: s.name,
                latestValue: s.data[s.data.length - 1] || 0,
                data: s.data
            }));

            if (this.currentTab === 'overview') {
                this.renderOverviewCharts();
            }
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки тренда филиалов:', e);
        }
    }

    // ===== ОТРИСОВКА ГРАФИКОВ ДЛЯ ВКЛАДКИ "ОБЗОР" =====
    renderOverviewCharts() {
        if (!this.filialTrendData) return;

        const { dates, series } = this.filialTrendData;
        const labels = dates.map(d => this.formatDateDisplay(d));

        // ТОП-5 филиалов по последней дате
        const sortedSeries = [...series].sort((a, b) => {
            const lastA = a.data[a.data.length - 1] || 0;
            const lastB = b.data[b.data.length - 1] || 0;
            return lastB - lastA;
        }).slice(0, 5);

        this.renderLineChart('overdueTrend', 'overdueTrendChart', labels, sortedSeries, '₽');
    }

    // ===== ГРАФИК ПО КОНТРАГЕНТАМ =====
    async loadCounterpartyTrend(from, to, filial, counterparty) {
        try {
            let url = `${this.apiBase}/api/counterparty-trend?from=${from}&to=${to}`;
            if (filial) url += `&filial=${encodeURIComponent(filial)}`;
            if (counterparty) url += `&counterparty=${encodeURIComponent(counterparty)}`;

            const resp = await fetch(url);
            const json = await resp.json();

            if (!json.success || !json.data || json.data.dates.length === 0) {
                this.destroyChart('counterpartyTrend');
                return;
            }

            const { dates, series } = json.data;
            const labels = dates.map(d => this.formatDateDisplay(d));
            this.renderLineChart('counterpartyTrend', 'counterpartyTrendChart', labels, series, '₽');
        } catch (e) {
            console.error('📊 ReportsManager: ОШИБКА загрузки тренда контрагентов:', e);
        }
    }

    // ===== РАСЧЁТ ВЗВЕШЕННОГО СРЕДНЕГО =====
    calculateWeightedAverage(dataPoints, periodStart, periodEnd) {
        if (!dataPoints || dataPoints.length === 0) return null;

        const start = new Date(periodStart);
        const end = new Date(periodEnd);
        const totalDays = (end - start) / (1000 * 60 * 60 * 24);

        if (totalDays <= 0) return null;

        // Сортируем точки по дате
        const sorted = [...dataPoints].sort((a, b) => new Date(a.date) - new Date(b.date));

        let weightedSum = 0;
        let totalWeight = 0;

        for (let i = 0; i < sorted.length; i++) {
            const current = sorted[i];
            const next = sorted[i + 1];

            const currentDate = new Date(current.date);
            const nextDate = next ? new Date(next.date) : end;

            // Количество дней, в течение которых действовало это значение
            const days = Math.min((nextDate - currentDate) / (1000 * 60 * 60 * 24), totalDays);

            if (days > 0) {
                weightedSum += current.overdue * days;
                totalWeight += days;
            }
        }

        if (totalWeight === 0) return null;

        return weightedSum / totalWeight;
    }

    // ===== ГРУППИРОВКА ДАННЫХ ПО ПОДПЕРИОДАМ =====
    groupDataByPeriod(filial, breakdown, periodStart, periodEnd) {
        if (!this.rawSwipeData) return [];

        // Фильтруем данные по филиалу
        const filialData = this.rawSwipeData.filter(d => d.filial === filial);

        const subPeriods = this.getSubPeriods(breakdown, periodStart, periodEnd);
        const result = [];

        for (const subPeriod of subPeriods) {
            // Фильтруем точки данных для этого подпериода
            const pointsInPeriod = filialData.filter(d => {
                const date = new Date(d.date);
                return date >= subPeriod.start && date <= subPeriod.end;
            });

            const weightedAvg = this.calculateWeightedAverage(pointsInPeriod, subPeriod.start, subPeriod.end);

            result.push({
                label: subPeriod.label,
                value: weightedAvg,
                start: subPeriod.start,
                end: subPeriod.end
            });
        }

        return result;
    }

    // ===== ПОЛУЧЕНИЕ ПОДПЕРИОДОВ =====
    getSubPeriods(breakdown, periodStart, periodEnd) {
        const start = new Date(periodStart);
        const end = new Date(periodEnd);
        const subPeriods = [];

        if (breakdown === 'day') {
            // По дням
            for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
                const dayStart = new Date(d);
                const dayEnd = new Date(d);
                dayEnd.setHours(23, 59, 59, 999);
                subPeriods.push({
                    label: this.formatDateDisplay(d.toISOString().split('T')[0]),
                    start: dayStart,
                    end: dayEnd
                });
            }
        } else if (breakdown === 'decade') {
            // По декадам (1-10, 11-20, 21-31)
            const year = start.getFullYear();
            const month = start.getMonth();
            
            for (let decade = 1; decade <= 3; decade++) {
                const decadeStart = new Date(year, month, (decade - 1) * 10 + 1);
                const decadeEnd = new Date(year, month, decade * 10);
                
                // Последняя декада может быть до 31
                if (decade === 3) {
                    decadeEnd.setDate(new Date(year, month + 1, 0).getDate());
                }

                // Проверяем, попадает ли декада в выбранный период
                if (decadeEnd >= start && decadeStart <= end) {
                    const actualStart = decadeStart < start ? start : decadeStart;
                    const actualEnd = decadeEnd > end ? end : decadeEnd;
                    
                    subPeriods.push({
                        label: `${decade}-я декада`,
                        start: actualStart,
                        end: actualEnd
                    });
                }
            }
        } else if (breakdown === 'month') {
            // По месяцам
            const monthNames = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
            
            for (let m = start.getMonth(); m <= end.getMonth() || start.getFullYear() < end.getFullYear(); m++) {
                const year = start.getFullYear() + Math.floor(m / 12);
                const actualMonth = m % 12;
                
                const monthStart = new Date(year, actualMonth, 1);
                const monthEnd = new Date(year, actualMonth + 1, 0);

                if (monthEnd >= start && monthStart <= end) {
                    const actualStart = monthStart < start ? start : monthStart;
                    const actualEnd = monthEnd > end ? end : monthEnd;
                    
                    subPeriods.push({
                        label: monthNames[actualMonth],
                        start: actualStart,
                        end: actualEnd
                    });
                }

                if (actualMonth === 11 && year >= end.getFullYear()) break;
            }
        } else if (breakdown === 'quarter') {
            // По кварталам
            const quarterNames = ['1-й кв', '2-й кв', '3-й кв', '4-й кв'];
            
            for (let q = 0; q < 4; q++) {
                const quarterStart = new Date(start.getFullYear(), q * 3, 1);
                const quarterEnd = new Date(start.getFullYear(), q * 3 + 3, 0);

                if (quarterEnd >= start && quarterStart <= end) {
                    const actualStart = quarterStart < start ? start : quarterStart;
                    const actualEnd = quarterEnd > end ? end : quarterEnd;
                    
                    subPeriods.push({
                        label: quarterNames[q],
                        start: actualStart,
                        end: actualEnd
                    });
                }
            }
        }

        return subPeriods;
    }

    // ===== ПОЛУЧЕНИЕ ЗАГОЛОВКА ТАБЛИЦЫ =====
    getTableTitle() {
        const from = document.getElementById('dashFromDate')?.value;
        const to = document.getElementById('dashToDate')?.value;

        if (!from || !to) return 'Все филиалы';

        const fromDate = new Date(from);
        const toDate = new Date(to);

        const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];

        let periodLabel = '';
        if (this.detailMainPeriod === 'decade') {
            periodLabel = `${monthNames[fromDate.getMonth()]} ${fromDate.getFullYear()} (по дням)`;
        } else if (this.detailMainPeriod === 'month') {
            periodLabel = `${monthNames[fromDate.getMonth()]} ${fromDate.getFullYear()} (по ${this.getBreakdownLabel(this.detailBreakdown)})`;
        } else if (this.detailMainPeriod === 'quarter') {
            const quarter = Math.floor(fromDate.getMonth() / 3) + 1;
            periodLabel = `${quarter} квартал ${fromDate.getFullYear()} (по ${this.getBreakdownLabel(this.detailBreakdown)})`;
        } else if (this.detailMainPeriod === 'year') {
            periodLabel = `${fromDate.getFullYear()} год (по ${this.getBreakdownLabel(this.detailBreakdown)})`;
        }

        return periodLabel || 'Все филиалы';
    }

    // ===== ОТРИСОВКА ТАБЛИЦЫ ДЕТАЛИЗАЦИИ =====
    renderFilialsDetailTable() {
        const tbody = document.getElementById('filialsDetailBody');
        const thead = document.getElementById('filialsDetailHeader');
        const subtitle = document.getElementById('detailTableSubtitle');

        if (!tbody || !thead) return;

        // Подпись под таблицей
        if (subtitle) {
            subtitle.textContent = 'Показано взвешенное среднее ПДЗ за каждый подпериод';
        }

        if (!this.rawSwipeData || this.rawSwipeData.length === 0) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="5">Нет данных</td></tr>';
            return;
        }

        // Получаем уникальные филиалы
        const filials = [...new Set(this.rawSwipeData.map(d => d.filial))];
        
        // Получаем подпериоды
        const from = document.getElementById('dashFromDate')?.value;
        const to = document.getElementById('dashToDate')?.value;
        const subPeriods = this.getSubPeriods(this.detailBreakdown, from, to);

        if (subPeriods.length === 0) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="5">Не удалось определить подпериоды</td></tr>';
            return;
        }

        // Формируем заголовок таблицы
        let headerHtml = `<th style="width: 50px; text-align: center;">#</th>`;
        headerHtml += `<th>Филиал</th>`;
        
        subPeriods.forEach((sp, idx) => {
            headerHtml += `<th class="dynamic-col" title="${this.formatDateDisplay(sp.start.toISOString().split('T')[0])} — ${this.formatDateDisplay(sp.end.toISOString().split('T')[0])}">${sp.label}</th>`;
        });

        headerHtml += `<th class="col-average">Среднее</th>`;
        headerHtml += `<th class="col-trend">Итого Δ%</th>`;
        headerHtml += `<th class="dynamic-col" style="text-align: center;">Тренд</th>`;

        thead.innerHTML = headerHtml;

        // Формируем строки таблицы
        let html = '';
        filials.forEach((filial, filialIdx) => {
            const periodData = this.groupDataByPeriod(filial, this.detailBreakdown, from, to);
            
            // Считаем общее среднее за период
            const allValues = periodData.filter(d => d.value !== null).map(d => d.value);
            const overallAvg = allValues.length > 0 ? allValues.reduce((a, b) => a + b, 0) / allValues.length : null;

            // Считаем тренд (первое vs последнее значение)
            const firstValue = periodData.find(d => d.value !== null)?.value || 0;
            const lastValue = [...periodData].reverse().find(d => d.value !== null)?.value || 0;
            const change = firstValue > 0 ? ((lastValue - firstValue) / firstValue * 100) : 0;
            const trend = change > 5 ? 'up' : (change < -5 ? 'down' : 'stable');
            const changeClass = trend === 'up' ? 'trend-up' : (trend === 'down' ? 'trend-down' : 'trend-stable');
            const trendIcon = trend === 'up' ? 'Рост' : (trend === 'down' ? 'Падение' : '–');
            const changeText = change >= 0 ? `+${change.toFixed(1)}%` : `${change.toFixed(1)}%`;

            html += `<tr>`;
            html += `<td style="text-align: center;">${filialIdx + 1}</td>`;
            html += `<td><strong>${filial}</strong></td>`;

            // Значения по подпериодам
            periodData.forEach(pd => {
                if (pd.value !== null) {
                    html += `<td class="number-cell dynamic-col">${this.formatNumberRU(pd.value)}</td>`;
                } else {
                    html += `<td class="number-cell dynamic-col">—</td>`;
                }
            });

            // Среднее
            if (overallAvg !== null) {
                html += `<td class="number-cell col-average"><strong>${this.formatNumberRU(overallAvg)}</strong></td>`;
            } else {
                html += `<td class="number-cell col-average">—</td>`;
            }

            // Изменение
            html += `<td class="number-cell col-trend ${changeClass}">${changeText}</td>`;
            
            // Тренд иконка
            html += `<td class="${changeClass}" style="text-align: center;"><span class="trend-icon">${trendIcon}</span></td>`;
            
            html += `</tr>`;
        });

        tbody.innerHTML = html;
    }

    // ===== ФОРМАТИРОВАНИЕ ЧИСЕЛ (РУССКИЙ ФОРМАТ) =====
    formatNumberRU(value) {
        if (value === null || value === undefined) return '—';
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(value);
    }

    formatCurrency(amount) {
        if (!amount && amount !== 0) return '—';
        return new Intl.NumberFormat('ru-RU', {
            style: 'currency',
            currency: 'RUB',
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
        }).format(amount);
    }

    formatDateDisplay(isoDate) {
        if (!isoDate) return '';
        const parts = isoDate.split('-');
        if (parts.length === 3) {
            return `${parts[2]}.${parts[1]}.${parts[0]}`;
        }
        return isoDate;
    }

    // ===== ЭКСПОРТ ТАБЛИЦЫ ДЕТАЛИЗАЦИИ В EXCEL (ExcelJS — красивое форматирование) =====
    async exportFilialsTable() {
        if (!this.rawSwipeData || this.rawSwipeData.length === 0) {
            alert('Нет данных для экспорта');
            return;
        }

        const from = document.getElementById('dashFromDate')?.value;
        const to = document.getElementById('dashToDate')?.value;
        const allSubPeriods = this.getSubPeriods(this.detailBreakdown, from, to);
        const filials = [...new Set(this.rawSwipeData.map(d => d.filial))].sort();

        if (allSubPeriods.length === 0 || filials.length === 0) {
            alert('Недостаточно данных для экспорта');
            return;
        }

        // ===== 1. ВЫЧИСЛЯЕМ ВСЕ ДАННЫЕ ЗАРАНЕЕ =====
        const filialRows = filials.map((filial, idx) => {
            const periodData = this.groupDataByPeriod(filial, this.detailBreakdown, from, to);
            const allValues = periodData.filter(d => d.value !== null).map(d => d.value);
            const overallAvg = allValues.length > 0 ? allValues.reduce((a, b) => a + b, 0) / allValues.length : null;
            const firstValue = periodData.find(d => d.value !== null)?.value || 0;
            const lastValue = [...periodData].reverse().find(d => d.value !== null)?.value || 0;
            const change = firstValue > 0 ? ((lastValue - firstValue) / firstValue * 100) : 0;
            const trendIcon = change > 5 ? 'Рост' : (change < -5 ? 'Падение' : '–');
            return { idx, filial, periodData, overallAvg, change, trendIcon };
        });

        // ===== 2. ФИЛЬТРУЕМ ПУСТЫЕ ПОДПЕРИОДЫ =====
        // Оставляем только те подпериоды, где хотя бы у одного филиала есть не-null значение
        const nonEmptyIndices = [];
        allSubPeriods.forEach((sp, spIdx) => {
            const hasData = filialRows.some(fr => fr.periodData[spIdx]?.value !== null);
            if (hasData) nonEmptyIndices.push(spIdx);
        });

        if (nonEmptyIndices.length === 0) {
            alert('Нет данных для экспорта (все подпериоды пустые)');
            return;
        }

        const subPeriods = nonEmptyIndices.map(i => allSubPeriods[i]);
        // Отображаем periodData каждого филиала — только непустые индексы
        const filialPeriodDataFiltered = filialRows.map(fr => ({
            ...fr,
            periodData: nonEmptyIndices.map(i => fr.periodData[i])
        }));

        // ===== 3. СОЗДАЁМ WORKBOOK ЧЕРЕЗ ExcelJS =====
        const ExcelJS = window.ExcelJS;
        if (!ExcelJS) {
            alert('Библиотека ExcelJS не загружена. Проверьте подключение CDN.');
            return;
        }

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Детализация ПДЗ', {
            views: [{ state: 'frozen', xSplit: 2, ySplit: 2 }]
        });

        const numDataCols = subPeriods.length;
        const totalCols = 2 + numDataCols + 3; // №, Филиал, ..., Среднее, Δ%, Тренд

        // === ЦВЕТОВАЯ ПАЛИТРА ===
        const BLUE_BG = '2563EB';
        const WHITE_FONT = 'FFFFFF';
        const TOTAL_BG = 'F1F5F9';
        const BORDER_COLOR = 'CBD5E1';
        const THIN_BORDER = { style: 'thin', color: { argb: BORDER_COLOR } };
        const ALL_BORDERS = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
        const ALL_BORDERS_TOTAL = {
            top: { style: 'medium', color: { argb: '475569' } },
            bottom: { style: 'medium', color: { argb: '475569' } },
            left: THIN_BORDER, right: THIN_BORDER
        };

        // === СТИЛИ ===
        const titleStyle = { bold: true, size: 10, name: 'Arial' };
        const headerStyle = { bold: true, size: 7, color: { argb: WHITE_FONT }, name: 'Arial' };
        const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE_BG } };
        const headerAlign = { horizontal: 'center', vertical: 'middle', wrapText: true };
        const dataAlignRight = { horizontal: 'right', vertical: 'middle' };
        const dataAlignCenter = { horizontal: 'center', vertical: 'middle' };
        const dataAlignLeft = { horizontal: 'left', vertical: 'middle' };
        const dataFont = { size: 7, name: 'Arial' };
        const totalStyle = { bold: true, size: 7, name: 'Arial' };
        const totalFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOTAL_BG } };
        const numberFormat = '#,##0.00';

        // ---- СТРОКА 1: ЗАГОЛОВОК ПЕРИОДА ----
        const titleRow = ws.getRow(1);
        ws.mergeCells(1, 1, 1, totalCols);
        const titleCell = titleRow.getCell(1);
        titleCell.value = this.getTableTitle();
        titleCell.font = titleStyle;
        titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
        titleRow.height = 28;

        // ---- СТРОКА 2: ЗАГОЛОВКИ ТАБЛИЦЫ ----
        const headerRow = ws.getRow(2);
        headerRow.height = 32;
        const headers = ['№', 'Филиал'];
        subPeriods.forEach(sp => headers.push(sp.label));
        headers.push('Среднее');
        headers.push('Итого Δ%');
        headers.push('Тренд');

        headers.forEach((h, ci) => {
            const cell = headerRow.getCell(ci + 1);
            cell.value = h;
            cell.font = headerStyle;
            cell.fill = headerFill;
            cell.alignment = headerAlign;
            cell.border = ALL_BORDERS;
        });

        // ---- СТРОКИ ДАННЫХ ----
        filialPeriodDataFiltered.forEach((fr, rowIdx) => {
            const rowNum = 3 + rowIdx;
            const row = ws.getRow(rowNum);

            // №
            const noCell = row.getCell(1);
            noCell.value = fr.idx + 1;
            noCell.font = dataFont;
            noCell.alignment = dataAlignCenter;
            noCell.border = ALL_BORDERS;

            // Филиал
            const nameCell = row.getCell(2);
            nameCell.value = fr.filial;
            nameCell.font = { size: 7, name: 'Arial', bold: true };
            nameCell.alignment = dataAlignLeft;
            nameCell.border = ALL_BORDERS;

            // Подпериоды
            fr.periodData.forEach((pd, ci) => {
                const cell = row.getCell(3 + ci);
                cell.value = pd.value !== null ? Math.round(pd.value * 100) / 100 : 0;
                cell.numFmt = numberFormat;
                cell.font = dataFont;
                cell.alignment = dataAlignRight;
                cell.border = ALL_BORDERS;
            });

            // Среднее
            const avgCol = 3 + numDataCols;
            const avgCell = row.getCell(avgCol);
            avgCell.value = fr.overallAvg !== null ? Math.round(fr.overallAvg * 100) / 100 : 0;
            avgCell.numFmt = numberFormat;
            avgCell.font = { size: 7, name: 'Arial', bold: true };
            avgCell.alignment = dataAlignRight;
            avgCell.border = ALL_BORDERS;

            // Δ%
            const deltaCol = avgCol + 1;
            const deltaCell = row.getCell(deltaCol);
            deltaCell.value = `${fr.change >= 0 ? '+' : ''}${fr.change.toFixed(1)}%`;
            deltaCell.font = { size: 7, name: 'Arial', bold: true };
            if (fr.change > 5) {
                deltaCell.font = { size: 7, name: 'Arial', bold: true, color: { argb: 'DC2626' } };
            } else if (fr.change < -5) {
                deltaCell.font = { size: 7, name: 'Arial', bold: true, color: { argb: '059669' } };
            }
            deltaCell.alignment = dataAlignRight;
            deltaCell.border = ALL_BORDERS;

            // Тренд
            const trendCol = deltaCol + 1;
            const trendCell = row.getCell(trendCol);
            trendCell.value = fr.trendIcon;
            if (fr.trendIcon === 'Рост') {
                trendCell.font = { size: 7, name: 'Arial', bold: true, color: { argb: 'DC2626' } };
            } else if (fr.trendIcon === 'Падение') {
                trendCell.font = { size: 7, name: 'Arial', bold: true, color: { argb: '059669' } };
            } else {
                trendCell.font = { size: 7, name: 'Arial', color: { argb: '94A3B8' } };
            }
            trendCell.alignment = dataAlignCenter;
            trendCell.border = ALL_BORDERS;
        });

        // ---- СТРОКА ИТОГО ----
        const totalRowNum = 3 + filialPeriodDataFiltered.length;
        const totalRow = ws.getRow(totalRowNum);
        totalRow.height = 26;

        // № и ИТОГО
        const totalLabelCell = totalRow.getCell(2);
        totalLabelCell.value = 'ИТОГО';
        totalLabelCell.font = totalStyle;
        totalLabelCell.alignment = dataAlignLeft;
        totalLabelCell.border = ALL_BORDERS_TOTAL;
        totalLabelCell.fill = totalFill;

        // Граница для № в ИТОГО
        totalRow.getCell(1).border = ALL_BORDERS_TOTAL;
        totalRow.getCell(1).fill = totalFill;

        // Суммы по подпериодам
        for (let ci = 0; ci < numDataCols; ci++) {
            let sum = 0;
            let count = 0;
            filialPeriodDataFiltered.forEach(fr => {
                const val = fr.periodData[ci]?.value;
                if (val !== null && val !== undefined) { sum += val; count++; }
            });
            const cell = totalRow.getCell(3 + ci);
            cell.value = Math.round(sum * 100) / 100;
            cell.numFmt = numberFormat;
            cell.font = totalStyle;
            cell.alignment = dataAlignRight;
            cell.border = ALL_BORDERS_TOTAL;
            cell.fill = totalFill;
        }

        // Среднее в ИТОГО
        const avgOfAvgs = filialPeriodDataFiltered
            .map(fr => fr.overallAvg)
            .filter(v => v !== null);
        const totalAvgCell = totalRow.getCell(3 + numDataCols);
        const ta = avgOfAvgs.length > 0
            ? avgOfAvgs.reduce((a, b) => a + b, 0) / avgOfAvgs.length
            : 0;
        totalAvgCell.value = Math.round(ta * 100) / 100;
        totalAvgCell.numFmt = numberFormat;
        totalAvgCell.font = totalStyle;
        totalAvgCell.alignment = dataAlignRight;
        totalAvgCell.border = ALL_BORDERS_TOTAL;
        totalAvgCell.fill = totalFill;

        // Δ% и Тренд в ИТОГО
        for (let ci = numDataCols + 1; ci <= numDataCols + 2; ci++) {
            const cell = totalRow.getCell(3 + ci);
            cell.border = ALL_BORDERS_TOTAL;
            cell.fill = totalFill;
        }

        // ===== 4. ЗАКРЕПЛЕНИЕ ПАНЕЛИ (заголовки + 2 колонки) =====
        ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];

        // ===== 5. АВТОФИЛЬТР =====
        ws.autoFilter = {
            from: { row: 2, column: 1 },
            to: { row: totalRowNum, column: totalCols }
        };

        // ===== 6. ШИРИНА КОЛОНОК =====
        ws.getColumn(1).width = 5;    // №
        ws.getColumn(2).width = 35;   // Филиал
        for (let ci = 0; ci < numDataCols; ci++) {
            ws.getColumn(3 + ci).width = 16;  // Подпериоды
        }
        ws.getColumn(3 + numDataCols).width = 16;     // Среднее
        ws.getColumn(4 + numDataCols).width = 14;     // Δ%
        ws.getColumn(5 + numDataCols).width = 11;     // Тренд

        // ===== 7. СОХРАНЕНИЕ =====
        const buffer = await wb.xlsx.writeBuffer();
        const safeTitle = this.getTableTitle()
            .replace(/[^а-яА-ЯёЁa-zA-Z0-9\s-]/g, '')
            .replace(/\s+/g, '_');
        const filename = `Детализация_ПДЗ_${safeTitle}_${new Date().toISOString().split('T')[0]}.xlsx`;

        // FileSaver
        if (typeof saveAs !== 'undefined') {
            saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
        } else {
            const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
    }

    // ===== ВКЛАДКА "СРАВНЕНИЕ": ЧЕКБОКСЫ ФИЛИАЛОВ =====
    renderFilialCheckboxes() {
        const container = document.getElementById('filialCheckboxes');
        if (!container || !this.allFilialData) return;

        // Сортируем по убыванию ПДЗ
        const sorted = [...this.allFilialData].sort((a, b) => b.latestValue - a.latestValue);

        let html = '';
        sorted.forEach(filial => {
            const isChecked = this.selectedFilials.has(filial.name);
            html += `
                <label class="filial-checkbox">
                    <input type="checkbox" value="${filial.name}" ${isChecked ? 'checked' : ''} 
                           data-name="${filial.name}" data-value="${filial.latestValue}">
                    <span title="${filial.name}">${filial.name}</span>
                    <span class="filial-amount">${this.formatCurrency(filial.latestValue)}</span>
                </label>
            `;
        });

        container.innerHTML = html;

        // Добавляем обработчики чекбоксов
        container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.addEventListener('change', (e) => {
                const name = e.target.dataset.name;
                if (e.target.checked) {
                    this.selectedFilials.add(name);
                } else {
                    this.selectedFilials.delete(name);
                }
                this.renderComparisonChart();
            });
        });
    }

    selectTop5Filials() {
        if (!this.allFilialData) return;
        
        const top5 = [...this.allFilialData]
            .sort((a, b) => b.latestValue - a.latestValue)
            .slice(0, 5)
            .map(f => f.name);

        this.selectedFilials = new Set(top5);
        this.renderFilialCheckboxes();
        this.renderComparisonChart();
    }

    selectAllFilials() {
        if (!this.allFilialData) return;
        this.selectedFilials = new Set(this.allFilialData.map(f => f.name));
        this.renderFilialCheckboxes();
        this.renderComparisonChart();
    }

    clearFilialSelection() {
        this.selectedFilials.clear();
        this.renderFilialCheckboxes();
        this.renderComparisonChart();
    }

    // ===== ГРАФИК СРАВНЕНИЯ ВЫБРАННЫХ ФИЛИАЛОВ =====
    renderComparisonChart() {
        if (!this.filialTrendData || this.selectedFilials.size === 0) {
            this.destroyChart('comparison');
            return;
        }

        const { dates, series } = this.filialTrendData;
        const labels = dates.map(d => this.formatDateDisplay(d));

        // Фильтруем серии по выбранным филиалам
        const filteredSeries = series.filter(s => this.selectedFilials.has(s.name));

        if (filteredSeries.length === 0) {
            this.destroyChart('comparison');
            return;
        }

        this.renderLineChart('comparison', 'comparisonChart', labels, filteredSeries, '₽');
    }

    // ===== ОТРИСОВКА ГРАФИКОВ =====
    renderLineChart(chartId, canvasId, labels, series, unit) {
        console.log('📊 ReportsManager: renderLineChart() вызван, canvasId:', canvasId);
        this.destroyChart(chartId);

        const canvas = document.getElementById(canvasId);
        if (!canvas) return;
        
        const ctx = canvas.getContext('2d');
        const colors = this.getChartColors(series.length);

        const datasets = series.map((s, i) => ({
            label: s.name,
            data: s.data,
            borderColor: colors[i],
            backgroundColor: colors[i] + '20',
            tension: 0.3,
            fill: series.length === 1,
            pointRadius: 4,
            pointHoverRadius: 6,
            borderWidth: 2
        }));

        this.charts[chartId] = new Chart(ctx, {
            type: 'line',
            data: { labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                aspectRatio: 2.5,
                plugins: {
                    legend: {
                        display: series.length > 1,
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            padding: 16,
                            font: { size: 12 }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                return ctx.dataset.label + ': ' +
                                    new Intl.NumberFormat('ru-RU', {
                                        style: 'currency',
                                        currency: 'RUB',
                                        minimumFractionDigits: 0,
                                        maximumFractionDigits: 0
                                    }).format(ctx.parsed.y);
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        type: 'logarithmic',
                        min: 100000,
                        max: 1000000000,
                        ticks: {
                            callback: function(value) {
                                const steps = [100000, 500000, 1000000, 10000000, 50000000, 100000000, 250000000, 500000000, 1000000000];
                                if (steps.includes(value)) {
                                    if (value >= 1e9) return (value / 1e9).toFixed(1).replace('.0', '') + ' млрд';
                                    if (value >= 1e6) return (value / 1e6).toFixed(1).replace('.0', '') + ' млн';
                                    if (value >= 1e3) return (value / 1e3).toFixed(1).replace('.0', '') + ' тыс';
                                    return value.toString();
                                }
                                return '';
                            },
                            font: { size: 12, weight: '500' },
                            color: '#475569',
                            padding: 8
                        },
                        grid: { color: '#e2e8f0', lineWidth: 1 },
                        border: { display: true, color: '#cbd5e1' }
                    },
                    x: {
                        ticks: { font: { size: 11 }, maxRotation: 45, color: '#64748b' },
                        grid: { display: false }
                    }
                }
            }
        });
    }

    destroyChart(chartId) {
        if (this.charts[chartId]) {
            this.charts[chartId].destroy();
            delete this.charts[chartId];
        }
    }

    // ===== УТИЛИТЫ =====
    getChartColors(count) {
        const palette = [
            '#2563eb', '#059669', '#d97706', '#dc2626', '#7c3aed',
            '#0891b2', '#be185d', '#65a30d', '#ea580c', '#4f46e5',
            '#0d9488', '#c026d3', '#16a34a', '#9333ea', '#0284c7'
        ];
        const result = [];
        for (let i = 0; i < count; i++) {
            result.push(palette[i % palette.length]);
        }
        return result;
    }
}