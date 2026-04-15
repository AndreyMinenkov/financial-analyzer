// reports-manager.js — Управление страницей отчётов и дашбордов
class ReportsManager {
    constructor() {
        this.charts = {};
        this.apiBase = 'http://31.130.155.16:5000';
        this.currentSection = null; // 'debt' или null
    }

    // ===== ИНИЦИАЛИЗАЦИЯ =====
    init() {
        this.setupListeners();
        this.setDefaultDates();
    }

    setupListeners() {
        // Открытие раздела "Дебиторка"
        const card = document.getElementById('reportCardDebt');
        if (card) {
            card.addEventListener('click', () => {
                this.openDashboard('debt');
            });
        }

        // Кнопка "Назад"
        const backBtn = document.getElementById('backToReportsBtn');
        if (backBtn) {
            backBtn.addEventListener('click', () => {
                this.closeDashboard();
            });
        }

        // Кнопка "Сформировать"
        const buildBtn = document.getElementById('buildDashboardBtn');
        if (buildBtn) {
            buildBtn.addEventListener('click', () => {
                this.buildDashboard();
            });
        }

        // При смене филиала — обновить список контрагентов
        const filialSelect = document.getElementById('dashFilialSelect');
        if (filialSelect) {
            filialSelect.addEventListener('change', () => {
                this.loadCounterpartyList();
            });
        }
    }

    setDefaultDates() {
        const today = new Date();
        const weekAgo = new Date(today);
        weekAgo.setDate(today.getDate() - 30);

        document.getElementById('dashToDate').value = this.formatDateISO(today);
        document.getElementById('dashFromDate').value = this.formatDateISO(weekAgo);
    }

    formatDateISO(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    // ===== НАВИГАЦИЯ =====
    openDashboard(section) {
        this.currentSection = section;
        document.getElementById('reportCardDebt').style.display = 'none';
        document.getElementById('dashboardPanel').style.display = 'block';
        this.loadFilialList();
    }

    closeDashboard() {
        this.currentSection = null;
        document.getElementById('reportCardDebt').style.display = 'flex';
        document.getElementById('dashboardPanel').style.display = 'none';
    }

    // ===== ЗАГРУЗКА СПИСКОВ ФИЛЬТРОВ =====
    async loadFilialList() {
        try {
            const from = document.getElementById('dashFromDate').value;
            const to = document.getElementById('dashToDate').value;
            const url = `${this.apiBase}/api/filial-list?from=${from}&to=${to}`;

            const resp = await fetch(url);
            const json = await resp.json();

            const select = document.getElementById('dashFilialSelect');
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
            console.error('Ошибка загрузки списка филиалов:', e);
        }
    }

    async loadCounterpartyList() {
        try {
            const filial = document.getElementById('dashFilialSelect').value;
            let url = `${this.apiBase}/api/counterparty-list`;
            if (filial) url += `?filial=${encodeURIComponent(filial)}`;

            const resp = await fetch(url);
            const json = await resp.json();

            const select = document.getElementById('dashCounterpartySelect');
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
            console.error('Ошибка загрузки списка контрагентов:', e);
        }
    }

    // ===== ПОСТРОЕНИЕ ДАШБОРДА =====
    async buildDashboard() {
        const from = document.getElementById('dashFromDate').value;
        const to = document.getElementById('dashToDate').value;
        const filial = document.getElementById('dashFilialSelect').value;
        const counterparty = document.getElementById('dashCounterpartySelect').value;

        if (!from || !to) {
            alert('Укажите период');
            return;
        }

        console.log('📊 Построение дашборда:', { from, to, filial, counterparty });

        // Параллельная загрузка всех данных
        await Promise.all([
            this.loadSummary(from, to),
            this.loadSwipeHistory(from, to),
            this.loadFilialTrend(from, to, filial),
            this.loadCounterpartyTrend(from, to, filial, counterparty)
        ]);
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
            console.error('Ошибка загрузки сводки:', e);
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

            if (json.success && json.data && json.data.length > 0) {
                let html = '';
                // Показываем в обратном порядке (новые сверху)
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
            console.error('Ошибка загрузки истории:', e);
        }
    }

    // ===== ГРАФИК ДИНАМИКИ ОБЩЕЙ ПДЗ =====
    async loadFilialTrend(from, to, filial) {
        try {
            let url = `${this.apiBase}/api/filial-trend?from=${from}&to=${to}`;
            if (filial) url += `&filial=${encodeURIComponent(filial)}`;

            const resp = await fetch(url);
            const json = await resp.json();

            if (!json.success || !json.data || json.data.dates.length === 0) {
                this.destroyChart('overdueTrend');
                return;
            }

            const { dates, series } = json.data;
            const labels = dates.map(d => this.formatDateDisplay(d));

            this.renderLineChart('overdueTrend', 'overdueTrendChart', labels, series, '₽');
        } catch (e) {
            console.error('Ошибка загрузки тренда филиалов:', e);
        }
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
            console.error('Ошибка загрузки тренда контрагентов:', e);
        }
    }

    // ===== ГРАФИК СРАВНЕНИЯ ФИЛИАЛОВ (последняя дата) =====
    async renderFilialComparison(from, to) {
        try {
            const url = `${this.apiBase}/api/filial-trend?from=${from}&to=${to}`;
            const resp = await fetch(url);
            const json = await resp.json();

            if (!json.success || !json.data || json.data.dates.length === 0) {
                this.destroyChart('filialComparison');
                return;
            }

            // Берём данные за последнюю доступную дату
            const lastDateIndex = json.data.dates.length - 1;
            const labels = json.data.series.map(s => s.name);
            const data = json.data.series.map(s => s.data[lastDateIndex] || 0);

            this.renderBarChart('filialComparison', 'filialComparisonChart', labels, data);
        } catch (e) {
            console.error('Ошибка графика сравнения:', e);
        }
    }

    // ===== ОТРИСОВКА ГРАФИКОВ =====
    renderLineChart(chartId, canvasId, labels, series, unit) {
        this.destroyChart(chartId);

        const ctx = document.getElementById(canvasId).getContext('2d');
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
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                if (value >= 1e9) return (value / 1e9).toFixed(1) + ' млрд';
                                if (value >= 1e6) return (value / 1e6).toFixed(1) + ' млн';
                                if (value >= 1e3) return (value / 1e3).toFixed(0) + ' тыс';
                                return value;
                            },
                            font: { size: 11 }
                        },
                        grid: { color: '#e2e8f0' }
                    },
                    x: {
                        ticks: { font: { size: 11 }, maxRotation: 45 },
                        grid: { display: false }
                    }
                }
            }
        });
    }

    renderBarChart(chartId, canvasId, labels, data) {
        this.destroyChart(chartId);

        const ctx = document.getElementById(canvasId).getContext('2d');
        const colors = this.getChartColors(data.length);

        this.charts[chartId] = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'ПДЗ',
                    data,
                    backgroundColor: colors.map(c => c + 'CC'),
                    borderColor: colors,
                    borderWidth: 1,
                    borderRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                aspectRatio: 2,
                indexAxis: 'y', // Горизонтальная диаграмма
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                return new Intl.NumberFormat('ru-RU', {
                                    style: 'currency',
                                    currency: 'RUB',
                                    minimumFractionDigits: 0
                                }).format(ctx.parsed.x);
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                if (value >= 1e9) return (value / 1e9).toFixed(1) + ' млрд';
                                if (value >= 1e6) return (value / 1e6).toFixed(1) + ' млн';
                                return value;
                            }
                        },
                        grid: { color: '#e2e8f0' }
                    },
                    y: {
                        ticks: { font: { size: 11 } },
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
}
