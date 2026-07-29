    // cashflow-allocator.js - Модуль распределения налогов по подразделениям
class CashFlowAllocator {
    constructor() {
        this.sourceData = [];
        this.filials = [];
        this.originalFile = null;       // ✅ Исходный файл для отправки на сервер
        this.originalHeaders = [];      // ✅ Заголовки исходного файла
        this.settings = {
            periodType: 'month',
            periodValue: '',
            dateFrom: '',
            dateTo: '',
            strategy: 'revenue',
            rounding: 'kopeks',
            excludeArticles: ['ВГО', 'Депозит', 'Перечисление']
        };
        this.init();
    }

    init() {
        console.log('Initializing CashFlowAllocator...');
        this.setupEventListeners();
        this.setDefaultPeriod();
    }

    setDefaultPeriod() {
        const today = new Date();
        const currentMonth = today.getMonth() + 1;
        const year = today.getFullYear();
        
        const monthSelect = document.getElementById('selectMonth');
        if (monthSelect) {
            const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 
                               'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
            const shortYear = year.toString().substr(-2);
            monthSelect.value = `${monthNames[currentMonth - 1]}.${shortYear}`;
        }
        
        const dateFrom = document.getElementById('customDateFrom');
        const dateTo = document.getElementById('customDateTo');
        if (dateFrom && dateTo) {
            const firstDay = new Date(year, currentMonth - 1, 1);
            const lastDay = new Date(year, currentMonth, 0);
            dateFrom.value = this.formatDateOutput(firstDay);
            dateTo.value = this.formatDateOutput(lastDay);
        }
    }

    // Исправленная функция - парсит даты в формате DD.MM.YYYY из Excel
    parseDate(dateValue) {
        if (!dateValue) return null;
        
        if (dateValue instanceof Date) {
            return dateValue;
        }
        
        if (typeof dateValue === 'number') {
            return new Date(Math.round((dateValue - 25569) * 86400 * 1000));
        }
        
        if (typeof dateValue === 'string') {
            const str = dateValue.trim();
            
            const ddmmyyyy = str.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
            if (ddmmyyyy) {
                const day = parseInt(ddmmyyyy[1], 10);
                const month = parseInt(ddmmyyyy[2], 10) - 1;
                const year = parseInt(ddmmyyyy[3], 10);
                const date = new Date(year, month, day);
                if (!isNaN(date.getTime())) {
                    return date;
                }
            }
            
            const mmddyyyy = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
            if (mmddyyyy) {
                const month = parseInt(mmddyyyy[1], 10) - 1;
                const day = parseInt(mmddyyyy[2], 10);
                const year = parseInt(mmddyyyy[3], 10);
                const date = new Date(year, month, day);
                if (!isNaN(date.getTime())) {
                    return date;
                }
            }
            
            const isoyyyy = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (isoyyyy) {
                const year = parseInt(isoyyyy[1], 10);
                const month = parseInt(isoyyyy[2], 10) - 1;
                const day = parseInt(isoyyyy[3], 10);
                const date = new Date(year, month, day);
                if (!isNaN(date.getTime())) {
                    return date;
                }
            }
            
            const date = new Date(str);
            if (!isNaN(date.getTime())) {
                return date;
            }
            
            console.warn('Невалидная дата:', dateValue);
            return null;
        }
        
        console.warn('Невалидная дата:', dateValue);
        return null;
    }

    formatDateOutput(date) {
        if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
            return '';
        }
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    setupEventListeners() {
        console.log('Setting up event listeners...');
        
        const selectCashflowBtn = document.getElementById('selectCashflowBtn');
        const cashflowFile = document.getElementById('cashflowFile');
        const cashflowDropArea = document.getElementById('cashflowDropArea');

        if (selectCashflowBtn) {
            selectCashflowBtn.addEventListener('click', () => {
                console.log('Button clicked, triggering file input...');
                cashflowFile.click();
            });
        }

        if (cashflowFile) {
            cashflowFile.addEventListener('change', (e) => {
                console.log('File selected:', e.target.files[0]?.name);
                if (e.target.files.length > 0) {
                    this.loadCashflowFile(e.target.files[0]);
                }
            });
        }

        if (cashflowDropArea) {
            cashflowDropArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                cashflowDropArea.style.borderColor = '#2563eb';
                cashflowDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
            });

            cashflowDropArea.addEventListener('dragleave', () => {
                cashflowDropArea.style.borderColor = '';
                cashflowDropArea.style.background = '';
            });

            cashflowDropArea.addEventListener('drop', (e) => {
                e.preventDefault();
                cashflowDropArea.style.borderColor = '';
                cashflowDropArea.style.background = '';
                if (e.dataTransfer.files.length > 0) {
                    this.loadCashflowFile(e.dataTransfer.files[0]);
                }
            });
        }

        document.querySelectorAll('input[name="periodType"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.togglePeriodControls(e.target.value);
            });
        });

        const applySettingsBtn = document.getElementById('applyCashflowSettingsBtn');
        if (applySettingsBtn) {
            applySettingsBtn.addEventListener('click', () => {
                this.allocateCashflow();
            });
        }

        const exportResultsBtn = document.getElementById('exportResultsBtn');
        if (exportResultsBtn) {
            exportResultsBtn.addEventListener('click', () => {
                this.exportResults();
            });
        }

        const clearCashflowBtn = document.getElementById('clearCashflowBtn');
        if (clearCashflowBtn) {
            clearCashflowBtn.addEventListener('click', () => {
                this.clearData();
            });
        }
    }

    togglePeriodControls(periodType) {
        document.querySelectorAll('.period-control').forEach(el => {
            el.style.display = 'none';
        });

        const controlId = `period${periodType.charAt(0).toUpperCase() + periodType.slice(1)}`;
        const control = document.getElementById(controlId);
        if (control) {
            control.style.display = 'block';
        }
    }

    async loadCashflowFile(file) {
        console.log('Загрузка файла ДДС:', file.name);
        
        try {
            // ✅ Сохраняем исходный файл для отправки на сервер (со всеми колонками)
            this.originalFile = file;
            
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

            if (!workbook.SheetNames.includes('Источник')) {
                throw new Error('Файл не содержит вкладку "Источник"');
            }

            const worksheet = workbook.Sheets['Источник'];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

            if (rows.length < 2) {
                throw new Error('Файл не содержит данных');
            }

            this.sourceData = this.parseSourceData(rows);
            this.filials = [...new Set(this.sourceData.map(r => r.filial).filter(f => f))];

            document.getElementById('cashflowFileInfo').innerHTML = 
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${this.sourceData.length} записей, ${this.filials.length} подразделений)`;
            
            document.getElementById('cashflowSettings').style.display = 'block';

            this.populateMonthSelect();

            console.log(`✅ Загружено ${this.sourceData.length} записей, ${this.filials.length} подразделений`);

        } catch (error) {
            console.error('Ошибка загрузки файла:', error);
            document.getElementById('cashflowFileInfo').innerHTML = 
                `<i class="fas fa-exclamation-circle" style="color: var(--error);"></i> Ошибка: ${error.message}`;
        }
    }

    parseSourceData(rows) {
        const headers = rows[0].map(h => String(h || '').trim());
        const data = [];

        const filialIdx = headers.findIndex(h => h.includes('Подразделение'));
        const articleIdx = headers.findIndex(h => h.includes('Статья') || h.includes('Назначение'));
        const amountIdx = headers.findIndex(h => h.includes('Факт') || h.includes('Сумма'));
        const dateIdx = headers.findIndex(h => h.includes('Дата'));
        const monthIdx = headers.findIndex(h => h.includes('Месяц'));

        console.log('Индексы колонок:', { filialIdx, articleIdx, amountIdx, dateIdx, monthIdx });

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const filial = filialIdx !== -1 ? String(row[filialIdx] || '').trim() : '';
            const article = articleIdx !== -1 ? String(row[articleIdx] || '').trim() : '';
            const amount = amountIdx !== -1 ? this.parseAmount(row[amountIdx]) : 0;
            const date = dateIdx !== -1 ? row[dateIdx] : null;
            
            // ✅ Месяц: если значение пустое или формула — вычисляем из даты
            let month = monthIdx !== -1 ? String(row[monthIdx] || '').trim() : '';
            if (!month || month.startsWith('=') || month.startsWith('=TEXT')) {
                month = this.monthFromDate(date);
            }

            if (!filial && amount === 0) continue;

            data.push({
                filial,
                article,
                amount,
                date: this.parseDate(date),
                month,
                isIncome: amount > 0,
                isExpense: amount < 0
            });
        }

        return data;
    }

    // ✅ Вспомогательный метод: получить месяц в формате "Апрель.26" из даты
    monthFromDate(dateValue) {
        if (!dateValue) return '';
        const d = this.parseDate(dateValue);
        if (!d || isNaN(d.getTime())) return '';
        const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                           'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
        const shortYear = String(d.getFullYear()).slice(-2);
        return monthNames[d.getMonth()] + '.' + shortYear;
    }

    parseAmount(value) {
        if (value === null || value === undefined) return 0;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            const cleaned = value.replace(/\s/g, '').replace(',', '.');
            return parseFloat(cleaned) || 0;
        }
        return 0;
    }

    populateMonthSelect() {
        const monthSelect = document.getElementById('selectMonth');
        if (!monthSelect) return;

        const months = [...new Set(this.sourceData.map(r => r.month).filter(m => m))];
        const currentValue = monthSelect.value;

        monthSelect.innerHTML = '<option value="">Выберите месяц</option>';
        months.forEach(month => {
            const option = document.createElement('option');
            option.value = month;
            option.textContent = month;
            monthSelect.appendChild(option);
        });

        if (months.includes(currentValue)) {
            monthSelect.value = currentValue;
        }
    }

    getSettings() {
        const periodType = document.querySelector('input[name="periodType"]:checked')?.value || 'month';
        const strategy = document.getElementById('allocationStrategy')?.value || 'revenue';
        const rounding = document.querySelector('input[name="rounding"]:checked')?.value || 'kopeks';
        
        let periodValue = '';
        let dateFrom = '';
        let dateTo = '';

        if (periodType === 'month') {
            periodValue = document.getElementById('selectMonth')?.value || '';
        } else if (periodType === 'quarter') {
            periodValue = document.getElementById('selectQuarter')?.value || '';
        } else if (periodType === 'year') {
            periodValue = document.getElementById('selectYearOnly')?.value || '';
        } else if (periodType === 'custom') {
            dateFrom = document.getElementById('customDateFrom')?.value || '';
            dateTo = document.getElementById('customDateTo')?.value || '';
        }

        const excludeArticles = [];
        if (document.getElementById('excludeVGO')?.checked) excludeArticles.push('ВГО');
        if (document.getElementById('excludeDeposits')?.checked) excludeArticles.push('Депозит');
        if (document.getElementById('excludeTransfers')?.checked) excludeArticles.push('Перечисление');

        return {
            periodType,
            periodValue,
            dateFrom,
            dateTo,
            strategy,
            rounding,
            excludeArticles
        };
    }

    async allocateCashflow() {
        console.log('=== НАЧАЛО РАСПРЕДЕЛЕНИЯ НАЛОГОВ ===');
        
        const settings = this.getSettings();
        console.log('Настройки:', settings);

        if (!this.originalFile) {
            alert('Сначала загрузите файл с данными (исходный файл не найден)');
            console.error('❌ originalFile is null — файл не был загружен через loadCashflowFile()');
            return;
        }

        if (this.sourceData.length === 0) {
            alert('Сначала загрузите файл с данными');
            return;
        }

        if (settings.periodType === 'month' && !settings.periodValue) {
            alert('Выберите месяц');
            return;
        }

        if (settings.periodType === 'custom' && (!settings.dateFrom || !settings.dateTo)) {
            alert('Укажите даты периода');
            return;
        }

        const filteredData = this.filterByPeriod(settings);
        console.log(`Отфильтровано ${filteredData.length} записей за период`);

        if (filteredData.length === 0) {
            alert('Нет данных за выбранный период');
            return;
        }

        this.showLoading();

        try {
            const formData = new FormData();
            
            // ✅ ИСПРАВЛЕНИЕ: Отправляем исходный файл (со всеми колонками)
            // allocator.py сам разберёт все колонки и отфильтрует по периоду
            console.log(`📤 Отправка файла: ${this.originalFile.name} (${this.originalFile.size} байт)`);
            formData.append('file', this.originalFile);
            formData.append('settings', JSON.stringify(settings));

            const response = await fetch('/api/allocate-cashflow', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || 'Ошибка сервера');
            }

            const responseBlob = await response.blob();
            const url = window.URL.createObjectURL(responseBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `ДДС_налоги_распределены_${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            // ИСПРАВЛЕНИЕ: Правильное декодирование UTF-8 из Base64
            const summaryHeader = response.headers.get('X-Allocation-Summary');
            if (summaryHeader) {
                try {
                    // FIX: atob() возвращает Latin-1, поэтому конвертируем через TextDecoder
                    const binary = atob(summaryHeader);
                    const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));
                    const jsonStr = new TextDecoder('utf-8').decode(bytes);
                    const summary = JSON.parse(jsonStr);
                    this.showPreview(summary);
                } catch (e) {
                    console.error('Ошибка декодирования заголовка:', e);
                }
            }

            console.log('✅ Распределение завершено, файл скачан');

        } catch (error) {
            console.error('Ошибка распределения:', error);
            alert('Ошибка: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    filterByPeriod(settings) {
        return this.sourceData.filter(row => {
            if (settings.excludeArticles.some(ex => row.article.includes(ex))) {
                return false;
            }

            if (settings.periodType === 'month') {
                return row.month === settings.periodValue;
            } else if (settings.periodType === 'quarter') {
                const quarterMonths = {
                    'К1': ['Январь', 'Февраль', 'Март'],
                    'К2': ['Апрель', 'Май', 'Июнь'],
                    'К3': ['Июль', 'Август', 'Сентябрь'],
                    'К4': ['Октябрь', 'Ноябрь', 'Декабрь']
                };
                const rowMonth = row.month.split('.')[0];
                return quarterMonths[settings.periodValue]?.includes(rowMonth);
            } else if (settings.periodType === 'year') {
                return row.month.includes(settings.periodValue);
            } else if (settings.periodType === 'custom') {
                if (!row.date) return false;
                const from = new Date(settings.dateFrom);
                const to = new Date(settings.dateTo);
                return row.date >= from && row.date <= to;
            }
            return true;
        });
    }

    showPreview(summary) {
        console.log('Превью результатов:', summary);
        
        document.getElementById('cashflowPreview').style.display = 'block';
        document.getElementById('cashflowResults').style.display = 'block';

        document.getElementById('previewFilialsCount').textContent = summary.filials_count || 0;
        document.getElementById('previewTotalRevenue').textContent = this.formatCurrency(summary.total_revenue || 0);

        const sharesContainer = document.getElementById('filialsSharesContainer');
        if (sharesContainer && summary.shares) {
            sharesContainer.innerHTML = '';
            const sortedFilials = Object.entries(summary.shares)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 10);

            sortedFilials.forEach(([filial, share]) => {
                const percent = (share * 100).toFixed(2);
                const barWidth = Math.max(5, percent * 3);
                
                const div = document.createElement('div');
                div.className = 'filial-share-item';
                div.style.cssText = 'display: flex; align-items: center; gap: 10px; margin-bottom: 8px; padding: 8px; background: var(--bg-tertiary); border-radius: var(--radius-md);';
                div.innerHTML = `
                    <div style="flex: 1; min-width: 0;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
                            <span style="font-size: 13px; font-weight: 500;">${this.escapeHtml(filial)}</span>
                            <span style="font-size: 13px; font-weight: 600;">${percent}%</span>
                        </div>
                        <div style="height: 6px; background: var(--border-light); border-radius: 3px; overflow: hidden;">
                            <div style="width: ${barWidth}%; height: 100%; background: var(--primary); border-radius: 3px;"></div>
                        </div>
                    </div>
                `;
                sharesContainer.appendChild(div);
            });
        }

        document.getElementById('exportCashflowBtn').disabled = false;
    }

    formatCurrency(amount) {
        return new Intl.NumberFormat('ru-RU', {
            style: 'currency',
            currency: 'RUB',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount);
    }

    escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    exportResults() {
        console.log('Экспорт результатов...');
        this.allocateCashflow();
    }

    clearData() {
        if (!confirm('Вы уверены, что хотите очистить все данные?')) {
            return;
        }

        this.sourceData = [];
        this.filials = [];
        
        document.getElementById('cashflowFile').value = '';
        document.getElementById('cashflowFileInfo').innerHTML = 
            '<i class="fas fa-info-circle"></i> Файл не выбран';
        document.getElementById('cashflowSettings').style.display = 'none';
        document.getElementById('cashflowPreview').style.display = 'none';
        document.getElementById('cashflowResults').style.display = 'none';
        document.getElementById('exportCashflowBtn').disabled = true;

        console.log('Данные очищены');
    }

    showLoading() {
        const overlay = document.createElement('div');
        overlay.className = 'loading-overlay active';
        overlay.innerHTML = '<div class="loading-spinner"></div>';
        document.body.appendChild(overlay);
    }

    hideLoading() {
        const overlay = document.querySelector('.loading-overlay');
        if (overlay) {
            overlay.remove();
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM loaded, initializing CashFlowAllocator...');
    window.cashflowAllocator = new CashFlowAllocator();
});
