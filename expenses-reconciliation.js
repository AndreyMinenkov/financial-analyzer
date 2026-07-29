// expenses-reconciliation.js — Сверка расходов из выписок с загруженным Excel-файлом
class ExpensesReconciliationManager {
    constructor(storage) {
        this.storage = storage;
        this.excelData = [];        // [{amount, purpose, counterparty, organization}, ...]
        this.reconciliationResult = null; // {commissions: [], deviations: [], matched: [], unmatchedExcel: []}
        this.searchTerm = '';
        this.init();
    }

    init() {
        console.log('Initializing ExpensesReconciliationManager...');
    }

    // ── Таблица расходов из выписок ──────────────────────
    updateTable() {
        const transactions = this.getFilteredExpenses();
        this.renderExpensesTable(transactions);
        this.updateSummary(transactions);
    }

    getFilteredExpenses() {
        let transactions = this.storage.getOutgoingTransactions();
        console.log('Всего исходящих транзакций:', transactions.length);

        if (this.searchTerm) {
            const term = this.searchTerm.toLowerCase();
            transactions = transactions.filter(t => {
                const counterparty = (t.counterCompany || '').toLowerCase();
                const purpose = (t.purpose || '').toLowerCase();
                const ourCompany = (t.ourCompany || '').toLowerCase();
                return counterparty.includes(term) || purpose.includes(term) || ourCompany.includes(term);
            });
        }

        return transactions;
    }

    renderExpensesTable(transactions) {
        const tbody = document.getElementById('expensesTableBody');
        if (!tbody) return;

        if (!transactions || transactions.length === 0) {
            tbody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="8">
                        ${this.storage.getTransactions().length === 0
                            ? 'Загрузите выписки на странице "Загрузка выписок"'
                            : 'Нет расходных операций по текущим фильтрам'}
                    </td>
                </tr>
            `;
            return;
        }

        let html = '';
        transactions.forEach(transaction => {
            const amountFormatted = this.storage.formatNumber(Math.abs(transaction.amount));
            html += `
                <tr>
                    <td>${transaction.counterCompany || ''}</td>
                    <td>${transaction.counterAccount || ''}</td>
                    <td>${transaction.ourCompany || ''}</td>
                    <td>${transaction.ourAccount || ''}</td>
                    <td class="number-cell">${amountFormatted}</td>
                    <td>${transaction.ourBank || ''}</td>
                    <td>${transaction.purpose || ''}</td>
                    <td>${transaction.date || ''}</td>
                </tr>
            `;
        });

        tbody.innerHTML = html;
    }

    updateSummary(transactions) {
        const totalAmount = transactions.reduce((sum, t) => sum + Math.abs(t.amount || 0), 0);
        const countEl = document.getElementById('totalExpensesCount');
        const amountEl = document.getElementById('totalExpensesAmount');
        if (countEl) countEl.textContent = transactions.length;
        if (amountEl) amountEl.textContent = this.storage.formatCurrency(totalAmount);
    }

    // ── Загрузка Excel-файла ─────────────────────────────
    async loadExcelFile(file) {
        if (!file) return { success: false, message: 'Файл не выбран' };

        try {
            const workbook = await this.readExcelFile(file);
            this.excelData = this.parseExpensesExcel(workbook);

            if (this.excelData.length === 0) {
                return { success: false, message: 'Не найдено данных в файле. Проверьте формат.' };
            }

            return {
                success: true,
                message: `Загружено ${this.excelData.length} строк из Excel`,
                count: this.excelData.length
            };
        } catch (error) {
            console.error('Ошибка загрузки Excel:', error);
            return { success: false, message: 'Ошибка чтения файла: ' + error.message };
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (error) {
                    reject(new Error('Ошибка чтения Excel файла: ' + error.message));
                }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseExpensesExcel(workbook) {
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (rows.length < 2) return [];

        // Определяем колонки по заголовкам
        let amountCol = -1, purposeCol = -1, counterpartyCol = -1, orgCol = -1;
        const headerRow = rows[0];

        headerRow.forEach((cell, i) => {
            const c = String(cell || '').toLowerCase().trim();
            if (c.includes('списание') || c.includes('сумма')) amountCol = i;
            if (c.includes('назначение')) purposeCol = i;
            if (c.includes('контрагент')) counterpartyCol = i;
            if (c.includes('организац') || c.includes('юрлиц') || c.includes('юридическое')) orgCol = i;
        });

        // Если не нашли по заголовкам — используем позиции по умолчанию
        if (amountCol === -1) amountCol = 0;
        if (purposeCol === -1) purposeCol = 1;
        if (counterpartyCol === -1) counterpartyCol = 2;
        if (orgCol === -1) orgCol = 3;

        console.log('Колонки Excel:', { amountCol, purposeCol, counterpartyCol, orgCol });

        const result = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const rawAmount = row[amountCol];
            let amount = 0;
            if (rawAmount !== undefined && rawAmount !== null && rawAmount !== '') {
                if (typeof rawAmount === 'number') {
                    amount = rawAmount;
                } else {
                    const str = String(rawAmount).replace(/[^\d.,\-]/g, '').replace(',', '.');
                    amount = parseFloat(str) || 0;
                }
            }
            // Берём абсолютное значение (расходы всегда положительные по модулю)
            amount = Math.abs(amount);

            if (amount === 0) continue;

            result.push({
                amount: Math.round(amount * 100) / 100,
                purpose: String(row[purposeCol] || '').trim(),
                counterparty: String(row[counterpartyCol] || '').trim(),
                organization: String(row[orgCol] || '').trim()
            });
        }

        console.log('Распаршено строк из Excel:', result.length);
        return result;
    }

    // ── Сверка ───────────────────────────────────────────
    reconcile() {
        const expenses = this.storage.getOutgoingTransactions();

        if (expenses.length === 0) {
            return { success: false, message: 'Нет расходных операций в выписках' };
        }
        if (this.excelData.length === 0) {
            return { success: false, message: 'Не загружен Excel-файл для сверки' };
        }

        const commissions = [];    // Банковские комиссии
        const deviations = [];     // Расхождения (есть в выписке, нет в Excel)
        const matched = [];        // Успешно сопоставленные
        const unmatchedExcel = []; // Есть в Excel, нет в выписке (копия excelData, из которой удаляем сопоставленные)

        // Копия excelData для отслеживания несопоставленных
        const excelRemaining = this.excelData.map((item, idx) => ({ ...item, _idx: idx }));

        for (const expense of expenses) {
            const absAmount = Math.abs(expense.amount);
            const counterCompany = (expense.counterCompany || '').trim();

            // Проверяем, является ли это банковской комиссией (по назначению платежа)
            const isBankCommission = this.isBankCommission(expense.purpose || '');

            if (isBankCommission) {
                commissions.push({
                    date: expense.date || '',
                    ourCompany: expense.ourCompany || '',
                    ourBank: expense.ourBank || '',
                    counterCompany: counterCompany,
                    amount: absAmount,
                    purpose: expense.purpose || '',
                    sourceFile: expense.sourceFile || ''
                });
                continue; // Комиссии не участвуют в сверке с Excel
            }

            // Ищем соответствие в Excel по сумме И контрагенту
            const matchIndex = this.findMatch(expense, excelRemaining);

            if (matchIndex !== -1) {
                const match = excelRemaining[matchIndex];
                matched.push({
                    date: expense.date || '',
                    amount: absAmount,
                    counterparty: counterCompany,
                    purpose: expense.purpose || '',
                    ourCompany: expense.ourCompany || '',
                    ourBank: expense.ourBank || '',
                    excelAmount: match.amount,
                    excelCounterparty: match.counterparty,
                    excelPurpose: match.purpose,
                    excelOrganization: match.organization
                });
                // Удаляем сопоставленный элемент из оставшихся
                excelRemaining.splice(matchIndex, 1);
            } else {
                // Отклонение: есть в выписке, но нет в Excel
                deviations.push({
                    type: 'only_in_statement',
                    date: expense.date || '',
                    amount: absAmount,
                    counterparty: counterCompany,
                    purpose: expense.purpose || '',
                    ourCompany: expense.ourCompany || '',
                    ourBank: expense.ourBank || '',
                    sourceFile: expense.sourceFile || ''
                });
            }
        }

        // Оставшиеся в Excel — это расходы из Excel, не найденные в выписках
        for (const item of excelRemaining) {
            deviations.push({
                type: 'only_in_excel',
                date: '',
                amount: item.amount,
                counterparty: item.counterparty,
                purpose: item.purpose,
                ourCompany: item.organization,
                ourBank: '',
                sourceFile: ''
            });
        }

        this.reconciliationResult = {
            commissions,
            deviations,
            matched,
            unmatchedExcel: excelRemaining,
            summary: {
                totalExpenses: expenses.length,
                totalExcelRows: this.excelData.length,
                commissionsCount: commissions.length,
                commissionsAmount: commissions.reduce((s, c) => s + c.amount, 0),
                matchedCount: matched.length,
                matchedAmount: matched.reduce((s, m) => s + m.amount, 0),
                deviationsCount: deviations.length,
                deviationsAmount: deviations.reduce((s, d) => s + (d.amount || 0), 0)
            }
        };

        console.log('Результат сверки:', this.reconciliationResult.summary);

        return {
            success: true,
            message: `Сверка завершена: найдено ${commissions.length} комиссий, ${matched.length} сопоставлено, ${deviations.length} отклонений`,
            result: this.reconciliationResult
        };
    }

    isBankCommission(purpose) {
        // Проверка по назначению платежа
        if (!purpose) return false;
        const upper = purpose.toUpperCase();

        // Шаг 1: точные фразы, однозначно указывающие на банковскую комиссию
        const exactKeywords = [
            'БАНКОВСКАЯ КОМИССИЯ',
            'КОМИССИЯ БАНКА',
            'КОМИССИОННОЕ ВОЗНАГРАЖДЕНИЕ',
            'КОМИССИОННЫЙ СБОР',
            'ВОЗНАГРАЖДЕНИЕ БАНКА',
            'СМС-ИНФОРМИРОВАНИЕ',
            'СМС ИНФОРМИРОВАНИЕ',
            'ДИСТАНЦИОННОЕ БАНКОВСКОЕ ОБСЛУЖИВАНИЕ',
            'УСЛУГИ БАНКА',
            'БАНКОВСКАЯ УСЛУГА',
            'БАНКОВСКИЙ ТАРИФ',
            'ТАРИФ БАНКА',
            'СОГЛАСНО ТАРИФАМ БАНКА',
            'ПЛАТА ЗА ВЫПУСК КАРТ',
            'ОБСЛУЖИВАНИЕ СЧЕТА',
            'ОБСЛУЖИВАНИЕ РАСЧЕТНОГО',
            'ВЕДЕНИЕ СЧЕТА',
            'ВЕДЕНИЕ РАСЧЕТНОГО'
        ];
        for (const kw of exactKeywords) {
            if (upper.includes(kw)) return true;
        }

        // Шаг 2: общие фразы с КОМИССИЯ — широкая проверка с фильтром исключений
        if (upper.includes('КОМИССИЯ')) {
            // Исключаем платежи, не являющиеся банковскими комиссиями
            const nonBankPatterns = [
                'ЗАРПЛАТА', 'ЗАРАБОТНАЯ', 'ВЫПЛАТА ЗАРПЛАТЫ',
                'ЛИЗИНГ', 'АРЕНДА', 'АРЕНДН',
                'УСЛУГИ СВЯЗИ', 'СОТОВАЯ СВЯЗЬ', 'ИНТЕРНЕТ',
                'АБОНЕНТСКАЯ ПЛАТА', 'АБОНЕНТСКОЕ',
                'КОММУНАЛЬНЫЕ', 'ЭЛЕКТРОЭНЕРГИЯ', 'ТЕПЛО',
                'ОХРАНА', 'СТРАХОВАНИЕ', 'СТРАХОВ',
                'ПОСТАВКА', 'ОТГРУЗКА', 'ТОВАР',
                'ШТРАФ', 'ПЕНИ', 'НЕУСТОЙКА',
                'ГОСПОШЛИНА', 'НАЛОГ', 'СБОР ЗА',
                'ПРЕДОПЛАТА', 'АВАНС'
            ];
            const isNonBank = nonBankPatterns.some(pattern => upper.includes(pattern));
            if (!isNonBank) return true;
        }

        return false;
    }

    findMatch(expense, excelRemaining) {
        const absAmount = Math.round(Math.abs(expense.amount) * 100) / 100;

        for (let i = 0; i < excelRemaining.length; i++) {
            const excelItem = excelRemaining[i];
            // Сверка только по сумме (точное совпадение до 2 знаков)
            const excelAmount = Math.round(excelItem.amount * 100) / 100;
            if (Math.abs(absAmount - excelAmount) > 0.01) continue;

            // Сумма совпала — считаем соответствием
            return i;
        }
        return -1;
    }

    normalizeName(name) {
        if (!name) return '';
        return name
            .replace(/^["']+|["']+$/g, '')
            .replace(/^ИНН\s+\d+\s+/, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    companiesMatch(name1, name2) {
        if (!name1 || !name2) return false;
        const n1 = name1.toUpperCase().replace(/["«»'']/g, '').replace(/\s+/g, ' ').trim();
        const n2 = name2.toUpperCase().replace(/["«»'']/g, '').replace(/\s+/g, ' ').trim();

        // Точное совпадение
        if (n1 === n2) return true;

        // Одно содержит другое
        if (n1.includes(n2) || n2.includes(n1)) return true;

        // Убираем организационно-правовую форму для сравнения
        const stripOPF = (s) => s.replace(/^(ООО|АО|ПАО|ЗАО|НКО|ИП)\s+/i, '').trim();
        return stripOPF(n1) === stripOPF(n2);
    }

    // ── Отрисовка результатов сверки ────────────────────
    renderReconciliationResult() {
        if (!this.reconciliationResult) return;

        // Собираем все строки в единый массив
        this.allReconRows = this._buildReconRows();
        this._currentReconFilter = 'all';

        // Скрываем исходную таблицу и показываем результат
        const tableContainer = document.getElementById('expensesTableContainer');
        const summaryCards = document.getElementById('expensesSummaryCards');
        const searchSection = document.querySelector('#expenses-reconciliation-page .search-section');
        if (tableContainer) tableContainer.style.display = 'none';
        if (summaryCards) summaryCards.style.display = 'none';
        if (searchSection) searchSection.style.display = 'none';

        // Обновляем сводку
        const s = this.reconciliationResult.summary;
        const onlyStatement = this.reconciliationResult.deviations.filter(d => d.type === 'only_in_statement').length;
        const onlyExcel = this.reconciliationResult.deviations.filter(d => d.type === 'only_in_excel').length;

        document.getElementById('recMatchedCount').textContent = s.matchedCount;
        document.getElementById('recOnlyStatementCount').textContent = onlyStatement;
        document.getElementById('recOnlyExcelCount').textContent = onlyExcel;
        document.getElementById('recCommissionsCount').textContent = s.commissionsCount;
        document.getElementById('recCommissionsAmount').textContent = this.storage.formatCurrency(s.commissionsAmount);

        // Комиссии — сворачиваемый блок
        const commBlock = document.getElementById('reconCommissionsBlock');
        if (s.commissionsCount > 0) {
            commBlock.style.display = '';
            this._renderCommissionsCompact(this.reconciliationResult.commissions);
        } else {
            commBlock.style.display = 'none';
        }

        // Единая таблица
        this._renderReconTable(this.allReconRows);

        // Показать панель
        document.getElementById('reconciliationResultSection').style.display = 'block';
        document.getElementById('exportExpensesReconciliationBtn').disabled = false;

        // Фильтр-табы
        this._setupReconFilterTabs();
    }

    _buildReconRows() {
        const rows = [];

        // Сопоставленные
        for (const m of this.reconciliationResult.matched) {
            rows.push({
                type: 'matched',
                date: m.date,
                stmtAmount: m.amount,
                excelAmount: m.excelAmount,
                stmtCounterparty: m.counterparty,
                excelCounterparty: m.excelCounterparty,
                stmtPurpose: m.purpose,
                excelPurpose: m.excelPurpose,
                org: m.ourCompany || m.excelOrganization
            });
        }

        // Только в выписке
        for (const d of this.reconciliationResult.deviations) {
            if (d.type === 'only_in_statement') {
                rows.push({
                    type: 'only_statement',
                    date: d.date,
                    stmtAmount: d.amount,
                    excelAmount: null,
                    stmtCounterparty: d.counterparty,
                    excelCounterparty: '',
                    stmtPurpose: d.purpose,
                    excelPurpose: '',
                    org: d.ourCompany
                });
            }
        }

        // Только в Excel
        for (const d of this.reconciliationResult.deviations) {
            if (d.type === 'only_in_excel') {
                rows.push({
                    type: 'only_excel',
                    date: '',
                    stmtAmount: null,
                    excelAmount: d.amount,
                    stmtCounterparty: '',
                    excelCounterparty: d.counterparty,
                    stmtPurpose: '',
                    excelPurpose: d.purpose,
                    org: d.ourCompany
                });
            }
        }

        // Сортировка: сначала несовпадения, потом сопоставленные
        rows.sort((a, b) => {
            const order = { only_statement: 0, only_excel: 1, matched: 2 };
            return order[a.type] - order[b.type];
        });

        return rows;
    }

    _statusIcon(type) {
        switch (type) {
            case 'matched':
                return '<span class="recon-status recon-status-matched" title="Сопоставлено"><i class="fas fa-check-circle"></i></span>';
            case 'only_statement':
                return '<span class="recon-status recon-status-statement" title="Только в выписке"><i class="fas fa-arrow-right"></i></span>';
            case 'only_excel':
                return '<span class="recon-status recon-status-excel" title="Только в Excel"><i class="fas fa-file-excel"></i></span>';
            default:
                return '';
        }
    }

    _renderReconTable(rows) {
        const tbody = document.getElementById('reconMainTableBody');
        if (!tbody) return;

        if (rows.length === 0) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="9">Нет данных для отображения</td></tr>';
            return;
        }

        let html = '';
        for (const r of rows) {
            const rowClass = `recon-row-${r.type}`;
            html += `<tr class="${rowClass}">`;
            html += `<td>${this._statusIcon(r.type)}</td>`;
            html += `<td>${r.date || '—'}</td>`;
            html += `<td class="number-cell">${r.stmtAmount != null ? this.storage.formatNumber(r.stmtAmount) : '—'}</td>`;
            html += `<td class="number-cell">${r.excelAmount != null ? this.storage.formatNumber(r.excelAmount) : '—'}</td>`;
            html += `<td>${this._cell(r.stmtCounterparty)}</td>`;
            html += `<td>${this._cell(r.excelCounterparty)}</td>`;
            html += `<td>${this._cell(r.stmtPurpose)}</td>`;
            html += `<td>${this._cell(r.excelPurpose)}</td>`;
            html += `<td>${this._cell(r.org)}</td>`;
            html += `</tr>`;
        }

        tbody.innerHTML = html;
        this._updateReconFilterStats(rows.length);
    }

    _cell(val) {
        return val || '<span class="recon-empty-cell">—</span>';
    }

    _renderCommissionsCompact(commissions) {
        const tbody = document.getElementById('commissionsTableBody');
        if (!tbody) return;

        // Группируем по юрлицу
        const grouped = {};
        for (const c of commissions) {
            const key = c.ourCompany || 'Неизвестно';
            if (!grouped[key]) grouped[key] = { items: [], total: 0 };
            grouped[key].items.push(c);
            grouped[key].total += c.amount;
        }

        let html = '';
        let grandTotal = 0;
        for (const [company, group] of Object.entries(grouped)) {
            html += `<tr class="recon-group-header"><td colspan="5"><strong>${company}</strong> — ${this.storage.formatCurrency(group.total)}</td></tr>`;
            for (const c of group.items) {
                html += `<tr>
                    <td>${c.date}</td>
                    <td>${c.counterCompany}</td>
                    <td class="number-cell">${this.storage.formatNumber(c.amount)}</td>
                    <td>${c.purpose}</td>
                    <td>${c.ourCompany}</td>
                </tr>`;
                grandTotal += c.amount;
            }
        }

        // Итоговая строка
        html += `<tr class="recon-total-row">
            <td colspan="2"><strong>ИТОГО</strong></td>
            <td class="number-cell"><strong>${this.storage.formatNumber(grandTotal)}</strong></td>
            <td colspan="2"></td>
        </tr>`;

        tbody.innerHTML = html;
    }

    _setupReconFilterTabs() {
        const tabs = document.querySelectorAll('.recon-filter-tab');
        tabs.forEach(tab => {
            const newTab = tab.cloneNode(true);
            tab.parentNode.replaceChild(newTab, tab);
        });

        document.querySelectorAll('.recon-filter-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.recon-filter-tab').forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                this._applyReconFilter(tab.dataset.filter);
            });
        });
    }

    _applyReconFilter(filter) {
        this._currentReconFilter = filter;
        let filtered = this.allReconRows;
        if (filter !== 'all') {
            filtered = this.allReconRows.filter(r => r.type === filter);
        }
        this._renderReconTable(filtered);
    }

    _updateReconFilterStats(total) {
        const el = document.getElementById('reconFilterStats');
        if (!el) return;
        const matched = this.allReconRows.filter(r => r.type === 'matched').length;
        const stmt = this.allReconRows.filter(r => r.type === 'only_statement').length;
        const excel = this.allReconRows.filter(r => r.type === 'only_excel').length;
        el.textContent = `Показано: ${total} строк (✅ ${matched} | 📄 ${stmt} | 📊 ${excel})`;
    }

    // ── Экспорт результатов ─────────────────────────────
    exportToExcel() {
        if (!this.reconciliationResult) {
            alert('Нет результатов сверки. Выполните сверку перед экспортом.');
            return;
        }

        const wb = XLSX.utils.book_new();

        // Лист 1: Комиссии
        const commissionsData = [['Дата', 'Контрагент (Банк)', 'Сумма', 'Банк получателя', 'Назначение', 'Юридическое лицо', 'Файл источника']];
        for (const c of this.reconciliationResult.commissions) {
            commissionsData.push([c.date, c.counterCompany, c.amount, c.ourBank, c.purpose, c.ourCompany, c.sourceFile]);
        }
        // Итого
        const commTotal = commissionsData.slice(1).reduce((s, r) => s + (r[2] || 0), 0);
        commissionsData.push(['ИТОГО', '', commTotal, '', '', '', '']);

        const wsCommissions = XLSX.utils.aoa_to_sheet(commissionsData);
        XLSX.utils.book_append_sheet(wb, wsCommissions, 'Комиссии');

        // Лист 2: Сопоставленные расходы
        const matchedData = [['Дата', 'Сумма (выписка)', 'Контрагент (выписка)', 'Назначение (выписка)',
            'Юрлицо', 'Банк', 'Сумма (Excel)', 'Контрагент (Excel)', 'Назначение (Excel)', 'Организация (Excel)']];
        for (const m of this.reconciliationResult.matched) {
            matchedData.push([m.date, m.amount, m.counterparty, m.purpose,
                m.ourCompany, m.ourBank, m.excelAmount, m.excelCounterparty, m.excelPurpose, m.excelOrganization]);
        }
        const wsMatched = XLSX.utils.aoa_to_sheet(matchedData);
        XLSX.utils.book_append_sheet(wb, wsMatched, 'Сопоставлено');

        // Лист 3: Отклонения
        const deviationsData = [['Дата', 'Сумма', 'Контрагент', 'Назначение', 'Юрлицо', 'Тип отклонения']];
        for (const d of this.reconciliationResult.deviations) {
            deviationsData.push([d.date, d.amount, d.counterparty, d.purpose, d.ourCompany,
                d.type === 'only_in_statement' ? 'Только в выписке' : 'Только в Excel']);
        }
        const devTotal = deviationsData.slice(1).reduce((s, r) => s + (r[1] || 0), 0);
        deviationsData.push(['ИТОГО', devTotal, '', '', '', '']);

        const wsDeviations = XLSX.utils.aoa_to_sheet(deviationsData);
        XLSX.utils.book_append_sheet(wb, wsDeviations, 'Отклонения');

        // Настройка ширины колонок
        for (const ws of [wsCommissions, wsMatched, wsDeviations]) {
            const cols = [];
            const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
            for (let col = 0; col <= range.e.c; col++) {
                let maxLen = 5;
                for (let row = 0; row <= range.e.r; row++) {
                    const cell = ws[XLSX.utils.encode_cell({ r: row, c: col })];
                    if (cell && cell.v) maxLen = Math.max(maxLen, String(cell.v).length);
                }
                cols.push({ wch: Math.min(maxLen + 3, 60) });
            }
            ws['!cols'] = cols;
        }

        const date = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `Сверка_оплат_${date}.xlsx`, { cellStyles: true });
    }

    // ── Очистка ─────────────────────────────────────────
    clearAll() {
        this.excelData = [];
        this.reconciliationResult = null;
        this.searchTerm = '';

        const searchInput = document.getElementById('searchExpenses');
        if (searchInput) searchInput.value = '';

        const fileInfo = document.getElementById('expensesExcelFileInfo');
        if (fileInfo) fileInfo.innerHTML = '<i class="fas fa-info-circle"></i> Столбцы: Списание, Назначение платежа, Контрагент, Организация';

        const reconcileBtn = document.getElementById('reconcileExpensesBtn');
        if (reconcileBtn) reconcileBtn.disabled = true;

        const exportBtn = document.getElementById('exportExpensesReconciliationBtn');
        if (exportBtn) exportBtn.disabled = true;

        const resultSection = document.getElementById('reconciliationResultSection');
        if (resultSection) resultSection.style.display = 'none';

        // Возвращаем исходную таблицу расходов
        const tableContainer = document.getElementById('expensesTableContainer');
        const summaryCards = document.getElementById('expensesSummaryCards');
        const searchSection = document.querySelector('#expenses-reconciliation-page .search-section');
        if (tableContainer) tableContainer.style.display = '';
        if (summaryCards) summaryCards.style.display = '';
        if (searchSection) searchSection.style.display = '';

        this.updateTable();
    }

    // ── Поиск ───────────────────────────────────────────
    searchTransactions(term) {
        this.searchTerm = term;
        this.updateTable();
    }

    clearSearch() {
        this.searchTerm = '';
        const input = document.getElementById('searchExpenses');
        if (input) input.value = '';
        this.updateTable();
    }
}