// balances-manager.js - Управление страницей остатков
class BalancesManager {
    constructor(storage) {
        this.storage = storage;
        this.editingAccount = null;
        this.init();
    }

    init() {
        console.log('Initializing BalancesManager...');
        this.setupEventListeners();
    }

    setupEventListeners() {
        // Делегирование событий для редактируемых ячеек
        document.getElementById('balancesTableBody').addEventListener('click', (e) => {
            const cell = e.target.closest('.editable');
            if (cell && cell.dataset.account) {
                this.openEditModal(cell.dataset.account, cell.dataset.field);
            }
        });
    }

    updateTable() {
        const accounts = this.storage.getAccounts();
        const accountNumbers = Object.keys(accounts);

        const tbody = document.getElementById('balancesTableBody');

        if (accountNumbers.length === 0) {
            tbody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="9">
                        ${this.storage.getStatements().length === 0
                            ? 'Загрузите выписки на странице "Загрузка выписок"'
                            : 'Нет данных об остатках'}
                    </td>
                </tr>
            `;
            return;
        }

        let html = '';
        let totalBalance = 0;
        let totalDeposit = 0;
        let totalInterest = 0;

        accountNumbers.forEach(accountNumber => {
            const account = accounts[accountNumber];
            const depositData = this.storage.getDepositForAccount(accountNumber);

            // Получаем самые свежие данные по счету
            const latestStatement = this.getLatestStatementForAccount(accountNumber);
            const balance = latestStatement?.balance || account?.balance || 0;
            const statementDate = latestStatement?.date || account?.date || '';

            // Данные по депозиту
            const depositAmount = depositData?.amount || 0;
            const interestRate = depositData?.rate || 0;
            const depositStartDate = depositData?.startDate || '';

            // Расчет процентов и реального остатка
            const { interest, realBalance } = this.calculateInterestAndRealBalance(
                accountNumber, balance, depositAmount, interestRate, depositStartDate
            );

            // Форматирование чисел
            const balanceFormatted = this.storage.formatNumber(balance);
            const depositFormatted = depositAmount > 0 ? this.storage.formatNumber(depositAmount) : '';
            const interestFormatted = interest > 0 ? this.storage.formatNumber(interest) : '';
            const realBalanceFormatted = this.storage.formatNumber(realBalance);

            // Суммирование итогов
            totalBalance += balance;
            totalDeposit += depositAmount;
            totalInterest += interest;

            html += `
                <tr>
                    <td>${account.company || ''}</td>
                    <td>${account.bank || ''}</td>
                    <td>${accountNumber}</td>
                    <td class="number-cell">${balanceFormatted}</td>
                    <td class="number-cell editable" data-account="${accountNumber}" data-field="amount">
                        ${depositFormatted}
                    </td>
                    <td class="number-cell editable" data-account="${accountNumber}" data-field="rate">
                        ${interestRate > 0 ? interestRate.toFixed(2) + '%' : ''}
                    </td>
                    <td class="number-cell">${interestFormatted}</td>
                    <td class="number-cell">${realBalanceFormatted}</td>
                    <td>${statementDate}</td>
                </tr>
            `;
        });

        tbody.innerHTML = html;

        // Обновляем итоги
        this.updateSummary(totalBalance, totalDeposit, totalInterest, accountNumbers.length);
    }

    calculateInterestAndRealBalance(accountNumber, balance, depositAmount, interestRate, depositStartDate) {
        let interest = 0;
        let realBalance = balance;

        // Счета банка МИБ - особый расчет
        const mibAccounts = ['40702810700990012381', '40702810100990012143'];

        if (mibAccounts.includes(accountNumber)) {
            // Для МИБ: реальный остаток = конечный остаток по выписке + начисленные проценты
            // Сумма депозита отображается только информационно, не добавляется к остатку
            if (depositAmount > 0 && interestRate > 0 && depositStartDate) {
                const calculationDate = document.getElementById('calculationDate').value;
                const days = this.storage.getDaysBetween(depositStartDate, calculationDate);
                if (days > 0) {
                    interest = this.storage.calculateInterest(depositAmount, interestRate, days);
                    realBalance = balance + interest; // Только проценты добавляются к остатку
                }
            }
            return { interest, realBalance };
        }

        // Обычные счета: реальный остаток = конечный остаток по выписке + сумма депозита + начисленные проценты
        if (depositAmount > 0 && interestRate > 0 && depositStartDate) {
            const calculationDate = document.getElementById('calculationDate').value;
            const days = this.storage.getDaysBetween(depositStartDate, calculationDate);

            if (days > 0) {
                interest = this.storage.calculateInterest(depositAmount, interestRate, days);
                realBalance = balance + depositAmount + interest; // Депозит + проценты
            }
        }

        return { interest, realBalance };
    }

    getLatestStatementForAccount(accountNumber) {
        const statements = this.storage.getStatements();
        const accountStatements = statements.filter(s => s.account === accountNumber);

        if (accountStatements.length === 0) return null;

        // Сортируем по дате (от новых к старым)
        accountStatements.sort((a, b) => {
            const dateA = this.storage.parseDate(a.date);
            const dateB = this.storage.parseDate(b.date);
            return dateB - dateA;
        });

        return accountStatements[0];
    }

    updateSummary(totalBalance, totalDeposit, totalInterest, accountCount) {
        document.getElementById('totalAccountsCount').textContent = accountCount;
        document.getElementById('totalBalanceAmount').textContent =
            this.storage.formatCurrency(totalBalance);
        document.getElementById('totalInterestsAmount').textContent =
            this.storage.formatCurrency(totalInterest);
    }

    openEditModal(accountNumber, field) {
        const accounts = this.storage.getAccounts();
        const account = accounts[accountNumber];
        const depositData = this.storage.getDepositForAccount(accountNumber) || {};

        // Заполняем форму
        document.getElementById('editCompany').value = account?.company || '';
        document.getElementById('editAccount').value = accountNumber;
        document.getElementById('editDepositAmount').value = depositData.amount || '';
        document.getElementById('editInterestRate').value = depositData.rate || '';
        document.getElementById('editStartDate').value = depositData.startDate || '';

        // Сохраняем редактируемый счет
        this.editingAccount = accountNumber;

        // Показываем модальное окно
        document.getElementById('depositModal').classList.add('active');
    }

    saveDeposit() {
        if (!this.editingAccount) return;

        const amount = parseFloat(document.getElementById('editDepositAmount').value) || 0;
        const rate = parseFloat(document.getElementById('editInterestRate').value) || 0;
        const startDate = document.getElementById('editStartDate').value;

        const depositData = {
            amount: amount,
            rate: rate,
            startDate: startDate
        };

        this.storage.setDepositForAccount(this.editingAccount, depositData);

        // Закрываем модальное окно
        document.getElementById('depositModal').classList.remove('active');
        this.editingAccount = null;

        // Обновляем таблицу
        this.updateTable();

        window.app.showNotification('Данные по депозиту сохранены', 'success');
    }

    calculateInterests() {
        const calculationDate = document.getElementById('calculationDate').value;
        if (!calculationDate) {
            alert('Пожалуйста, укажите дату расчета');
            return;
        }

        this.updateTable();
        window.app.showNotification('Проценты рассчитаны', 'success');
    }

    async loadDepositData(file) {
        if (!file) return;

        try {
            console.log('Загрузка файла депозитов:', file.name, 'тип:', file.type);

            if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                await this.loadExcelDepositData(file);
            } else {
                const text = await this.readFile(file);
                const depositData = this.parseDepositData(text);

                // Сохраняем данные по депозитам
                Object.entries(depositData).forEach(([account, data]) => {
                    this.storage.setDepositForAccount(account, data);
                });

                window.app.showNotification(`Загружены данные по ${Object.keys(depositData).length} депозитам`, 'success');
            }

            this.updateTable();

        } catch (error) {
            console.error('Error loading deposit data:', error);
            window.app.showNotification('Ошибка загрузки данных по депозитам', 'error');
        }
    }

    async loadExcelDepositData(file) {
        try {
            console.log('Начало обработки Excel файла');
            const workbook = await this.readExcelFile(file);
            const depositData = this.parseDepositDataFromExcel(workbook);

            console.log('Найдено счетов с депозитами:', Object.keys(depositData).length);

            // Сохраняем данные по депозитам
            Object.entries(depositData).forEach(([account, data]) => {
                this.storage.setDepositForAccount(account, data);
            });

            window.app.showNotification(`Загружены данные по ${Object.keys(depositData).length} депозитам из Excel`, 'success');

        } catch (error) {
            console.error('Ошибка обработки Excel файла:', error);
            throw new Error('Ошибка обработки Excel файла: ' + error.message);
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    console.log('Чтение Excel файла завершено');
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (error) {
                    console.error('Ошибка парсинга Excel:', error);
                    reject(new Error('Ошибка чтения Excel файла: ' + error.message));
                }
            };
            reader.onerror = (e) => {
                console.error('Ошибка чтения файла:', e);
                reject(new Error('Ошибка чтения файла'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    parseDepositDataFromExcel(workbook) {
        const depositData = {};

        // Ищем лист "Свод"
        let sheetName = workbook.SheetNames.find(name =>
            name.toLowerCase().includes('свод')
        ) || workbook.SheetNames[0];

        console.log('Используем лист:', sheetName);
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Ищем заголовки
        let accountCol = 1; // Столбец B по умолчанию
        let amountCol = 2;  // Столбец C
        let rateCol = 3;    // Столбец D
        let startDateCol = 5; // Столбец F

        if (jsonData.length > 0) {
            const headerRow = jsonData[0];
            for (let i = 0; i < headerRow.length; i++) {
                const cellValue = String(headerRow[i] || '').toLowerCase();
                if (cellValue.includes('счет') || cellValue.includes('номер')) accountCol = i;
                if (cellValue.includes('сумма') && cellValue.includes('депозит')) amountCol = i;
                if (cellValue.includes('ставка')) rateCol = i;
                if (cellValue.includes('дата') && cellValue.includes('начал')) startDateCol = i;
            }
        }

        console.log('Столбцы для парсинга:', { accountCol, amountCol, rateCol, startDateCol });

        // Обрабатываем строки
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length <= Math.max(accountCol, amountCol, rateCol, startDateCol)) {
                continue;
            }

            const accountRaw = String(row[accountCol] || '').trim();
            const amountRaw = row[amountCol];
            const rateRaw = row[rateCol];
            const startDateRaw = row[startDateCol];

            console.log(`Строка ${i}:`, { accountRaw, amountRaw, rateRaw, startDateRaw });

            // Извлекаем номер счета (20 цифр)
            const accountMatch = accountRaw.match(/\d{20}/);
            if (!accountMatch) {
                console.log(`Не найден номер счета в строке ${i}: "${accountRaw}"`);
                continue;
            }

            const account = accountMatch[0];

            // Парсим сумму - только числа
            let amount = 0;
            if (amountRaw !== undefined && amountRaw !== null && amountRaw !== '') {
                if (typeof amountRaw === 'number') {
                    amount = amountRaw;
                } else {
                    const amountStr = String(amountRaw);
                    // Удаляем все не-цифры, кроме точки и минуса
                    const cleanStr = amountStr.replace(/[^\d.-]/g, '');
                    amount = parseFloat(cleanStr.replace(',', '.')) || 0;
                }
            }

            // Парсим ставку - только числа
            let rate = 0;
            if (rateRaw !== undefined && rateRaw !== null && rateRaw !== '') {
                if (typeof rateRaw === 'number') {
                    rate = rateRaw;
                } else {
                    const rateStr = String(rateRaw);
                    // Удаляем все не-цифры, кроме точки и минуса
                    const cleanStr = rateStr.replace(/[^\d.-]/g, '');
                    rate = parseFloat(cleanStr.replace(',', '.')) || 0;
                }
            }

            // Парсим дату
            let startDate = '';
            if (startDateRaw !== undefined && startDateRaw !== null && startDateRaw !== '') {
                if (typeof startDateRaw === 'number') {
                    // Даты Excel (число дней с 1 января 1900)
                    const date = new Date((startDateRaw - 25569) * 86400 * 1000);
                    startDate = date.toISOString().split('T')[0];
                } else if (typeof startDateRaw === 'string') {
                    const str = startDateRaw.trim();
                    // Проверяем разные форматы дат
                    if (str.match(/^\d{4}-\d{2}-\d{2}$/)) {
                        startDate = str; // YYYY-MM-DD
                    } else if (str.match(/^\d{2}\.\d{2}\.\d{4}$/)) {
                        // DD.MM.YYYY
                        const parts = str.split('.');
                        const date = new Date(parts[2], parts[1] - 1, parts[0]);
                        startDate = date.toISOString().split('T')[0];
                    } else {
                        // Пробуем как обычную дату
                        const date = new Date(str);
                        if (!isNaN(date.getTime())) {
                            startDate = date.toISOString().split('T')[0];
                        }
                    }
                }
            }

            // Сохраняем только если есть депозит (сумма > 0) И ставка > 0
            if (account && amount > 0 && rate > 0) {
                depositData[account] = {
                    amount: amount,
                    rate: rate,
                    startDate: startDate || new Date().toISOString().split('T')[0]
                };

                console.log(`✓ Счет ${account}: сумма=${amount}, ставка=${rate}%, дата начала=${depositData[account].startDate}`);
            } else {
                console.log(`✗ Счет ${account}: сумма=${amount}, ставка=${rate}% (не сохраняем)`);
            }
        }

        console.log('Итого найдено счетов с депозитами:', Object.keys(depositData).length);
        return depositData;
    }

    readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error('Ошибка чтения файла'));
            reader.readAsText(file, 'UTF-8');
        });
    }

    parseDepositData(text) {
        const lines = text.split('\n');
        const depositData = {};

        lines.forEach(line => {
            const trimmed = line.trim();
            if (!trimmed) return;

            // Пробуем разные форматы: счет, сумма, ставка, дата начала
            const parts = trimmed.split(/[\t,;]/).map(p => p.trim());

            if (parts.length >= 3) {
                const account = parts[0].replace(/\s/g, '');
                const amount = parseFloat(parts[1].replace(',', '.')) || 0;
                const rate = parseFloat(parts[2].replace(',', '.')) || 0;
                const startDate = parts[3] || '';

                if (account && amount > 0 && rate > 0) {
                    depositData[account] = {
                        amount: amount,
                        rate: rate,
                        startDate: startDate
                    };
                }
            }
        });

        return depositData;
    }

    exportToExcel() {
        const accounts = this.storage.getAccounts();
        const accountNumbers = Object.keys(accounts);

        if (accountNumbers.length === 0) {
            alert('Нет данных для экспорта');
            return;
        }

        const calculationDate = document.getElementById('calculationDate').value;

        // Сортируем счета по порядку из accountMapping
        let sortedAccountNumbers = [...accountNumbers];
        try {
            const parser = new BankStatementParser();
            const accountMapping = parser.loadAccountMapping();
            const accountOrder = Object.keys(accountMapping);

            sortedAccountNumbers.sort((a, b) => {
                const indexA = accountOrder.indexOf(a);
                const indexB = accountOrder.indexOf(b);

                if (indexA !== -1 && indexB !== -1) return indexA - indexB;
                if (indexA !== -1) return -1;
                if (indexB !== -1) return 1;
                return a.localeCompare(b);
            });
        } catch (error) {
            console.warn('Не удалось отсортировать счета:', error);
        }

        // Подготовка данных для Excel (без столбца "Ставка, %")
        const data = [
            ['Компания', 'Банк', 'Счёт', 'Остаток по выписке', 'Вернувшийся депозит',
             'Начисленные проценты', 'Реальный остаток', 'Дата']
        ];

        sortedAccountNumbers.forEach(accountNumber => {
            const account = accounts[accountNumber];
            const depositData = this.storage.getDepositForAccount(accountNumber);

            // Получаем самые свежие данные
            const latestStatement = this.getLatestStatementForAccount(accountNumber);
            const balance = latestStatement?.balance || account?.balance || 0;
            const statementDate = latestStatement?.date || account?.date || '';

            // Данные по депозиту
            const depositAmount = depositData?.amount || 0;
            const interestRate = depositData?.rate || 0;
            const depositStartDate = depositData?.startDate || '';

            // Расчет процентов и реального остатка
            const { interest, realBalance } = this.calculateInterestAndRealBalance(
                accountNumber, balance, depositAmount, interestRate, depositStartDate
            );

            // Форматирование
            const balanceFormatted = this.storage.formatNumber(balance);
            const depositFormatted = this.storage.formatNumber(depositAmount);
            const interestFormatted = this.storage.formatNumber(interest);
            const realBalanceFormatted = this.storage.formatNumber(realBalance);

            data.push([
                account.company || '',
                account.bank || '',
                accountNumber,
                balanceFormatted,
                depositFormatted,
                interestFormatted,
                realBalanceFormatted,
                statementDate,
                depositStartDate,
                calculationDate
            ]);
        });

        // Создание книги Excel
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Остатки');

        // Настройка ширины колонок
        const colWidths = [
            { wch: 30 }, // Компания
            { wch: 15 }, // Банк
            { wch: 20 }, // Счет
            { wch: 20 }, // Остаток
            { wch: 15 }, // Вернувшийся депозит
            { wch: 15 }, // Начисленные проценты
            { wch: 20 }, // Реальный остаток
            { wch: 12 } // Дата
        ];
        ws['!cols'] = colWidths;

        // Экспорт файла
        const date = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `Остатки_${date}.xlsx`);

        window.app.showNotification(`Экспортировано ${sortedAccountNumbers.length} счетов`, 'success');
    }
}
