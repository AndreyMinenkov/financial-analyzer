// receipts-manager.js - Управление страницей поступлений
class ReceiptsManager {
    constructor(storage) {
        this.storage = storage;
        this.searchTerm = '';
        this.init();
    }

    init() {
        console.log('Initializing ReceiptsManager...');
    }

    updateTable() {
        const transactions = this.getFilteredTransactions();
        this.renderTable(transactions);
        this.updateSummary(transactions);
    }

    getFilteredTransactions() {
        let transactions = this.storage.getIncomingTransactions();
        console.log('Всего входящих транзакций:', transactions.length);

        // Исключаем транзакции по назначению платежа
        const excludePurposePhrases = [
            'перечисление собственных средств',
            'перечисление собственных дс под операционные расходы',
            'сс перечисление собственных средств',
            'перечисление средств на выплату заработной платы'
        ];

        // Исключаем транзакции по заказчику (наши компании)
        const excludeCustomerNames = [
            'сервис-интегратор ооо',
            'си уат ооо',
            'сервис-интегратор логистика ооо',
            'соир ооо',
            'сервис цм ооо',
            'управляющая компания сервис-интегратор ооо',
            'сервис-интегратор арктика ооо',
            'сервис-интегратор ао',
            'сервис-интегратор ут ооо',
            'сервис-интегратор сахалин ооо',
            'си логистика ооо',
            'сервис-интератор красноярск',
            // Производные названия
            'ооо соир',
            'сервис обслуживания и ремонта',
            'ооо си уат',
            'сервис-интегратор управление автотранспортом',
            'ооо си логистика',
            'ооо сервис-интегратор',
            'ооо сервис цм',
            'ооо управляющая компания сервис-интегратор',
            'ооо сервис-интегратор арктика',
            'ао сервис-интегратор',
            'ооо сервис-интегратор ут',
            'ооо сервис-интегратор сахалин'
        ];

        // Применяем фильтры исключения
        transactions = transactions.filter(t => {
            const purpose = (t.purpose || '').toLowerCase();
            const customer = this.getDisplayCounterCompany(t).toLowerCase();

            // Проверяем по назначению платежа
            for (const phrase of excludePurposePhrases) {
                if (purpose.includes(phrase)) {
                    console.log('Исключено по назначению:', purpose.substring(0, 50));
                    return false;
                }
            }

            // Проверяем по заказчику
            for (const name of excludeCustomerNames) {
                if (customer.includes(name)) {
                    console.log('Исключено по заказчику:', customer);
                    return false;
                }
            }

            return true;
        });

        console.log('После исключений осталось:', transactions.length);

        // Применяем поиск
        if (this.searchTerm) {
            const term = this.searchTerm.toLowerCase();
            transactions = transactions.filter(t => {
                const payer = (t.payer || '').toLowerCase();
                const purpose = (t.purpose || '').toLowerCase();
                return payer.includes(term) || purpose.includes(term);
            });
        }

        return transactions;
    }

    renderTable(transactions) {
        const tbody = document.getElementById('receiptsTableBody');

        if (!transactions || transactions.length === 0) {
            tbody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="8">
                        ${this.storage.getTransactions().length === 0
                            ? 'Загрузите выписки на странице "Загрузка выписок"'
                            : 'Нет данных по текущим фильтрам'}
                    </td>
                </tr>
            `;
            return;
        }

        let html = '';
        transactions.forEach(transaction => {
            // Получаем ИНН из загруженных данных или пытаемся извлечь
            let payerINN = transaction.payerINN || '';
            if (!payerINN && this.storage.getINNData()[transaction.payer]) {
                payerINN = this.storage.getINNData()[transaction.payer];
            }

            // Форматируем сумму
            const amountFormatted = this.storage.formatNumber(transaction.amount);

            html += `
                <tr>
                    <td>${this.getDisplayCounterCompany(transaction)}</td>
                    <td>${payerINN}</td>
                    <td>${transaction.ourCompany || ''}</td>
                    <td>${transaction.ourAccount || transaction.recipientAccount || ''}</td>
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
        const totalAmount = transactions.reduce((sum, t) => sum + (t.amount || 0), 0);

        document.getElementById('totalReceiptsCount').textContent = transactions.length;
        document.getElementById('totalReceiptsAmount').textContent =
            this.storage.formatCurrency(totalAmount);
    }

    getBankName(account) {
        if (!account) return '';
        const cleanAccount = account.replace(/\s/g, '');
        const mapping = {
            '40702810400000204768': 'ПСБ',
            '40702810240000407651': 'Сбер',
            '40702810000000011018': 'СДМ',
            '40702810907700000421': 'БКС',
            '40702810300000011971': 'МКБ',
            '40702810040000071672': 'Сбер',
            '40702810805010002132': 'МКБ',
            '40702810740000405629': 'Сбер',
            '40702810340000082125': 'Сбер',
            '40702810500000141745': 'ГПБ',
            '40702810500000009494': 'СДМ',
            '40702810100990012143': 'МИБ',
            '40702810240000097197': 'Сбер',
            '40701810540000401219': 'Сбер'
        };
        return mapping[cleanAccount] || 'Неизвестный банк';
    }

    isOurCompany(companyName) {
        if (!companyName) return false;
        const upperName = companyName.toUpperCase();
        return (
            upperName.includes("СЕРВИС-ИНТЕГРАТОР") ||
            upperName.includes("СИ УАТ") ||
            upperName.includes("СЕРВИС ЦМ") ||
            upperName.includes("СОИР") ||
            upperName.includes("Управляющая компания Сервис-Интегратор ООО") ||
            upperName.includes("СЕРВИС-ИНТЕГРАТОР УТ") ||
            upperName.includes("СЕРВИС-ИНТЕГРАТОР САХАЛИН") ||
            upperName.includes("СЕРВИС-ИНТЕГРАТОР ЛОГИСТИКА") ||
            upperName.includes("СЕРВИС-ИНТЕГРАТОР АО") ||
            upperName.includes("СЕРВИС-ИНТЕГРАТОР АРКТИКА")
        );
    }

    getDisplayCounterCompany(transaction) {
        const purpose = (transaction.purpose || "").toLowerCase();

        // Правило 1: СУЭК-Кузбасс АО (ищем "суэк-кузбасс" или часть "куз")
        if (purpose.includes("суэк-кузбасс") || purpose.includes("куз")) {
            return "СУЭК-Кузбасс АО";
        }

        // Правило 2: Запсибнефтехим ООО
        if (purpose.includes("запсибнефтехим")) {
            return "ЗАПСИБНЕФТЕХИМ ООО (Выплата финансирования)";
        }

        // Критерий 2: %% по депозитам
        const depositPhrases = [
            "уплата %",
            "уплата процентов по депозиту",
            "выплата начисленных процентов по депозиту",
            "Выплата начисленных процентов",
            "выплата начисленных процентов по заявке на размещение денежных средств",
            "размещение на депозите",
            "размещение вклада",
            "выплата процентов по депозиту"
        ];
        for (const phrase of depositPhrases) {
            if (purpose.includes(phrase)) {
                return "%% по депозитам";
            }
        }

        // Критерий 3: %% по НСО
        if (purpose.includes("уплата процентов по сделке нсо")) {
            return "%% по НСО";
        }

        // Критерий 4: Продажа ТС
        const tsPhrases = [
            "дкп",
            "договор купле-продажи",
            "договор купле продажи"
        ];
        for (const phrase of tsPhrases) {
            if (purpose.includes(phrase)) {
                return "Продажа ТС";
            }
        }
        // Проверка на госномер в формате А000АА123 (буква, три цифры, три буквы)
        const plateRegex = /[авекмнорстух]\d{3}[авекмнорстух]{2}\d{2,3}/i;
        if (plateRegex.test(purpose)) {
            return "Продажа ТС";
        }

        // Иначе возвращаем исходного заказчика
        return transaction.payer || transaction.counterCompany || "";
    }

    async loadINNData(file) {
        if (!file) return;

        try {
            // Читаем Excel-файл
            const workbook = await this.readExcelFile(file);
            const innData = this.parseINNDataFromExcel(workbook);

            // Сохраняем данные ИНН в хранилище
            this.storage.setINNData(innData);

            // Обновляем транзакции: заменяем наименования заказчиков по совпадению ИНН
            this.updateTransactionsWithINNData(innData);

            // Обновляем таблицу
            this.updateTable();

            window.app.showNotification(`Загружено ${Object.keys(innData).length} ИНН и обновлены транзакции`, 'success');

        } catch (error) {
            console.error('Error loading INN data:', error);
            window.app.showNotification('Ошибка загрузки ИНН', 'error');
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
            reader.onerror = (e) => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseINNDataFromExcel(workbook) {
        const innData = {};

        // Берем первый лист
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Конвертируем в JSON (массив объектов)
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Ищем столбцы "ИНН" и "Наименование"
        let innColumnIndex = -1;
        let nameColumnIndex = -1;

        if (rows.length > 0) {
            const headerRow = rows[0];
            for (let i = 0; i < headerRow.length; i++) {
                const cell = String(headerRow[i]).toLowerCase();
                if (cell.includes('инн')) innColumnIndex = i;
                if (cell.includes('наименование') || cell.includes('название')) nameColumnIndex = i;
            }
        }

        // Если не нашли заголовки, предполагаем, что первый столбец - ИНН, второй - Наименование
        if (innColumnIndex === -1) innColumnIndex = 0;
        if (nameColumnIndex === -1) nameColumnIndex = 1;

        // Обрабатываем строки, начиная со второй (индекс 1)
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length <= Math.max(innColumnIndex, nameColumnIndex)) continue;

            let inn = String(row[innColumnIndex] || '').trim();
            let name = String(row[nameColumnIndex] || '').trim();

            // Очищаем ИНН: оставляем только цифры
            inn = inn.replace(/\D/g, '');

            if (inn.length >= 10 && inn.length <= 12 && name) {
                innData[inn] = name;
            }
        }

        return innData;
    }

    updateTransactionsWithINNData(innData) {
        // Получаем все транзакции
        const allTransactions = this.storage.getTransactions();
        let updated = false;

        const updatedTransactions = allTransactions.map(transaction => {
            // Работаем только с входящими транзакциями, у которых есть payerINN
            if (transaction.direction === 'incoming' && transaction.payerINN) {
                const correctName = innData[transaction.payerINN];
                if (correctName) {
                    // Обновляем поля с наименованием плательщика/контрагента
                    transaction.payer = correctName;
                    transaction.counterCompany = correctName;
                    updated = true;
                }
            }
            return transaction;
        });

        if (updated) {
            // Сохраняем обновленные транзакции обратно в хранилище
            this.storage.setTransactions(updatedTransactions);
        }
    }

    exportToExcel() {
        const transactions = this.getFilteredTransactions();

        if (transactions.length === 0) {
            alert('Нет данных для экспорта');
            return;
        }

        // Подготовка данных для Excel
        const data = [
            ['Заказчик', 'ИНН заказчика', 'Юридическое лицо', 'Счет получателя',
             'Сумма', 'Банк получателя', 'Назначение платежа', 'Дата', 'Файл источника']
        ];

        transactions.forEach(t => {
            // Банк получателя
            const bank = t.ourBank || '';

            data.push([
                this.getDisplayCounterCompany(t),
                t.payerINN || '',
                t.ourCompany || '',
                t.ourAccount || t.recipientAccount || '',
                t.amount,
                bank,
                t.purpose || '',
                t.date || '',
                t.sourceFile || ''
            ]);
        });

        // Создание рабочей книги
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Поступления');

        // Устанавливаем числовой формат для столбца "Сумма"
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 4 });
            if (!ws[cellAddress]) continue;
            ws[cellAddress].t = 'n';
            ws[cellAddress].z = '#,##0.00';
        }

        // Автонастройка ширины столбцов
        const maxWidth = data.reduce((max, row) => Math.max(max, row.length), 0);
        const colWidths = [];
        for (let i = 0; i < maxWidth; i++) {
            let maxLength = 0;
            data.forEach(row => {
                const cellValue = row[i] || '';
                const length = String(cellValue).length;
                if (length > maxLength) maxLength = length;
            });
            colWidths.push({ wch: Math.min(maxLength + 2, 50) });
        }
        ws['!cols'] = colWidths;

        // Сохранение файла
        XLSX.writeFile(wb, `Поступления_${new Date().toISOString().split('T')[0]}.xlsx`);
    }

    // Метод для очистки поиска
    clearSearch() {
        this.searchTerm = '';
        document.getElementById('searchReceipts').value = '';
        this.updateTable();
    }

    // Метод для поиска
    searchTransactions(term) {
        this.searchTerm = term;
        this.updateTable();
    }
}
