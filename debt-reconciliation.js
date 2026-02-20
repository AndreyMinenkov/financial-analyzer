// debt-reconciliation.js - Модуль для сверки долгов с полным сохранением форматирования через сервер
class DebtReconciliationManager {
    constructor(storage) {
        this.storage = storage;
        this.debtData = [];           // данные из файла 1
        this.debtHeaders = [];         // заголовки файла 1
        this.debtFile = null;          // оригинальный файл
        this.debtFileName = '';        // имя файла
        this.receiptsData = [];        // данные из файла 2 (только с датами)
        this.processedDocuments = [];   // логи обработанных документов
        this.currentDate = new Date();
        
        // Загружаем список целевых контрагентов из localStorage или используем стандартный
        this.loadTargetContractors();
        
        this.stats = {
            totalDocuments: 0,
            foundDocuments: 0,
            updatedDocuments: 0,
            errors: []
        };

        // Индексы колонок в файле 1
        this.COLUMNS = {
            DOCUMENT_NAME: 0,      // A
            DEBT_AMOUNT: 11,       // L
            OVERDUE: 14,           // O
            DAYS: 17,              // R
            NOT_OVERDUE: 19,       // T - не просрочено
            INTERVAL_1_15: 20,     // U - 1-15 дней
            INTERVAL_16_29: 21,    // V - 16-29 дней
            INTERVAL_30_89: 22,    // W - 30-89 дней
            INTERVAL_90_179: 23,   // X - 90-179 дней
            INTERVAL_180_PLUS: 24, // Y - 180+ дней
        };
    }

    // Загрузка списка целевых контрагентов из localStorage
    loadTargetContractors() {
        const stored = localStorage.getItem('targetContractors');
        if (stored) {
            try {
                this.TARGET_CONTRAGENTS = JSON.parse(stored);
                console.log('Загружен список контрагентов:', this.TARGET_CONTRAGENTS);
            } catch (e) {
                console.error('Ошибка загрузки списка контрагентов', e);
                this.TARGET_CONTRAGENTS = ['ВАНКОРНЕФТЬ АО', 'РН-Ванкор ООО'];
            }
        } else {
            // Стандартный список по умолчанию
            this.TARGET_CONTRAGENTS = ['ВАНКОРНЕФТЬ АО', 'РН-Ванкор ООО'];
        }
    }

    // Сохранение списка целевых контрагентов в localStorage
    saveTargetContractors(contractors) {
        this.TARGET_CONTRAGENTS = contractors.filter(c => c.trim() !== '');
        localStorage.setItem('targetContractors', JSON.stringify(this.TARGET_CONTRAGENTS));
        console.log('Сохранен список контрагентов:', this.TARGET_CONTRAGENTS);
    }

    // Получение списка целевых контрагентов
    getTargetContractors() {
        return [...this.TARGET_CONTRAGENTS];
    }

    // Добавление контрагента в список
    addTargetContractor(contractor) {
        if (contractor && contractor.trim() !== '' && !this.TARGET_CONTRAGENTS.includes(contractor.trim())) {
            this.TARGET_CONTRAGENTS.push(contractor.trim());
            this.saveTargetContractors(this.TARGET_CONTRAGENTS);
            return true;
        }
        return false;
    }

    // Удаление контрагента из списка
    removeTargetContractor(contractor) {
        const index = this.TARGET_CONTRAGENTS.indexOf(contractor);
        if (index !== -1) {
            this.TARGET_CONTRAGENTS.splice(index, 1);
            this.saveTargetContractors(this.TARGET_CONTRAGENTS);
            return true;
        }
        return false;
    }

    formatDate(date) {
        if (!date) return '';
        const d = new Date(date);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return year + '-' + month + '-' + day;
    }

    getTodayString() {
        return this.formatDate(this.currentDate);
    }

    clearData() {
        this.debtData = [];
        this.debtHeaders = [];
        this.debtFile = null;
        this.debtFileName = '';
        this.receiptsData = [];
        this.processedDocuments = [];
        this.stats = {
            totalDocuments: 0,
            foundDocuments: 0,
            updatedDocuments: 0,
            errors: []
        };
    }

    async loadDebtRegistryFile(file) {
        console.log('Загрузка файла реестра ДЗ:', file.name);
        try {
            const arrayBuffer = await this.readFileAsArrayBuffer(file);

            const workbook = XLSX.read(arrayBuffer, {
                type: 'array',
                cellDates: true,
                raw: true
            });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            this.debtData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: null,
                raw: true
            });

            this.debtHeaders = this.debtData[0] || [];
            this.debtFile = file;
            this.debtFileName = file.name;

            console.log('Файл загружен, строк:', this.debtData.length);

            return {
                success: true,
                message: 'Загружено ' + this.debtData.length + ' строк',
                data: this.debtData
            };
        } catch (error) {
            console.error('Ошибка загрузки:', error);
            return {
                success: false,
                message: 'Ошибка загрузки файла: ' + error.message
            };
        }
    }

    async loadReceiptsFile(file) {
        console.log('Загрузка файла поступлений:', file.name);
        try {
            const arrayBuffer = await this.readFileAsArrayBuffer(file);

            const workbook = XLSX.read(arrayBuffer, {
                type: 'array',
                cellDates: true,
                raw: true
            });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            const rows = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: null,
                raw: true
            });

            if (rows.length < 2) {
                throw new Error('Файл не содержит данных');
            }

            const headers = rows[0] || [];

            const docNameCol = this.findColumnIndex(headers, 'Документ реализации');
            const dateCol = this.findColumnIndex(headers, 'Оплата по подписанию');
            const amountCol = this.findColumnIndex(headers, 'Сумма');
            const kontragentCol = this.findColumnIndex(headers, 'Контрагент');

            console.log('Найденные колонки:', {
                документ: docNameCol !== -1 ? (docNameCol + 1) + ' (' + headers[docNameCol] + ')' : 'не найдена',
                дата: dateCol !== -1 ? (dateCol + 1) + ' (' + headers[dateCol] + ')' : 'не найдена',
                сумма: amountCol !== -1 ? (amountCol + 1) + ' (' + headers[amountCol] + ')' : 'не найдена',
                контрагент: kontragentCol !== -1 ? (kontragentCol + 1) + ' (' + headers[kontragentCol] + ')' : 'не найдена'
            });

            if (dateCol === -1) {
                return {
                    success: false,
                    message: 'Не найдена колонка "Оплата по подписанию" в файле поступлений'
                };
            }

            // Собираем записи с датами
            this.receiptsData = [];

            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row || row.length === 0) continue;

                let dateValue = null;
                if (dateCol !== -1 && row[dateCol]) {
                    dateValue = this.parseExcelDate(row[dateCol]);
                }

                // Если есть дата, сохраняем запись
                if (dateValue) {
                    let docName = '';
                    if (docNameCol !== -1 && row[docNameCol]) {
                        docName = String(row[docNameCol]).trim();
                    }

                    let amount = 0;
                    if (amountCol !== -1 && row[amountCol]) {
                        amount = this.parseExcelNumber(row[amountCol]);
                    }

                    let kontragent = '';
                    if (kontragentCol !== -1 && row[kontragentCol]) {
                        kontragent = String(row[kontragentCol]).trim();
                    }

                    if (docName) {
                        this.receiptsData.push({
                            documentName: docName,
                            expectedDate: dateValue,
                            amount: amount,
                            kontragent: kontragent
                        });
                        console.log('Найден документ: ' + docName + ', дата: ' + this.formatDate(dateValue) + ', контрагент: ' + kontragent);
                    }
                }
            }

            console.log('Всего найдено записей с датами:', this.receiptsData.length);

            return {
                success: true,
                message: 'Загружено ' + this.receiptsData.length + ' записей с датами',
                data: this.receiptsData
            };
        } catch (error) {
            console.error('Ошибка загрузки:', error);
            return {
                success: false,
                message: 'Ошибка загрузки файла: ' + error.message
            };
        }
    }

    findColumnIndex(headers, searchText) {
        if (!headers || headers.length === 0) return -1;

        const searchLower = searchText.toLowerCase();
        for (let i = 0; i < headers.length; i++) {
            if (headers[i] && String(headers[i]).toLowerCase().indexOf(searchLower) !== -1) {
                return i;
            }
        }
        return -1;
    }

    readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                resolve(e.target.result);
            };
            reader.onerror = function() {
                reject(new Error('Ошибка чтения файла'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    parseExcelDate(value) {
        if (!value) return null;

        if (value instanceof Date) {
            return value;
        }

        if (typeof value === 'number') {
            return new Date((value - 25569) * 86400 * 1000);
        }

        if (typeof value === 'string') {
            const trimmed = value.trim();

            const yyyymmdd = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
            if (yyyymmdd) {
                return new Date(parseInt(yyyymmdd[1]), parseInt(yyyymmdd[2]) - 1, parseInt(yyyymmdd[3]));
            }

            const ddmmyyyy = trimmed.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
            if (ddmmyyyy) {
                return new Date(parseInt(ddmmyyyy[3]), parseInt(ddmmyyyy[2]) - 1, parseInt(ddmmyyyy[1]));
            }

            const date = new Date(trimmed);
            if (!isNaN(date.getTime())) {
                return date;
            }
        }

        return null;
    }

    parseExcelNumber(value) {
        if (value === undefined || value === null) return 0;
        if (typeof value === 'number') {
            return value;
        }
        if (typeof value === 'string') {
            const cleaned = value.replace(/\s/g, '').replace(',', '.');
            const num = parseFloat(cleaned);
            return isNaN(num) ? 0 : num;
        }
        return 0;
    }

    isDocumentRow(row) {
        if (!row || row.length === 0) return false;
        const value = row[this.COLUMNS.DOCUMENT_NAME];
        if (!value) return false;

        const str = String(value);
        return str.indexOf('Акт') !== -1 ||
               str.indexOf('Реализация') !== -1 ||
               str.indexOf('Корректировка') !== -1 ||
               str.indexOf('Поступление') !== -1;
    }

    // Находит контрагента для строки документа
    findKontragentForRow(rowIndex) {
        for (let i = rowIndex - 1; i >= 14; i--) {  // начиная с 14 строки (после заголовков)
            const row = this.debtData[i];
            if (!row) continue;

            const cellValue = row[0]; // колонка A
            if (!cellValue) continue;

            const strVal = String(cellValue).trim();
            // Проверяем, является ли строка контрагентом (ООО или АО, но не ДТ)
            if ((strVal.includes('ООО') || strVal.includes('АО')) && !strVal.startsWith('ДТ ')) {
                return strVal;
            }
        }
        return null;
    }

    reconcile() {
        console.log('Начало сверки...');
        this.stats = {
            totalDocuments: 0,
            foundDocuments: 0,
            updatedDocuments: 0,
            errors: []
        };

        if (this.debtData.length === 0) {
            return {
                success: false,
                message: 'Не загружен реестр ДЗ'
            };
        }

        // Создаем карту документов из файла 2 (только те, у которых есть дата)
        const receiptsMap = new Map();
        this.receiptsData.forEach(function(item) {
            receiptsMap.set(item.documentName, item);
        });

        console.log('Создана карта документов из файла поступлений, размер:', receiptsMap.size);
        console.log('Целевые контрагенты:', this.TARGET_CONTRAGENTS);

        const today = this.currentDate;
        const self = this;

        // Проходим по всем строкам файла 1 и ищем документы
        for (let i = 0; i < this.debtData.length; i++) {
            const row = this.debtData[i];
            if (!row || row.length === 0) continue;

            if (this.isDocumentRow(row)) {
                this.stats.totalDocuments++;

                const docName = String(row[this.COLUMNS.DOCUMENT_NAME] || '').trim();

                // Находим контрагента для этого документа
                const kontragent = this.findKontragentForRow(i);

                // Проверяем, входит ли контрагент в целевой список
                const isTargetKontragent = kontragent && this.TARGET_CONTRAGENTS.some(target =>
                    kontragent.includes(target)
                );

                // Если контрагент не в списке целевых - пропускаем документ
                if (!isTargetKontragent) {
                    console.log(`Пропущен (контрагент ${kontragent} не в списке): ${docName}`);
                    continue;
                }

                console.log(`\nОбработка документа: ${docName}`);
                console.log(`  Контрагент: ${kontragent}`);

                // Ищем документ в карте поступлений
                const receiptItem = receiptsMap.get(docName);

                let expectedDate = null;
                let hasDate = false;

                if (receiptItem) {
                    expectedDate = receiptItem.expectedDate;
                    hasDate = true;
                    console.log(`  Найден в файле поступлений с датой: ${this.formatDate(expectedDate)}`);
                } else {
                    console.log(`  Не найден в файле поступлений - будет считаться непросроченным`);
                }

                this.stats.foundDocuments++;

                const debtAmount = this.parseExcelNumber(row[this.COLUMNS.DEBT_AMOUNT] || 0);
                console.log(`  Сумма долга: ${debtAmount}`);

                if (debtAmount > 0) {
                    // Всегда обрабатываем документ, даже если нет даты
                    const updated = this.updateDocumentRow(i, debtAmount, expectedDate, today, hasDate);
                    if (updated) {
                        this.stats.updatedDocuments++;
                        this.processedDocuments.push({
                            documentName: docName,
                            action: 'updated',
                            date: expectedDate ? this.formatDate(expectedDate) : null,
                            amount: debtAmount,
                            rowIndex: i,
                            rowNumber: i + 1
                        });
                        console.log(`  Документ добавлен в processedDocuments`);
                    }
                }
            }
        }

        console.log('\nСверка завершена. Найдено документов целевых контрагентов:', this.stats.foundDocuments, 'Обновлено:', this.stats.updatedDocuments);
        console.log('processedDocuments содержит', this.processedDocuments.length, 'записей');

        return {
            success: true,
            message: 'Сверка завершена. Найдено документов: ' + this.stats.foundDocuments + ', обновлено: ' + this.stats.updatedDocuments,
            stats: this.stats
        };
    }

    updateDocumentRow(rowIndex, debtAmount, expectedDate, today, hasDate) {
        const row = this.debtData[rowIndex];
        if (!row) return false;

        let changed = false;

        console.log(`  Обновление строки ${rowIndex}: hasDate=${hasDate}, expectedDate=${expectedDate ? this.formatDate(expectedDate) : 'null'}`);

        // Очищаем все интервалы
        const intervalCols = [
            this.COLUMNS.NOT_OVERDUE,      // T
            this.COLUMNS.INTERVAL_1_15,    // U
            this.COLUMNS.INTERVAL_16_29,   // V
            this.COLUMNS.INTERVAL_30_89,   // W
            this.COLUMNS.INTERVAL_90_179,  // X
            this.COLUMNS.INTERVAL_180_PLUS // Y
        ];

        // Очищаем интервалы
        for (let j = 0; j < intervalCols.length; j++) {
            const col = intervalCols[j];
            if (row[col] !== 0) {
                row[col] = 0;
                changed = true;
            }
        }

        // Если даты нет - документ непросроченный
        if (!hasDate || expectedDate === null || expectedDate >= today) {
            console.log(`  -> НЕ ПРОСРОЧЕНО (причина: ${!hasDate ? 'нет даты' : 'дата в будущем'})`);

            // O (просрочено) - очищаем
            if (row[this.COLUMNS.OVERDUE] !== 0) {
                row[this.COLUMNS.OVERDUE] = 0;
                changed = true;
            }

            // R (дни) - очищаем
            if (row[this.COLUMNS.DAYS] !== 0) {
                row[this.COLUMNS.DAYS] = 0;
                changed = true;
            }

            // T (не просрочено) - устанавливаем сумму
            if (row[this.COLUMNS.NOT_OVERDUE] !== debtAmount) {
                row[this.COLUMNS.NOT_OVERDUE] = debtAmount;
                changed = true;
            }

        } else if (expectedDate < today) {
            // ПРОСРОЧЕНО
            const daysOverdue = Math.floor((today - expectedDate) / (1000 * 60 * 60 * 24));
            console.log(`  -> ПРОСРОЧЕНО на ${daysOverdue} дн.`);

            // O (просрочено) - сумма долга
            if (row[this.COLUMNS.OVERDUE] !== debtAmount) {
                row[this.COLUMNS.OVERDUE] = debtAmount;
                changed = true;
            }

            // R (дни просрочки)
            if (row[this.COLUMNS.DAYS] !== daysOverdue) {
                row[this.COLUMNS.DAYS] = daysOverdue;
                changed = true;
            }

            // T (не просрочено) - очищаем
            if (row[this.COLUMNS.NOT_OVERDUE] !== 0) {
                row[this.COLUMNS.NOT_OVERDUE] = 0;
                changed = true;
            }

            // Определяем интервал по дням
            let intervalCol = this.COLUMNS.INTERVAL_1_15; // U по умолчанию
            if (daysOverdue >= 1 && daysOverdue <= 15) {
                intervalCol = this.COLUMNS.INTERVAL_1_15;      // U
            } else if (daysOverdue >= 16 && daysOverdue <= 29) {
                intervalCol = this.COLUMNS.INTERVAL_16_29;     // V
            } else if (daysOverdue >= 30 && daysOverdue <= 89) {
                intervalCol = this.COLUMNS.INTERVAL_30_89;     // W
            } else if (daysOverdue >= 90 && daysOverdue <= 179) {
                intervalCol = this.COLUMNS.INTERVAL_90_179;    // X
            } else if (daysOverdue >= 180) {
                intervalCol = this.COLUMNS.INTERVAL_180_PLUS;  // Y
            }

            if (row[intervalCol] !== debtAmount) {
                row[intervalCol] = debtAmount;
                changed = true;
            }
        }

        return changed;
    }

    async exportToExcel() {
        console.log('=== ОТПРАВКА НА СЕРВЕР ===');
        console.log('Количество документов для отправки:', this.processedDocuments.length);

        if (!this.debtFile) {
            console.error('ОШИБКА: файл не загружен');
            return { success: false, message: 'Нет данных для экспорта' };
        }

        if (this.processedDocuments.length === 0) {
            console.log('Нет обновленных документов');
            return { success: false, message: 'Нет обновленных документов для целевых контрагентов' };
        }

        try {
            const formData = new FormData();
            formData.append('file', this.debtFile);
            formData.append('data', JSON.stringify({
                updatedDocuments: this.processedDocuments
            }));

            console.log('Отправляем на сервер...');

            const serverResponse = await fetch('http://localhost:5000/save-excel', {
                method: 'POST',
                body: formData
            });

            if (!serverResponse.ok) {
                const error = await serverResponse.json();
                throw new Error(error.error || 'Ошибка сервера');
            }

            const blob = await serverResponse.blob();

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'ДЗ_обновленный_' + this.formatDate(this.currentDate) + '.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            console.log('Файл успешно сохранен');

            return {
                success: true,
                message: 'Файл сохранен с полным форматированием'
            };

        } catch (error) {
            console.error('Ошибка:', error);
            return {
                success: false,
                message: 'Ошибка при сохранении: ' + error.message
            };
        }
    }

    getStats() {
        return {
            totalDocuments: this.stats.totalDocuments,
            foundDocuments: this.stats.foundDocuments,
            updatedDocuments: this.stats.updatedDocuments,
            errors: this.stats.errors,
            debtRows: this.debtData.length,
            receiptsWithDates: this.receiptsData.length,
            processedCount: this.processedDocuments.length
        };
    }

    getProcessedLog() {
        return this.processedDocuments;
    }

    // Сохранение оригинального файла для тестирования
    async saveOriginalForTest() {
        console.log('=== ТЕСТОВОЕ СОХРАНЕНИЕ ОРИГИНАЛА ===');
        
        if (!this.debtFile) {
            console.error('ОШИБКА: файл не загружен');
            return { 
                success: false, 
                message: 'Сначала загрузите файл реестра ДЗ' 
            };
        }

        try {
            // Просто сохраняем оригинальный файл с новым именем
            const dateStr = this.formatDate(this.currentDate);
            const fileName = `ДЗ_оригинал_${dateStr}.xlsx`;
            
            // Создаем Blob из оригинального файла
            const blob = new Blob([await this.debtFile.arrayBuffer()], 
                { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            
            // Сохраняем через FileSaver
            saveAs(blob, fileName);
            
            console.log('Оригинальный файл сохранен:', fileName);
            
            return {
                success: true,
                message: `Оригинальный файл сохранен как ${fileName}`
            };
            
        } catch (error) {
            console.error('Ошибка при сохранении оригинала:', error);
            return {
                success: false,
                message: 'Ошибка при сохранении: ' + error.message
            };
        }
    }
}
