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

        // Данные для сводных таблиц
        this.siUatFile = null;         // файл СИ УАТ
        this.siUatFileName = '';       // имя файла СИ УАТ
        this.summaryDT = {             // свод задолженности ДТ
            legal: 0,
            notRecoverable: 0,
            recoverable: 0
        };
        this.summarySIUAT = {          // свод задолженности СИ УАТ
            totalDebt: 0,
            totalOverdue: 0,
            legal: 0,
            notRecoverable: 0,
            recoverable: 0
        };
        this.currentSubdivisionData = {}; // данные текущего дня по филиалам

        // Загружаем список целевых контрагентов из localStorage или используем стандартный
        this.loadTargetContractors();

        // Загружаем сохранённые сводные данные
        this.loadSummaryData();

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

    // Загрузка сводных данных из localStorage
    loadSummaryData() {
        try {
            const dtData = localStorage.getItem('summaryDT');
            if (dtData) {
                this.summaryDT = JSON.parse(dtData);
            }
            const siuatData = localStorage.getItem('summarySIUAT');
            if (siuatData) {
                const parsed = JSON.parse(siuatData);
                // Объединяем с дефолтными значениями для новых полей
                this.summarySIUAT = {
                    totalDebt: parsed.totalDebt || 0,
                    totalOverdue: parsed.totalOverdue || 0,
                    legal: parsed.legal || 0,
                    notRecoverable: parsed.notRecoverable || 0,
                    recoverable: parsed.recoverable || 0
                };
            }
        } catch (e) {
            console.error('Ошибка загрузки сводных данных:', e);
        }
    }

    // Сохранение сводных данных в localStorage
    saveSummaryData() {
        try {
            localStorage.setItem('summaryDT', JSON.stringify(this.summaryDT));
            localStorage.setItem('summarySIUAT', JSON.stringify(this.summarySIUAT));
            console.log('Сводные данные сохранены');
        } catch (e) {
            console.error('Ошибка сохранения сводных данных:', e);
        }
    }

    // Получить данные предыдущего дня из localStorage (единый ключ для ручных данных)
    getPreviousDayData() {
        try {
            const data = localStorage.getItem('previousDayDebt_manual');
            return data ? JSON.parse(data) : {};
        } catch (e) {
            console.error('Ошибка загрузки данных предыдущего дня:', e);
            return {};
        }
    }

    // Сохранить данные текущего дня в localStorage (для использования завтра — единый ключ)
    saveCurrentDayData() {
        try {
            localStorage.setItem('previousDayDebt_manual', JSON.stringify(this.currentSubdivisionData));
            console.log('Данные текущего дня сохранены для использования завтра');
        } catch (e) {
            console.error('Ошибка сохранения данных текущего дня:', e);
        }
    }

    // Собрать данные по филиалам из debtData (только из колонки OVERDUE)
    collectSubdivisionData() {
        const subdivisionData = {};
        let currentFilial = null;
        let filialCount = 0;
        let docCount = 0;
        let totalOverdue = 0;

        console.log('=== collectSubdivisionData START ===');
        console.log('debtData строк:', this.debtData.length);
        console.log('processedDocuments (обновлённые через reconcile):', this.processedDocuments.length);

        // Создаём Set обработанных документов для отладки
        const processedRowSet = new Set(this.processedDocuments.map(d => d.rowIndex));

        // Отслеживаем уже обработанные строки чтобы избежать дублирования
        const processedRows = new Set();

        for (let i = 0; i < this.debtData.length; i++) {
            const row = this.debtData[i];
            if (!row || row.length === 0) continue;

            const cellValue = row[0];
            if (!cellValue) continue;

            const strVal = String(cellValue).trim();

            // Филиал — строка начинается с "ДТ "
            if (strVal.startsWith('ДТ ')) {
                currentFilial = strVal;
                if (!subdivisionData[currentFilial]) {
                    subdivisionData[currentFilial] = 0;
                    filialCount++;
                }
            }

            // Документ — добавляем просрочку к текущему филиалу
            // ВАЖНО: Пропускаем строки контрагентов, договоров и итого
            if (this.isDocumentRow(row) && currentFilial && !processedRows.has(i)) {
                // Проверяем, что это не строка контрагента или договора
                const isKontragentOrDogovor = strVal.includes('Договор') || 
                                              strVal.includes('договор') ||
                                              (!strVal.includes('Акт') && !strVal.includes('Реализация') && 
                                               !strVal.includes('Корректировка') && !strVal.includes('Поступление') &&
                                               !strVal.includes('Взаимозачет') && !strVal.includes('Взаимозачёт') &&
                                               !strVal.includes('Списание') && !strVal.includes('УПД'));
                
                // Пропускаем если это не документ
                if (isKontragentOrDogovor) {
                    continue;
                }
                
                // ВАЖНО: берем значение из колонки OVERDUE (просрочено)
                const rawValue = row[this.COLUMNS.OVERDUE];
                const overdue = this.parseExcelNumber(rawValue || 0);
                
                // Добавляем только если строка ещё не была обработана
                subdivisionData[currentFilial] += overdue;
                totalOverdue += overdue;
                docCount++;
                processedRows.add(i);  // Помечаем как обработанную

                // Логируем первые 5 документов для отладки
                if (docCount <= 5) {
                    const source = processedRowSet.has(i) ? 'RECONCILED' : 'ORIGINAL';
                    console.log(`  Дока #${docCount} [${source}]: ${strVal.substring(0, 40)}... | raw=${rawValue} | overdue=${overdue} | филиал=${currentFilial}`);
                }
            }
        }

        // Округляем до 2 знаков
        for (const key in subdivisionData) {
            subdivisionData[key] = Math.round(subdivisionData[key] * 100) / 100;
        }

        this.currentSubdivisionData = subdivisionData;
        console.log('collectSubdivisionData: филиалов=' + filialCount + ', документов=' + docCount);
        console.log('collectSubdivisionData: общая просрочка=' + totalOverdue);
        console.log('collectSubdivisionData: данные по подразделениям:', JSON.stringify(subdivisionData));
        console.log('=== collectSubdivisionData END ===');

        return subdivisionData;
    }

    // Загрузка файла СИ УАТ
    async loadSiUatFile(file) {
        console.log('Загрузка файла СИ УАТ:', file.name);
        try {
            this.siUatFile = file;
            this.siUatFileName = file.name;
            return {
                success: true,
                message: 'Файл СИ УАТ загружен: ' + file.name
            };
        } catch (error) {
            console.error('Ошибка загрузки файла СИ УАТ:', error);
            return {
                success: false,
                message: 'Ошибка загрузки файла СИ УАТ: ' + error.message
            };
        }
    }

    // Загрузка данных предыдущего дня из Excel файла (полная замена данных)
    async loadPreviousDayDataFromFile(file) {
        console.log('Загрузка данных предыдущего дня из файла:', file.name);
        try {
            const arrayBuffer = await this.readFileAsArrayBuffer(file);
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, raw: true });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: true });

            if (rows.length < 2) {
                return { success: false, message: 'Файл не содержит данных (минимум 2 строки: заголовок + данные)' };
            }

            const headers = rows[0] || [];
            console.log('Заголовки:', headers);

            // Ищем колонки "Подразделение" и "Сумма ПДЗ"
            const subdivisionCol = this.findColumnIndex(headers, 'Подразделение');
            const amountCol = this.findColumnIndex(headers, 'Сумма ПДЗ');

            if (subdivisionCol === -1) {
                return { success: false, message: 'Не найдена колонка "Подразделение". Заголовки: ' + headers.join(', ') };
            }
            if (amountCol === -1) {
                return { success: false, message: 'Не найдена колонка "Сумма ПДЗ". Заголовки: ' + headers.join(', ') };
            }

            console.log(`Найдены колонки: Подразделение=${subdivisionCol + 1}, Сумма ПДЗ=${amountCol + 1}`);

            // Полная замена данных — создаём новый объект
            const previousDayData = {};
            let parsedCount = 0;
            let totalAmount = 0;

            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row || row.length === 0) continue;

                const subdivision = row[subdivisionCol];
                const amount = row[amountCol];

                if (subdivision && amount !== undefined && amount !== null) {
                    const subdivisionName = String(subdivision).trim();
                    const amountValue = this.parseExcelNumber(amount);

                    // Сохраняем все подразделения, включая с нулевой суммой
                    if (subdivisionName) {
                        previousDayData[subdivisionName] = amountValue;
                        parsedCount++;
                        totalAmount += amountValue;
                    }
                }
            }

            if (parsedCount === 0) {
                return { success: false, message: 'Не найдено данных для загрузки' };
            }

            // Полная замена: удаляем старые данные и записываем новые
            try {
                localStorage.setItem('previousDayDebt_manual', JSON.stringify(previousDayData));
                console.log(`Данные предыдущего дня полностью заменены: ${parsedCount} записей, общая сумма: ${totalAmount.toFixed(2)}`);
            } catch (e) {
                console.error('Ошибка сохранения в localStorage:', e);
                return { success: false, message: 'Ошибка сохранения данных: ' + e.message };
            }

            // Обновляем currentSubdivisionData для отображения в таблице
            this.currentSubdivisionData = previousDayData;

            console.log(`Загружено ${parsedCount} записей:`, previousDayData);

            return {
                success: true,
                message: `Загружено ${parsedCount} подразделений. Общая сумма: ${totalAmount.toFixed(2)}`,
                data: previousDayData,
                count: parsedCount,
                total: totalAmount
            };
        } catch (error) {
            console.error('Ошибка загрузки файла:', error);
            return { success: false, message: 'Ошибка загрузки файла: ' + error.message };
        }
    }

    // Очистка данных предыдущего дня
    clearPreviousDayData() {
        try {
            localStorage.removeItem('previousDayDebt_manual');
            this.currentSubdivisionData = {};
            console.log('Данные предыдущего дня очищены');
            return { success: true, message: 'Данные очищены' };
        } catch (e) {
            console.error('Ошибка очистки данных:', e);
            return { success: false, message: 'Ошибка очистки: ' + e.message };
        }
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
        this.siUatFile = null;
        this.siUatFileName = '';
        this.currentSubdivisionData = {};
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

            // Сбрасываем данные предыдущей сессии, чтобы не использовать старые значения
            this.currentSubdivisionData = {};
            this.processedDocuments = [];

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
        
        // Расширенный список типов документов
        const documentKeywords = [
            'Акт', 'Реализация', 'Корректировка', 'Поступление',
            'Взаимозачет', 'Взаимозачёт', 'Списание', 'УПД', 'Счет-фактура',
            'Товарная накладная', 'ТОРГ-12', 'Универсальный передаточный'
        ];
        
        return documentKeywords.some(keyword => str.includes(keyword));
    }

    // Находит контрагента для строки документа (расширенная версия)
    findKontragentForRow(rowIndex) {
        for (let i = rowIndex - 1; i >= 14; i--) {
            const row = this.debtData[i];
            if (!row) continue;
            const cellValue = row[0];
            if (!cellValue) continue;
            const strVal = String(cellValue).trim();
            
            // Проверяем, является ли строка филиалом
            if (strVal.startsWith('ДТ ')) {
                return null;  // дошли до филиала - контрагент не найден
            }
            
            // Проверяем, является ли строка контрагентом
            // Теперь считаем контрагентом любую непустую строку, которая:
            // - не начинается с ДТ
            // - не содержит "Договор"
            // - не содержит слова из списка документов
            // - имеет длину > 2 символов
            if (strVal.length > 2 && 
                !strVal.includes('Договор') && 
                !strVal.includes('Акт') && 
                !strVal.includes('Реализация') && 
                !strVal.includes('Корректировка') && 
                !strVal.includes('Поступление') && 
                !strVal.startsWith('ДТ ')) {
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
                
                // Проверяем, является ли контрагент целевым
                // Если контрагент не найден (например, документ висит прямо на филиале) - считаем его целевым?
                // По логике, если контрагент не определен, документ обрабатывать не нужно
                if (!kontragent) {
                    console.log(`Пропущен (контрагент не найден): ${docName}`);
                    continue;
                }

                // Проверяем, входит ли контрагент в целевой список
                const isTargetKontragent = this.TARGET_CONTRAGENTS.some(target =>
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
                    console.log(`  Не найден в файле поступлений - документ останется без изменений`);
                }

                this.stats.foundDocuments++;

                const debtAmount = this.parseExcelNumber(row[this.COLUMNS.DEBT_AMOUNT] || 0);
                console.log(`  Сумма долга: ${debtAmount}`);

                // ОБНОВЛЯЕМ ТОЛЬКО ЕСЛИ ЕСТЬ ДАТА В ФАЙЛЕ ПОСТУПЛЕНИЙ
                if (debtAmount > 0 && hasDate) {
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
                } else {
                    console.log(`  Документ пропущен (нет даты в файле поступлений)`);
                }
            }
        }

        console.log('\nСверка завершена. Найдено документов целевых контрагентов:', this.stats.foundDocuments, 'Обновлено:', this.stats.updatedDocuments);
        console.log('processedDocuments содержит', this.processedDocuments.length, 'записей');

        // ВАЖНО: Принудительно пересобираем данные по филиалам из ОБНОВЛЕННОГО debtData
        this.collectSubdivisionData();

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

        // ВАЖНО: Принудительно пересобираем данные по филиалам из ТЕКУЩЕГО debtData
        // Это гарантирует, что currentSubdivisionData будет содержать актуальные данные из файла
        console.log('ПРИНУДИТЕЛЬНЫЙ ПЕРЕСБОР данных по филиалам из debtData...');
        this.collectSubdivisionData();

        // Проверяем, что данные собраны
        if (Object.keys(this.currentSubdivisionData).length === 0) {
            console.error('ОШИБКА: Не удалось собрать данные по подразделениям. Убедитесь, что файл содержит филиалы (ДТ ...) и документы.');
            return { success: false, message: 'Нет данных по подразделениям. Проверьте структуру файла.' };
        }

        console.log('=== ДАННЫЕ ДЛЯ ОТПРАВКИ НА СЕРВЕР ===');
        console.log('currentDayData (из debtData, колонка O):', JSON.stringify(this.currentSubdivisionData));

        // Получаем данные предыдущего дня из localStorage
        const previousDayData = this.getPreviousDayData();
        console.log('previousDayData (из localStorage):', JSON.stringify(previousDayData));

        // Рассчитываем общие суммы для сводки ДТ
        let totalDebt = 0;
        let totalOverdue = 0;
        for (let i = 0; i < this.debtData.length; i++) {
            const row = this.debtData[i];
            if (!row) continue;
            if (this.isDocumentRow(row)) {
                totalDebt += this.parseExcelNumber(row[this.COLUMNS.DEBT_AMOUNT] || 0);
                totalOverdue += this.parseExcelNumber(row[this.COLUMNS.OVERDUE] || 0);
            }
        }
        totalDebt = Math.round(totalDebt * 100) / 100;
        totalOverdue = Math.round(totalOverdue * 100) / 100;

        try {
            const formData = new FormData();
            formData.append('file', this.debtFile);

            // Формируем объект данных для сводных таблиц
            const summaryData = {
                updatedDocuments: this.processedDocuments,
                // Данные для таблицы динамики
                previousDayData: previousDayData,
                currentDayData: this.currentSubdivisionData,  // ВАЖНО: это данные из debtData, а не из localStorage
                currentDate: this.formatDate(this.currentDate),
                previousDate: 'предыдущий рабочий день',
                // Свод задолженности ДТ
                summaryDT: {
                    totalDebt: totalDebt,
                    totalOverdue: totalOverdue,
                    legal: this.summaryDT.legal,
                    notRecoverable: this.summaryDT.notRecoverable,
                    recoverable: this.summaryDT.recoverable
                },
                // Свод задолженности СИ УАТ
                summarySIUAT: {
                    totalDebt: this.summarySIUAT.totalDebt || 0,
                    totalOverdue: this.summarySIUAT.totalOverdue || 0,
                    legal: this.summarySIUAT.legal,
                    notRecoverable: this.summarySIUAT.notRecoverable,
                    recoverable: this.summarySIUAT.recoverable
                },
                // Метаданные файла СИ УАТ
                siUatFileName: this.siUatFileName || ''
            };

            // Если файл СИ УАТ загружен — добавляем его
            if (this.siUatFile) {
                formData.append('siUatFile', this.siUatFile);
            }

            formData.append('data', JSON.stringify(summaryData));

            console.log('Отправляем на сервер...');
            console.log('Размер данных:', JSON.stringify(summaryData).length, 'байт');

            // Увеличиваем таймаут и добавляем обработку ошибок
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 300000); // 5 минут таймаут

            const serverResponse = await fetch('http://31.130.155.16:5000/save-excel', {
                method: 'POST',
                body: formData,
                signal: controller.signal
            }).finally(() => clearTimeout(timeoutId));

            if (!serverResponse.ok) {
                let errorMessage = 'Ошибка сервера';
                try {
                    const errorData = await serverResponse.json();
                    errorMessage = errorData.error || errorMessage;
                } catch (e) {
                    errorMessage = `HTTP ${serverResponse.status}: ${serverResponse.statusText}`;
                }
                throw new Error(errorMessage);
            }

            const blob = await serverResponse.blob();
            console.log('Получен ответ, размер:', blob.size, 'байт');

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'ДЗ_обновленный_' + this.formatDate(this.currentDate) + '.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            console.log('Файл успешно сохранен');
            console.log('=== ИТОГОВЫЕ ДАННЫЕ ДЛЯ ТАБЛИЦЫ ДИНАМИКИ ===');
            console.log('СТОЛБЕЦ 2 (текущий день, из debtData):', JSON.stringify(this.currentSubdivisionData));
            console.log('СТОЛБЕЦ 3 (предыдущий день, из localStorage):', JSON.stringify(this.getPreviousDayData()));
            console.log('Данные НЕ сохранены в localStorage автоматически.');
            console.log('Для сохранения текущих данных в localStorage используйте "Настройки сводных" → "Сохранить".');

            return {
                success: true,
                message: 'Файл сохранен с полным форматированием'
            };

        } catch (error) {
            console.error('Ошибка при отправке на сервер:', error);
            
            if (error.name === 'AbortError') {
                return {
                    success: false,
                    message: 'Превышено время ожидания ответа от сервера. Попробуйте уменьшить количество документов.'
                };
            }
            
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