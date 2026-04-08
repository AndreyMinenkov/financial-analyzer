// supplier-payments.js — Модуль "Оплаты Поставщикам"
class SupplierPaymentsManager {
    constructor(contractorsLibrary) {
        this.library = contractorsLibrary;
        this.loadedData = [];       // [{sheetName, headers, rows}]
        this.pivotTables = [];      // [{sheetName, pivotData, headers}]
    }

    // ===== ЗАГРУЗКА ФАЙЛА =====
    async loadExcelFile(file) {
        if (!file) return { success: false, message: 'Файл не выбран' };

        try {
            const workbook = await this.readExcelFile(file);
            this.loadedData = [];
            this.pivotTables = [];

            // Обрабатываем каждую вкладку
            for (const sheetName of workbook.SheetNames) {
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

                if (rows.length < 2) continue; // Пропускаем пустые листы

                const headers = rows[0].map(h => String(h || '').trim());
                const dataRows = rows.slice(1).filter(row => row.some(cell => cell !== null && cell !== ''));

                // Определяем колонки
                const columns = this.detectColumns(headers);
                if (!columns.payer) {
                    console.warn(`Лист "${sheetName}": не найдена колонка "Получатель", пропускаем`);
                    continue;
                }

                this.loadedData.push({
                    sheetName,
                    headers,
                    rows: dataRows,
                    columns
                });

                // Формируем сводную таблицу
                const pivot = this.buildPivotTable(dataRows, columns);
                this.pivotTables.push({
                    sheetName,
                    pivotData: pivot.data,
                    pivotHeaders: pivot.headers,
                    explanations: pivot.explanations
                });
            }

            return {
                success: true,
                message: `Загружено ${this.loadedData.length} реестров`,
                count: this.loadedData.length
            };
        } catch (error) {
            console.error('Ошибка загрузки файла:', error);
            return { success: false, message: 'Ошибка загрузки: ' + error.message };
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    resolve(workbook);
                } catch (error) {
                    reject(new Error('Ошибка чтения Excel: ' + error.message));
                }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    // ===== ОПРЕДЕЛЕНИЕ КОЛОНОК =====
    detectColumns(headers) {
        const mapping = {
            payer: ['получатель', 'контрагент', 'плательщик'],
            amount: ['сумма', 'сумма заявки', 'сумма оплаты'],
            subdivision: ['подразделение', 'дт ', 'филиал'],
            purpose: ['назначение платежа', 'назначение', 'комментарий'],
            status: ['статус', 'состояние'],
            date: ['дата', 'дата заявки'],
            article: ['статья ддс', 'статья', 'статья платежей'],
            priority: ['приоритет'],
            organization: ['организация'],
            applicant: ['заявитель'],
            application: ['заявка', 'номер заявки']
        };

        const result = {};

        headers.forEach((header, index) => {
            const lowerHeader = header.toLowerCase();

            for (const [key, patterns] of Object.entries(mapping)) {
                for (const pattern of patterns) {
                    if (lowerHeader.includes(pattern)) {
                        if (!result[key]) {
                            result[key] = index;
                        }
                        break;
                    }
                }
            }
        });

        return result;
    }

    // ===== ПОСТРОЕНИЕ СВОДНОЙ ТАБЛИЦЫ =====
    buildPivotTable(rows, columns) {
        const contractorMap = new Map(); // contractor -> { subdivision -> sum }
        const subdivisions = new Set();
        const explanations = new Map(); // contractor -> explanation text

        rows.forEach(row => {
            const payer = this.cleanString(row[columns.payer]);
            if (!payer) return;

            const amount = this.parseAmount(row[columns.amount]);
            const subdivision = columns.subdivision !== undefined
                ? this.cleanString(row[columns.subdivision])
                : 'Общий';

            if (subdivision) subdivisions.add(subdivision);

            if (!contractorMap.has(payer)) {
                contractorMap.set(payer, {});
            }

            const contractorData = contractorMap.get(payer);
            contractorData[subdivision] = (contractorData[subdivision] || 0) + amount;

            // Сохраняем назначение платежа для пояснений
            if (columns.purpose !== undefined && !explanations.has(payer)) {
                const purpose = this.cleanString(row[columns.purpose]);
                if (purpose) {
                    explanations.set(payer, purpose);
                }
            }
        });

        // Сортируем подразделения
        const sortedSubdivisions = Array.from(subdivisions).sort();

        // Формируем данные сводной таблицы
        const pivotData = [];
        const sortedContractors = Array.from(contractorMap.keys()).sort();

        sortedContractors.forEach(contractor => {
            const data = contractorMap.get(contractor);
            let total = 0;

            const row = { contractor };
            sortedSubdivisions.forEach(sub => {
                const value = data[sub] || 0;
                row[sub] = value;
                total += value;
            });
            row.total = total;

            // Определяем пояснение
            row.explanation = this.getExplanation(contractor, explanations.get(contractor) || '');

            pivotData.push(row);
        });

        return {
            data: pivotData,
            headers: sortedSubdivisions,
            explanations
        };
    }

    // ===== БИБЛИОТЕКА КОНТРАГЕНТОВ =====
    getExplanation(contractorName, fallbackPurpose) {
        // Ищем в библиотеке
        const libraryEntry = this.library.findByContractor(contractorName);
        if (libraryEntry) {
            return libraryEntry.explanation;
        }

        // Не нашли — используем назначение платежа и добавляем в библиотеку
        if (fallbackPurpose) {
            this.library.addContractor({
                name: contractorName,
                organization: '',
                explanation: fallbackPurpose
            });
            return fallbackPurpose;
        }

        return '';
    }

    // ===== ЭКСПОРТ В EXCEL (через сервер для сохранения форматирования) =====
    async exportToExcel(originalFile) {
        if (this.pivotTables.length === 0) {
            return { success: false, message: 'Нет данных для экспорта' };
        }

        if (!originalFile) {
            return { success: false, message: 'Нет оригинального файла' };
        }

        try {
            // Подготавливаем данные для отправки на сервер
            const pivotTablesData = this.pivotTables.map(pivot => ({
                sheetName: pivot.sheetName,
                headers: pivot.pivotHeaders,
                data: pivot.pivotData.map(row => ({
                    contractor: row.contractor,
                    total: row.total,
                    explanation: row.explanation || '',
                    ...Object.fromEntries(pivot.pivotHeaders.map(h => [h, row[h] || 0]))
                }))
            }));

            const formData = new FormData();
            formData.append('file', originalFile);
            formData.append('data', JSON.stringify({
                pivotTables: pivotTablesData
            }));

            console.log('Отправка на сервер для обработки...');
            console.log('Сводных таблиц:', pivotTablesData.length);

            // Отправляем на сервер
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 300000);

            const response = await fetch('http://31.130.155.16:5000/save-suppliers', {
                method: 'POST',
                body: formData,
                signal: controller.signal
            }).finally(() => clearTimeout(timeoutId));

            if (!response.ok) {
                let errorMessage = 'Ошибка сервера';
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.error || errorMessage;
                } catch (e) {
                    errorMessage = `HTTP ${response.status}: ${response.statusText}`;
                }
                throw new Error(errorMessage);
            }

            const blob = await response.blob();
            console.log('Получен ответ, размер:', blob.size, 'байт');

            // Скачиваем файл
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const date = new Date().toISOString().split('T')[0];
            a.download = `Оплаты_поставщикам_${date}.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            console.log('Файл успешно сохранён');
            return { success: true, message: 'Файл успешно сохранён' };

        } catch (error) {
            console.error('Ошибка экспорта:', error);
            if (error.name === 'AbortError') {
                return {
                    success: false,
                    message: 'Превышено время ожидания ответа от сервера'
                };
            }
            return { success: false, message: 'Ошибка экспорта: ' + error.message };
        }
    }

    createPivotWorksheet(pivot) {
        const { pivotData, pivotHeaders } = pivot;

        // Заголовок
        const data = [
            ['Сводная таблица оплат по подразделениям'],
            []
        ];

        // Шапка таблицы
        const headerRow = ['Контрагент', ...pivotHeaders, 'Итого', 'Пояснение'];
        data.push(headerRow);

        // Данные
        pivotData.forEach(row => {
            const values = pivotHeaders.map(h => row[h] || 0);
            data.push([
                row.contractor,
                ...values,
                row.total,
                row.explanation
            ]);
        });

        // Итоговая строка
        const totalRow = ['ИТОГО'];
        pivotHeaders.forEach(h => {
            const sum = pivotData.reduce((acc, r) => acc + (r[h] || 0), 0);
            totalRow.push(sum);
        });
        const grandTotal = pivotData.reduce((acc, r) => acc + r.total, 0);
        totalRow.push(grandTotal);
        totalRow.push('');
        data.push(totalRow);

        const ws = XLSX.utils.aoa_to_sheet(data);

        // ===== ФОРМАТИРОВАНИЕ =====
        const range = XLSX.utils.decode_range(ws['!ref']);

        // Заголовки (строка 2, индекс 2) — тёмно-синий фон, белый текст
        const headerStyle = {
            fill: { fgColor: { rgb: '1F3864' } },
            font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
        };

        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 2, c: col });
            if (ws[cellAddress]) {
                ws[cellAddress].s = headerStyle;
            }
        }

        // Числовые ячейки — числовой формат
        for (let row = 3; row <= range.e.r; row++) {
            for (let col = 1; col <= pivotHeaders.length + 1; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                if (ws[cellAddress] && typeof ws[cellAddress].v === 'number') {
                    ws[cellAddress].t = 'n';
                    ws[cellAddress].z = '#,##0.00';
                }
            }
        }

        // Строка с пояснениями — жёлтый фон
        for (let row = 3; row < range.e.r; row++) {
            const explanationCell = XLSX.utils.encode_cell({ r: row, c: pivotHeaders.length + 2 });
            if (ws[explanationCell] && ws[explanationCell].v) {
                // Жёлтый фон для строк с пояснениями
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    if (ws[cellAddress]) {
                        if (!ws[cellAddress].s) ws[cellAddress].s = {};
                        ws[cellAddress].s.fill = { fgColor: { rgb: 'FFF2CC' } };
                    }
                }
            }
        }

        // Итоговая строка — зелёный фон, жирный
        const lastRow = range.e.r;
        const totalStyle = {
            fill: { fgColor: { rgb: 'C6EFCE' } },
            font: { bold: true, sz: 11 },
            alignment: { horizontal: 'right' }
        };

        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: lastRow, c: col });
            if (ws[cellAddress]) {
                ws[cellAddress].s = totalStyle;
                if (typeof ws[cellAddress].v === 'number') {
                    ws[cellAddress].t = 'n';
                    ws[cellAddress].z = '#,##0.00';
                }
            }
        }

        // Ширина колонок
        ws['!cols'] = [
            { wch: 35 }, // Контрагент
            ...pivotHeaders.map(() => ({ wch: 18 })), // Подразделения
            { wch: 18 }, // Итого
            { wch: 50 }  // Пояснение
        ];

        // Объединение заголовка
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: pivotHeaders.length + 2 } }
        ];

        // Стиль заголовка
        ws['A1'].s = {
            font: { bold: true, sz: 14 },
            alignment: { horizontal: 'center' }
        };

        return ws;
    }

    createLibraryWorksheet() {
        const contractors = this.library.getAll();

        const data = [
            ['Библиотека контрагентов'],
            [],
            ['Получатель', 'Юридическое Лицо', 'Пояснения']
        ];

        contractors.forEach(c => {
            data.push([c.name, c.organization, c.explanation]);
        });

        const ws = XLSX.utils.aoa_to_sheet(data);
        const range = XLSX.utils.decode_range(ws['!ref']);

        // Заголовки — тёмно-синий фон
        const headerStyle = {
            fill: { fgColor: { rgb: '1F3864' } },
            font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 }
        };

        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 2, c: col });
            if (ws[cellAddress]) {
                ws[cellAddress].s = headerStyle;
            }
        }

        ws['!cols'] = [
            { wch: 40 }, // Получатель
            { wch: 30 }, // Юридическое Лицо
            { wch: 60 }  // Пояснения
        ];

        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }
        ];

        ws['A1'].s = {
            font: { bold: true, sz: 14 },
            alignment: { horizontal: 'center' }
        };

        return ws;
    }

    // ===== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ =====
    cleanString(value) {
        if (!value) return '';
        return String(value).trim();
    }

    parseAmount(value) {
        if (!value) return 0;
        if (typeof value === 'number') return value;
        const cleaned = String(value).replace(/\s/g, '').replace(',', '.');
        return parseFloat(cleaned) || 0;
    }

    // ===== ОЧИСТКА ДАННЫХ =====
    clearData() {
        this.loadedData = [];
        this.pivotTables = [];
    }

    // ===== ПОЛУЧЕНИЕ ДАННЫХ ДЛЯ UI =====
    getLoadedData() {
        return this.loadedData;
    }

    getPivotTables() {
        return this.pivotTables;
    }
}
