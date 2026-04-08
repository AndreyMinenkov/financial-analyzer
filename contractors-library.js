// contractors-library.js — Библиотека контрагентов
class ContractorsLibrary {
    constructor() {
        this.contractors = this.loadFromStorage();
    }

    // ===== ЗАГРУЗКА ИЗ LOCALSTORAGE =====
    loadFromStorage() {
        try {
            const stored = localStorage.getItem('contractorsLibrary');
            if (stored) {
                return JSON.parse(stored);
            }
        } catch (e) {
            console.error('Ошибка загрузки библиотеки контрагентов:', e);
        }
        return [];
    }

    // ===== СОХРАНЕНИЕ В LOCALSTORAGE =====
    saveToStorage() {
        try {
            localStorage.setItem('contractorsLibrary', JSON.stringify(this.contractors));
        } catch (e) {
            console.error('Ошибка сохранения библиотеки контрагентов:', e);
        }
    }

    // ===== ДОБАВЛЕНИЕ КОНТРАГЕНТА =====
    addContractor(data) {
        // Проверяем, есть ли уже такой контрагент
        const existing = this.contractors.find(c =>
            this.normalizeName(c.name) === this.normalizeName(data.name)
        );

        if (existing) {
            // Обновляем существующую запись
            if (data.organization) existing.organization = data.organization;
            if (data.explanation) existing.explanation = data.explanation;
            this.saveToStorage();
            return { success: false, message: 'Контрагент уже существует, обновлены данные', updated: existing };
        }

        // Генерируем уникальный ID (timestamp + случайный суффикс)
        const id = `${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;

        const contractor = {
            id,
            name: data.name || '',
            organization: data.organization || '',
            explanation: data.explanation || ''
        };

        this.contractors.push(contractor);
        this.saveToStorage();
        return { success: true, message: 'Контрагент добавлен', contractor };
    }

    // ===== ПОИСК КОНТРАГЕНТА =====
    findByContractor(name) {
        const normalizedName = this.normalizeName(name);
        return this.contractors.find(c => this.normalizeName(c.name) === normalizedName);
    }

    // ===== ПОЛУЧЕНИЕ ВСЕХ КОНТРАГЕНТОВ =====
    getAll() {
        return [...this.contractors].sort((a, b) => a.name.localeCompare(b.name));
    }

    // ===== УДАЛЕНИЕ КОНТРАГЕНТА =====
    removeContractor(id) {
        const index = this.contractors.findIndex(c => c.id === id);
        if (index !== -1) {
            const removed = this.contractors.splice(index, 1)[0];
            this.saveToStorage();
            return { success: true, message: 'Контрагент удалён', contractor: removed };
        }
        return { success: false, message: 'Контрагент не найден' };
    }

    // ===== ОБНОВЛЕНИЕ КОНТРАГЕНТА =====
    updateContractor(id, data) {
        const contractor = this.contractors.find(c => c.id === id);
        if (contractor) {
            if (data.name !== undefined) contractor.name = data.name;
            if (data.organization !== undefined) contractor.organization = data.organization;
            if (data.explanation !== undefined) contractor.explanation = data.explanation;
            this.saveToStorage();
            return { success: true, message: 'Контрагент обновлён', contractor };
        }
        return { success: false, message: 'Контрагент не найден' };
    }

    // ===== ИМПОРТ ИЗ EXCEL =====
    async importFromExcel(file) {
        try {
            const workbook = await this.readExcelFile(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

            if (rows.length < 2) {
                return { success: false, message: 'Файл пуст или не содержит данных' };
            }

            const headers = rows[0].map(h => String(h || '').trim());
            console.log('Заголовки файла библиотеки:', headers);

            // Ищем колонки по разным вариантам написания
            const nameCol = this.findColumn(headers, [
                'получатель', 'наименование', 'название', 'контрагент', 'имя', 'заказчик'
            ]);
            const explCol = this.findColumn(headers, [
                'пояснения', 'пояснение', 'типичное пояснение', 'описание', 'комментарий', 'назначение платежа'
            ]);
            const orgCol = this.findColumn(headers, [
                'юридическое лицо', 'юр лицо', 'организация', 'org', 'компания', 'юл'
            ]);

            console.log('Найденные колонки:', { nameCol, explCol, orgCol });

            if (nameCol === -1) {
                return { 
                    success: false, 
                    message: 'Не найдена колонка "Получатель" или "Наименование". Ожидаемые колонки: Получатель, Пояснения, Юридическое Лицо' 
                };
            }

            let added = 0;
            let updated = 0;
            let skipped = 0;

            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row) {
                    skipped++;
                    continue;
                }

                // Получаем значение из колонки "Получатель"
                const name = this.cleanCellValue(row[nameCol]);
                if (!name) {
                    skipped++;
                    continue;
                }

                // Получаем пояснение и организацию (порядок зависит от найденных колонок)
                const explanation = explCol !== -1 ? this.cleanCellValue(row[explCol]) : '';
                const organization = orgCol !== -1 ? this.cleanCellValue(row[orgCol]) : '';

                console.log(`Строка ${i + 1}: Получатель="${name}", Пояснение="${explanation}", Организация="${organization}"`);

                const result = this.addContractor({ name, organization, explanation });
                if (result.success) {
                    added++;
                } else if (result.updated) {
                    updated++;
                }
            }

            return {
                success: true,
                message: `Импортировано: добавлено ${added}, обновлено ${updated}`,
                added,
                updated
            };
        } catch (error) {
            console.error('Ошибка импорта:', error);
            return { success: false, message: 'Ошибка импорта: ' + error.message };
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
                    reject(new Error('Ошибка чтения Excel: ' + error.message));
                }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    findColumn(headers, patterns) {
        for (let i = 0; i < headers.length; i++) {
            const headerLower = headers[i].toLowerCase();
            for (const pattern of patterns) {
                if (headerLower.includes(pattern.toLowerCase())) {
                    return i;
                }
            }
        }
        return -1;
    }

    cleanCellValue(value) {
        if (value === null || value === undefined) return '';
        return String(value).trim();
    }

    // ===== ЭКСПОРТ В EXCEL =====
    exportToExcel() {
        if (this.contractors.length === 0) {
            return { success: false, message: 'Библиотека пуста' };
        }

        const data = [
            ['Библиотека контрагентов'],
            [],
            ['Получатель', 'Юридическое Лицо', 'Пояснения']
        ];

        this.contractors.forEach(c => {
            data.push([c.name, c.organization, c.explanation]);
        });

        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Библиотека контрагентов');

        // Ширина колонок
        ws['!cols'] = [
            { wch: 40 }, // Получатель
            { wch: 30 }, // Юридическое Лицо
            { wch: 60 }  // Пояснения
        ];

        // Заголовок
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }
        ];

        const date = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `Библиотека_контрагентов_${date}.xlsx`);

        return { success: true, message: 'Библиотека экспортирована' };
    }

    // ===== ОЧИСТКА БИБЛИОТЕКИ =====
    clearAll() {
        if (confirm('Вы уверены, что хотите очистить всю библиотеку контрагентов?')) {
            this.contractors = [];
            this.saveToStorage();
            return { success: true, message: 'Библиотека очищена' };
        }
        return { success: false, message: 'Отменено' };
    }

    // ===== НОРМАЛИЗАЦИЯ НАЗВАНИЯ =====
    normalizeName(name) {
        if (!name) return '';
        return name.toUpperCase().replace(/\s+/g, ' ').trim();
    }

    // ===== СТАТИСТИКА =====
    getStats() {
        return {
            total: this.contractors.length,
            withExplanation: this.contractors.filter(c => c.explanation).length,
            withOrganization: this.contractors.filter(c => c.organization).length
        };
    }
}
