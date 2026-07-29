// contractors-library.js — Библиотека контрагентов (API + localStorage fallback)
class ContractorsLibrary {
    constructor() {
        this.contractors = [];
        this.apiBase = window.location.origin + '/api/contractors';
        this.useApi = true;
        this._loaded = false;
    }

    // ===== ИНИЦИАЛИЗАЦИЯ: ЗАГРУЗКА ДАННЫХ =====
    async init() {
        if (this._loaded) return;
        // Пробуем загрузить через API
        const apiResult = await this.loadFromApi();
        if (apiResult && apiResult.length > 0) {
            this.useApi = true;
            console.log(`Библиотека контрагентов загружена через API: ${this.contractors.length} записей`);
        } else if (apiResult && apiResult.length === 0) {
            // API доступен, но БД пуста — пробуем мигрировать из localStorage
            this.useApi = true;
            console.log('БД пуста, проверяем localStorage...');
            const stored = this.loadFromStorage();
            if (stored.length > 0) {
                console.log(`Найдено ${stored.length} записей в localStorage, мигрируем в БД...`);
                await this.migrateLocalToApi();
            } else {
                console.log('localStorage тоже пуст — библиотека пуста');
            }
        } else {
            // API недоступен — fallback на localStorage
            this.useApi = false;
            this.contractors = this.loadFromStorage();
            console.log(`Библиотека контрагентов загружена из localStorage: ${this.contractors.length} записей`);
        }
        this._loaded = true;
    }

    // ===== ЗАГРУЗКА ЧЕРЕЗ API =====
    async loadFromApi() {
        try {
            const response = await fetch(this.apiBase);
            if (!response.ok) return null;
            const result = await response.json();
            if (result.success && Array.isArray(result.data)) {
                this.contractors = result.data;
                return result.data;
            }
            return null;
        } catch (e) {
            console.warn('API библиотеки недоступен, использую localStorage:', e.message);
            return null;
        }
    }

    // ===== МИГРАЦИЯ localStorage → API =====
    async migrateLocalToApi() {
        const stored = this.loadFromStorage();
        if (stored.length === 0) return;

        console.log(`🔄 Миграция ${stored.length} записей из localStorage в БД...`);
        try {
            const contractors = stored.map(item => ({
                name: item.name || '',
                organization: item.organization || '',
                explanation: item.explanation || ''
            }));

            const response = await fetch(`${this.apiBase}/import`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ contractors })
            });

            const result = await response.json();
            if (result.success) {
                console.log(`✅ Миграция завершена: добавлено ${result.added}, обновлено ${result.updated}`);
                // Очищаем localStorage после успешной миграции
                localStorage.removeItem('contractorsLibrary');
                // Перезагружаем данные из API
                await this.loadFromApi();
            } else {
                console.warn('⚠️ Миграция не удалась:', result.error);
            }
        } catch (e) {
            console.warn('⚠️ Ошибка миграции localStorage:', e.message);
        }
    }

    // ===== ЗАГРУЗКА ИЗ LOCALSTORAGE (fallback) =====
    loadFromStorage() {
        try {
            const stored = localStorage.getItem('contractorsLibrary');
            if (stored) {
                return JSON.parse(stored);
            }
        } catch (e) {
            console.error('Ошибка загрузки библиотеки из localStorage:', e);
        }
        return [];
    }

    // ===== СОХРАНЕНИЕ В LOCALSTORAGE (fallback) =====
    saveToStorage() {
        try {
            localStorage.setItem('contractorsLibrary', JSON.stringify(this.contractors));
        } catch (e) {
            console.error('Ошибка сохранения библиотеки в localStorage:', e);
        }
    }

    // ===== ДОБАВЛЕНИЕ КОНТРАГЕНТА =====
    async addContractor(data) {
        await this.ensureLoaded();

        if (this.useApi) {
            return await this.addContractorViaApi(data);
        } else {
            return this.addContractorLocal(data);
        }
    }

    async addContractorViaApi(data) {
        try {
            const response = await fetch(this.apiBase, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    name: data.name || '',
                    organization: data.organization || '',
                    explanation: data.explanation || ''
                })
            });

            const result = await response.json();

            if (result.success) {
                // Обновляем локальный кэш
                await this.loadFromApi();
                if (result.action === 'inserted') {
                    return { success: true, message: 'Контрагент добавлен', contractor: data };
                } else {
                    return { success: false, message: 'Контрагент уже существует, обновлены данные', updated: data };
                }
            }
            return { success: false, message: result.error || 'Ошибка сохранения' };
        } catch (e) {
            console.error('Ошибка добавления через API:', e);
            // Fallback на localStorage
            return this.addContractorLocal(data);
        }
    }

    addContractorLocal(data) {
        const existing = this.contractors.find(c =>
            this.normalizeName(c.name) === this.normalizeName(data.name)
        );

        if (existing) {
            if (data.organization) existing.organization = data.organization;
            if (data.explanation) existing.explanation = data.explanation;
            this.saveToStorage();
            return { success: false, message: 'Контрагент уже существует, обновлены данные', updated: existing };
        }

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
    async removeContractor(id) {
        await this.ensureLoaded();

        if (this.useApi) {
            try {
                const response = await fetch(`${this.apiBase}/${id}`, {
                    method: 'DELETE'
                });
                const result = await response.json();
                if (result.success) {
                    await this.loadFromApi();
                    return { success: true, message: 'Контрагент удалён' };
                }
                return { success: false, message: result.message || 'Ошибка удаления' };
            } catch (e) {
                console.error('Ошибка удаления через API:', e);
            }
        }

        // Локальное удаление
        const index = this.contractors.findIndex(c => c.id === id);
        if (index !== -1) {
            const removed = this.contractors.splice(index, 1)[0];
            this.saveToStorage();
            return { success: true, message: 'Контрагент удалён', contractor: removed };
        }
        return { success: false, message: 'Контрагент не найден' };
    }

    // ===== ОБНОВЛЕНИЕ КОНТРАГЕНТА =====
    async updateContractor(id, data) {
        await this.ensureLoaded();

        if (this.useApi) {
            try {
                const response = await fetch(`${this.apiBase}/${id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                });
                const result = await response.json();
                if (result.success) {
                    await this.loadFromApi();
                    return { success: true, message: 'Контрагент обновлён' };
                }
                return { success: false, message: result.error || 'Ошибка обновления' };
            } catch (e) {
                console.error('Ошибка обновления через API:', e);
            }
        }

        // Локальное обновление
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
        await this.ensureLoaded();

        try {
            const workbook = await this.readExcelFile(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

            if (rows.length < 2) {
                return { success: false, message: 'Файл пуст или не содержит данных' };
            }

            const headers = rows[0].map(h => String(h || '').trim());
            console.log('Заголовки файла библиотеки:', headers);

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

            const contractorsToImport = [];
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;

                const name = this.cleanCellValue(row[nameCol]);
                if (!name) continue;

                const explanation = explCol !== -1 ? this.cleanCellValue(row[explCol]) : '';
                const organization = orgCol !== -1 ? this.cleanCellValue(row[orgCol]) : '';

                contractorsToImport.push({ name, organization, explanation });
            }

            if (contractorsToImport.length === 0) {
                return { success: false, message: 'Не найдено данных для импорта' };
            }

            // Пробуем импортировать через API
            if (this.useApi) {
                try {
                    const response = await fetch(`${this.apiBase}/import`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ contractors: contractorsToImport })
                    });
                    const result = await response.json();
                    if (result.success) {
                        await this.loadFromApi();
                        return {
                            success: true,
                            message: `Импортировано: добавлено ${result.added}, обновлено ${result.updated}`,
                            added: result.added,
                            updated: result.updated
                        };
                    }
                } catch (e) {
                    console.error('Ошибка импорта через API:', e);
                }
            }

            // Локальный импорт
            let added = 0;
            let updated = 0;

            for (const item of contractorsToImport) {
                const result = this.addContractorLocal(item);
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

        ws['!cols'] = [
            { wch: 40 },
            { wch: 30 },
            { wch: 60 }
        ];

        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }
        ];

        const date = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `Библиотека_контрагентов_${date}.xlsx`);

        return { success: true, message: 'Библиотека экспортирована' };
    }

    // ===== ОЧИСТКА БИБЛИОТЕКИ =====
    async clearAll() {
        await this.ensureLoaded();

        if (this.useApi) {
            try {
                const response = await fetch(`${this.apiBase}/clear`, { method: 'POST' });
                const result = await response.json();
                if (result.success) {
                    this.contractors = [];
                    return { success: true, message: 'Библиотека очищена' };
                }
            } catch (e) {
                console.error('Ошибка очистки через API:', e);
            }
        }

        // Локальная очистка
        this.contractors = [];
        this.saveToStorage();
        return { success: true, message: 'Библиотека очищена' };
    }

    // ===== ОБНОВЛЕНИЕ ДАННЫХ (перезагрузка с сервера) =====
    async refresh() {
        this._loaded = false;
        await this.init();
        return this.getAll();
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

    // ===== ПРОВЕРКА ЗАГРУЗКИ =====
    async ensureLoaded() {
        if (!this._loaded) {
            await this.init();
        }
    }
}