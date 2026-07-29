// ============================================================
// BankFlow → Financial Analyzer — SettingsManager
// 4 вкладки: Счета, Исключения, Категоризация, Синонимы
// + Поиск, редактирование через модальное окно, выделение строк
// ============================================================
class SettingsManager {
    constructor(storage) {
        this.storage = storage;
        this.activeTab = 'accounts';
        this.searchTerm = '';            // поисковый фильтр
        this.selectedRowId = null;      // выделенная строка (data-id)
        this.editType = null;           // тип редактируемой записи (accounts/exclusions/categorization/aliases)
        this.editData = null;           // кэш данных для модалки
    }

    async init() {
        this.setupTabs();
        this.setupForms();
        this.setupTooltips();
        this.setupSearch();
        this.setupEditModal();
        this.setupTableClickHandlers();
        await this.loadAllData();
    }

    // ── Табы ─────────────────────────────────────────────
    setupTabs() {
        const tabs = document.querySelectorAll('.settings-tab');
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                this.activeTab = tab.dataset.tab;
                document.querySelectorAll('.settings-panel').forEach(p => p.style.display = 'none');
                const panel = document.getElementById(`settings-${this.activeTab}`);
                if (panel) panel.style.display = 'block';
                // Сброс прокрутки при переключении вкладок
                const pc = document.querySelector('.page-content');
                if (pc) pc.scrollTop = 0;
                window.scrollTo(0, 0);
                // Очистка поиска и выделения
                this.searchTerm = '';
                this.selectedRowId = null;
                const si = document.getElementById('settingsSearch');
                if (si) si.value = '';
            });
        });
    }

    // ── Формы добавления ─────────────────────────────────
    setupForms() {
        const addAccountBtn = document.getElementById('settingsAddAccountBtn');
        if (addAccountBtn) addAccountBtn.addEventListener('click', () => this.addAccount());
        const addExclusionBtn = document.getElementById('addExclusionBtn');
        if (addExclusionBtn) addExclusionBtn.addEventListener('click', () => this.addExclusionRule());
        const addCategorizationBtn = document.getElementById('addCategorizationBtn');
        if (addCategorizationBtn) addCategorizationBtn.addEventListener('click', () => this.addCategorizationRule());
        const addAliasBtn = document.getElementById('addAliasBtn');
        if (addAliasBtn) addAliasBtn.addEventListener('click', () => this.addAlias());

        // Импорт счетов из Excel
        const importBtn = document.getElementById('importAccountsBtn');
        const importFile = document.getElementById('importAccountsFile');
        if (importBtn && importFile) {
            importBtn.addEventListener('click', () => importFile.click());
            importFile.addEventListener('change', (e) => {
                this.importAccountsFromExcel(e.target.files[0]);
                e.target.value = '';
            });
        }

        // Enter на полях ввода → добавить
        this.bindEnterOnForms();
    }

    // ── Enter в формах добавления → вызов соответствующей кнопки ──
    bindEnterOnForms() {
        const panels = [
            { selector: '#settingsNewAccountNumber, #settingsNewAccountCompany, #settingsNewAccountBank', btnId: 'settingsAddAccountBtn' },
            { selector: '#exclusionPattern', btnId: 'addExclusionBtn' },
            { selector: '#catPattern, #catDisplayName', btnId: 'addCategorizationBtn' },
            { selector: '#aliasPattern, #aliasCanonical', btnId: 'addAliasBtn' }
        ];
        panels.forEach(p => {
            document.querySelectorAll(p.selector).forEach(el => {
                el.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') {
                        e.preventDefault();
                        const btn = document.getElementById(p.btnId);
                        if (btn) btn.click();
                    }
                });
            });
        });
    }

    // ── Поисковая строка ────────────────────────────────
    setupSearch() {
        const searchInput = document.getElementById('settingsSearch');
        const clearBtn = document.getElementById('settingsClearSearch');
        if (!searchInput || !clearBtn) return;

        searchInput.addEventListener('input', () => {
            this.searchTerm = searchInput.value.toLowerCase();
            this.loadAllData(); // перерисовываем активную таблицу
        });

        clearBtn.addEventListener('click', () => {
            searchInput.value = '';
            this.searchTerm = '';
            this.loadAllData();
        });
    }

    // ── Модальное окно редактирования ────────────────────
    setupEditModal() {
        const closeBtn = document.getElementById('closeSettingsEditModal');
        const cancelBtn = document.getElementById('cancelSettingsEditBtn');
        const saveBtn = document.getElementById('saveSettingsEditBtn');
        const overlay = document.getElementById('settingsEditModal');

        if (!overlay) return;

        [closeBtn, cancelBtn].forEach(b => {
            if (b) b.addEventListener('click', () => {
                overlay.style.display = 'none';
                this.editType = null;
                this.editData = null;
            });
        });

        if (saveBtn) {
            saveBtn.addEventListener('click', () => this.saveSettingsEdit());
        }

        // Клик вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target === overlay) {
                overlay.style.display = 'none';
                this.editType = null;
                this.editData = null;
            }
        });

        // Enter в модальном окне → сохранить
        overlay.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && overlay.style.display !== 'none') {
                e.preventDefault();
                this.saveSettingsEdit();
            }
            if (e.key === 'Escape' && overlay.style.display !== 'none') {
                overlay.style.display = 'none';
                this.editType = null;
                this.editData = null;
            }
        });
    }

    // ── Обработчики кликов на таблицах ──────────────────
    setupTableClickHandlers() {
        const tables = ['accountsTableBody', 'exclusionTableBody', 'categorizationTableBody', 'aliasesTableBody'];
        tables.forEach(tbodyId => {
            const tbody = document.getElementById(tbodyId);
            if (!tbody) return;
            // Используем делегирование событий
            tbody.addEventListener('click', (e) => {
                const btn = e.target.closest('button');
                if (btn) return; // клик по кнопке — не обрабатываем здесь
                const row = e.target.closest('tr[data-id]');
                if (!row) return;
                const id = row.dataset.id;
                this.selectRow(tbodyId, id);
            });
            tbody.addEventListener('dblclick', (e) => {
                const btn = e.target.closest('button');
                if (btn) return;
                const row = e.target.closest('tr[data-id]');
                if (!row) return;
                const id = row.dataset.id;
                const type = row.dataset.type;
                this.openEditModal(type, id);
            });
        });
    }

    selectRow(tbodyId, id) {
        // Снимаем выделение со всех строк во всех таблицах
        document.querySelectorAll('tr.selected-settings').forEach(r => r.classList.remove('selected-settings'));
        const tbody = document.getElementById(tbodyId);
        if (!tbody) return;
        const row = tbody.querySelector(`tr[data-id="${id}"]`);
        if (row) {
            row.classList.add('selected-settings');
        }
        this.selectedRowId = id;
    }

    // ── Открытие модалки редактирования ──────────────────
    openEditModal(type, id) {
        this.editType = type;
        const overlay = document.getElementById('settingsEditModal');
        const title = document.getElementById('settingsEditModalTitle');
        const body = document.getElementById('settingsEditModalBody');
        if (!overlay || !title || !body) return;

        // Находим данные записи
        let item = null;
        switch (type) {
            case 'accounts':
                item = this._accountsData?.find(a => String(a.id) === String(id));
                title.textContent = 'Редактировать счёт';
                body.innerHTML = `
                    <input type="hidden" id="settingsEditId" value="${this.escapeHtml(String(item?.id || ''))}">
                    <div class="form-group">
                        <label for="settingsEditAccountNumber">Номер счёта</label>
                        <input type="text" id="settingsEditAccountNumber" class="form-input" value="${this.escapeHtml(item?.account_number || '')}" maxlength="20">
                    </div>
                    <div class="form-group">
                        <label for="settingsEditAccountCompany">Название компании</label>
                        <input type="text" id="settingsEditAccountCompany" class="form-input" value="${this.escapeHtml(item?.company_name || '')}">
                    </div>
                    <div class="form-group">
                        <label for="settingsEditAccountBank">Банк</label>
                        <input type="text" id="settingsEditAccountBank" class="form-input" value="${this.escapeHtml(item?.bank_name || '')}">
                    </div>
                `;
                break;
            case 'exclusions':
                item = this._exclusionsData?.find(r => String(r.id) === String(id));
                title.textContent = 'Редактировать правило исключения';
                body.innerHTML = `
                    <input type="hidden" id="settingsEditId" value="${this.escapeHtml(String(item?.id || ''))}">
                    <div class="form-group">
                        <label for="settingsEditExclusionType">Тип</label>
                        <select id="settingsEditExclusionType" class="form-input">
                            <option value="purpose" ${item?.type === 'purpose' ? 'selected' : ''}>Назначение</option>
                            <option value="counterparty" ${item?.type === 'counterparty' ? 'selected' : ''}>Контрагент</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="settingsEditExclusionPattern">Шаблон</label>
                        <input type="text" id="settingsEditExclusionPattern" class="form-input" value="${this.escapeHtml(item?.pattern || '')}">
                    </div>
                    <div class="form-group">
                        <label style="display:flex;align-items:center;gap:8px;">
                            <input type="checkbox" id="settingsEditExclusionIsRegex" ${item?.is_regex ? 'checked' : ''}> Regex
                        </label>
                    </div>
                `;
                break;
            case 'categorization':
                item = this._categorizationData?.find(r => String(r.id) === String(id));
                title.textContent = 'Редактировать правило категоризации';
                body.innerHTML = `
                    <input type="hidden" id="settingsEditId" value="${this.escapeHtml(String(item?.id || ''))}">
                    <div class="form-group">
                        <label for="settingsEditCatField">Поле</label>
                        <select id="settingsEditCatField" class="form-input">
                            <option value="purpose" ${item?.field === 'purpose' ? 'selected' : ''}>Назначение</option>
                            <option value="counterparty" ${item?.field === 'counterparty' ? 'selected' : ''}>Контрагент</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="settingsEditCatPattern">Шаблон</label>
                        <input type="text" id="settingsEditCatPattern" class="form-input" value="${this.escapeHtml(item?.pattern || '')}">
                    </div>
                    <div class="form-group">
                        <label for="settingsEditCatDisplayName">Отображаемое имя</label>
                        <input type="text" id="settingsEditCatDisplayName" class="form-input" value="${this.escapeHtml(item?.display_name || '')}">
                    </div>
                `;
                break;
            case 'aliases':
                item = this._aliasesData?.find(a => String(a.id) === String(id));
                title.textContent = 'Редактировать синоним';
                body.innerHTML = `
                    <input type="hidden" id="settingsEditId" value="${this.escapeHtml(String(item?.id || ''))}">
                    <div class="form-group">
                        <label for="settingsEditAliasPattern">Шаблон</label>
                        <input type="text" id="settingsEditAliasPattern" class="form-input" value="${this.escapeHtml(item?.pattern || '')}">
                    </div>
                    <div class="form-group">
                        <label for="settingsEditAliasCanonical">Каноническое название</label>
                        <input type="text" id="settingsEditAliasCanonical" class="form-input" value="${this.escapeHtml(item?.canonical || '')}">
                    </div>
                    <div class="form-group">
                        <label for="settingsEditAliasMatchType">Тип совпадения</label>
                        <select id="settingsEditAliasMatchType" class="form-input">
                            <option value="contains" ${item?.match_type === 'contains' ? 'selected' : ''}>Содержит</option>
                            <option value="exact" ${item?.match_type === 'exact' ? 'selected' : ''}>Точное</option>
                            <option value="regex" ${item?.match_type === 'regex' ? 'selected' : ''}>Regex</option>
                        </select>
                    </div>
                `;
                break;
            default:
                return;
        }

        if (!item) return;

        overlay.style.display = 'flex';
        // Фокус на первом поле ввода
        const firstInput = body.querySelector('input[type="text"], input:not([type="hidden"])');
        if (firstInput) setTimeout(() => firstInput.focus(), 100);
    }

    // ── Сохранение изменений из модалки ─────────────────
    async saveSettingsEdit() {
        const type = this.editType;
        const id = document.getElementById('settingsEditId')?.value;
        if (!type || !id) return;

        let payload = {};
        let url = '';

        switch (type) {
            case 'accounts': {
                const number = document.getElementById('settingsEditAccountNumber')?.value?.trim() || '';
                const company = document.getElementById('settingsEditAccountCompany')?.value?.trim() || '';
                const bank = document.getElementById('settingsEditAccountBank')?.value?.trim() || '';
                if (!number || !company) {
                    if (window.app) window.app.showNotification('Номер счёта и компания обязательны', 'error');
                    return;
                }
                payload = { account_number: number, company_name: company, bank_name: bank };
                url = `/api/account-mapping/${id}`;
                break;
            }
            case 'exclusions': {
                const ruleType = document.getElementById('settingsEditExclusionType')?.value || 'purpose';
                const pattern = document.getElementById('settingsEditExclusionPattern')?.value?.trim() || '';
                const isRegex = document.getElementById('settingsEditExclusionIsRegex')?.checked || false;
                if (!pattern) {
                    if (window.app) window.app.showNotification('Укажите шаблон', 'error');
                    return;
                }
                payload = { rule_type: ruleType, pattern, is_regex: isRegex };
                url = `/api/exclusion-rules/${id}`;
                break;
            }
            case 'categorization': {
                const field = document.getElementById('settingsEditCatField')?.value || 'purpose';
                const pattern = document.getElementById('settingsEditCatPattern')?.value?.trim() || '';
                const displayName = document.getElementById('settingsEditCatDisplayName')?.value?.trim() || '';
                if (!pattern || !displayName) {
                    if (window.app) window.app.showNotification('Заполните все поля', 'error');
                    return;
                }
                payload = { field, pattern, display_name: displayName };
                url = `/api/categorization-rules/${id}`;
                break;
            }
            case 'aliases': {
                const pattern = document.getElementById('settingsEditAliasPattern')?.value?.trim() || '';
                const canonical = document.getElementById('settingsEditAliasCanonical')?.value?.trim() || '';
                const matchType = document.getElementById('settingsEditAliasMatchType')?.value || 'contains';
                if (!pattern || !canonical) {
                    if (window.app) window.app.showNotification('Заполните все поля', 'error');
                    return;
                }
                payload = { pattern, canonical_name: canonical, match_type: matchType };
                url = `/api/company-aliases/${id}`;
                break;
            }
        }

        try {
            const resp = await fetch(url, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await resp.json();
            if (result.success) {
                if (window.app) window.app.showNotification('Изменения сохранены', 'success');
                document.getElementById('settingsEditModal').style.display = 'none';
                this.editType = null;
                this.editData = null;
                await this.loadAllData();
                if (type !== 'accounts') await this.storage.loadConfig();
            } else {
                if (window.app) window.app.showNotification(result.error || 'Ошибка сохранения', 'error');
            }
        } catch (e) {
            console.error('Ошибка сохранения редактирования:', e);
            if (window.app) window.app.showNotification('Ошибка соединения с сервером', 'error');
        }
    }

    // ── Загрузка данных ──────────────────────────────────
    async loadAllData() {
        await this.loadAccounts();
        await this.loadExclusionRules();
        await this.loadCategorizationRules();
        await this.loadAliases();
    }

    // ── Фильтрация строк по поисковому запросу ──────────
    filterVisibleRows(tbody, term) {
        if (!tbody) return;
        const rows = tbody.querySelectorAll('tr[data-id]');
        rows.forEach(row => {
            if (!term) {
                row.style.display = '';
                return;
            }
            const text = row.textContent.toLowerCase();
            row.style.display = text.includes(term) ? '' : 'none';
        });
    }

    async apiGet(url) {
        const resp = await fetch(url);
        const data = await resp.json();
        return data.success ? data.data : [];
    }

    async apiPost(url, body) {
        const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        return await resp.json();
    }

    async apiDelete(url) {
        const resp = await fetch(url, { method: 'DELETE' });
        return await resp.json();
    }

    async apiToggle(url) {
        const resp = await fetch(url, { method: 'PUT' });
        return await resp.json();
    }

    // ════════════════════════════════════════════════════════
    //  1. СЧЕТА
    // ════════════════════════════════════════════════════════
    async loadAccounts() {
        const data = await this.apiGet('/api/account-mapping?all=true');
        this._accountsData = data;
        const tbody = document.getElementById('accountsTableBody');
        if (!tbody) return;
        if (!data.length) {
            const term = this.searchTerm;
            tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:20px;color:#999;">${term ? 'Ничего не найдено' : 'Нет счетов. Добавьте счёт или загрузите выписку.'}</td></tr>`;
            return;
        }
        let html = '';
        data.forEach(a => {
            const rowText = (a.account_number + ' ' + a.company_name + ' ' + a.bank_name).toLowerCase();
            const hidden = this.searchTerm && !rowText.includes(this.searchTerm) ? ' style="display:none;"' : '';
            html += `
                <tr data-id="${a.id}" data-type="accounts" class="${a.is_active ? '' : 'inactive-row'}"${hidden}>
                    <td>${this.escapeHtml(a.account_number)}</td>
                    <td>${this.escapeHtml(a.company_name)}</td>
                    <td>${this.escapeHtml(a.bank_name)}</td>
                    <td>${a.is_active ? '✅' : '❌'}</td>
                    <td>
                        <button class="btn btn-sm ${a.is_active ? 'btn-danger' : 'btn-success'}" onclick="window.settingsManager.toggleAccount(${a.id})">${a.is_active ? 'Отключить' : 'Включить'}</button>
                    </td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
        // Восстанавливаем выделение
        if (this.selectedRowId) {
            const row = tbody.querySelector(`tr[data-id="${this.selectedRowId}"]`);
            if (row) row.classList.add('selected-settings');
            else this.selectedRowId = null;
        }
    }

    async addAccount() {
        const number = document.getElementById('settingsNewAccountNumber').value.trim();
        const company = document.getElementById('settingsNewAccountCompany').value.trim();
        const bank = document.getElementById('settingsNewAccountBank').value.trim();

        if (!number || !company) {
            if (window.app) window.app.showNotification('Номер счёта и компания обязательны', 'error');
            return;
        }

        const result = await this.apiPost('/api/account-mapping', {
            account_number: number,
            company_name: company,
            bank_name: bank
        });

        if (result.success) {
            if (window.app) window.app.showNotification('Счёт добавлен', 'success');
            document.getElementById('settingsNewAccountNumber').value = '';
            document.getElementById('settingsNewAccountCompany').value = '';
            document.getElementById('settingsNewAccountBank').value = '';
            await this.loadAccounts();
            await this.storage.loadConfig();
        } else {
            if (window.app) window.app.showNotification(result.error || 'Ошибка', 'error');
        }
    }

    async toggleAccount(id) {
        const result = await this.apiToggle(`/api/account-mapping/${id}/toggle`);
        if (result.success) {
            if (window.app) window.app.showNotification(result.is_active ? 'Счёт включён' : 'Счёт отключён', 'info');
            await this.loadAccounts();
            await this.storage.loadConfig();
        }
    }

    // ════════════════════════════════════════════════════════
    //  2. ПРАВИЛА ИСКЛЮЧЕНИЯ
    // ════════════════════════════════════════════════════════
    async loadExclusionRules() {
        const data = await this.apiGet('/api/exclusion-rules');
        this._exclusionsData = data;
        const tbody = document.getElementById('exclusionTableBody');
        if (!tbody) return;
        if (!data.length) {
            const term = this.searchTerm;
            tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:20px;color:#999;">${term ? 'Ничего не найдено' : 'Нет правил исключения'}</td></tr>`;
            return;
        }
        let html = '';
        data.forEach(r => {
            const rowText = (r.type + ' ' + r.pattern + ' ' + (r.is_regex ? 'Regex' : 'Текст')).toLowerCase();
            const hidden = this.searchTerm && !rowText.includes(this.searchTerm) ? ' style="display:none;"' : '';
            html += `
                <tr data-id="${r.id}" data-type="exclusions"${hidden}>
                    <td>${r.type === 'purpose' ? 'Назначение' : 'Контрагент'}</td>
                    <td><code>${this.escapeHtml(r.pattern)}</code></td>
                    <td>${r.is_regex ? 'Regex' : 'Текст'}</td>
                    <td><button class="btn btn-sm btn-danger" onclick="window.settingsManager.deleteExclusionRule(${r.id})">Удалить</button></td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
        if (this.selectedRowId) {
            const row = tbody.querySelector(`tr[data-id="${this.selectedRowId}"]`);
            if (row) row.classList.add('selected-settings');
            else this.selectedRowId = null;
        }
    }

    async addExclusionRule() {
        const type = document.getElementById('exclusionType').value;
        const pattern = document.getElementById('exclusionPattern').value.trim();
        const isRegex = document.getElementById('exclusionIsRegex').checked;

        if (!pattern) {
            if (window.app) window.app.showNotification('Укажите шаблон', 'error');
            return;
        }

        const result = await this.apiPost('/api/exclusion-rules', {
            rule_type: type,
            pattern: pattern,
            is_regex: isRegex
        });

        if (result.success) {
            if (window.app) window.app.showNotification('Правило добавлено', 'success');
            document.getElementById('exclusionPattern').value = '';
            await this.loadExclusionRules();
            await this.storage.loadConfig();
        } else {
            if (window.app) window.app.showNotification(result.error || 'Ошибка', 'error');
        }
    }

    async deleteExclusionRule(id) {
        if (!confirm('Удалить правило исключения?')) return;
        await this.apiDelete(`/api/exclusion-rules/${id}`);
        if (window.app) window.app.showNotification('Правило удалено', 'info');
        if (this.selectedRowId === String(id)) this.selectedRowId = null;
        await this.loadExclusionRules();
        await this.storage.loadConfig();
    }

    // ════════════════════════════════════════════════════════
    //  3. ПРАВИЛА КАТЕГОРИЗАЦИИ
    // ════════════════════════════════════════════════════════
    async loadCategorizationRules() {
        const data = await this.apiGet('/api/categorization-rules');
        this._categorizationData = data;
        const tbody = document.getElementById('categorizationTableBody');
        if (!tbody) return;
        if (!data.length) {
            const term = this.searchTerm;
            tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:20px;color:#999;">${term ? 'Ничего не найдено' : 'Нет правил категоризации'}</td></tr>`;
            return;
        }
        let html = '';
        data.forEach(r => {
            const rowText = (r.field + ' ' + r.pattern + ' ' + r.display_name).toLowerCase();
            const hidden = this.searchTerm && !rowText.includes(this.searchTerm) ? ' style="display:none;"' : '';
            html += `
                <tr data-id="${r.id}" data-type="categorization"${hidden}>
                    <td>${r.field === 'purpose' ? 'Назначение' : 'Контрагент'}</td>
                    <td><code>${this.escapeHtml(r.pattern)}</code></td>
                    <td><strong>${this.escapeHtml(r.display_name)}</strong></td>
                    <td><button class="btn btn-sm btn-danger" onclick="window.settingsManager.deleteCategorizationRule(${r.id})">Удалить</button></td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
        if (this.selectedRowId) {
            const row = tbody.querySelector(`tr[data-id="${this.selectedRowId}"]`);
            if (row) row.classList.add('selected-settings');
            else this.selectedRowId = null;
        }
    }

    async addCategorizationRule() {
        const field = document.getElementById('catField').value;
        const pattern = document.getElementById('catPattern').value.trim();
        const displayName = document.getElementById('catDisplayName').value.trim();

        if (!pattern || !displayName) {
            if (window.app) window.app.showNotification('Заполните все поля', 'error');
            return;
        }

        const result = await this.apiPost('/api/categorization-rules', {
            field, pattern, display_name: displayName
        });

        if (result.success) {
            if (window.app) window.app.showNotification('Правило категоризации добавлено', 'success');
            document.getElementById('catPattern').value = '';
            document.getElementById('catDisplayName').value = '';
            await this.loadCategorizationRules();
            await this.storage.loadConfig();
        } else {
            if (window.app) window.app.showNotification(result.error || 'Ошибка', 'error');
        }
    }

    async deleteCategorizationRule(id) {
        if (!confirm('Удалить правило категоризации?')) return;
        await this.apiDelete(`/api/categorization-rules/${id}`);
        if (window.app) window.app.showNotification('Правило удалено', 'info');
        if (this.selectedRowId === String(id)) this.selectedRowId = null;
        await this.loadCategorizationRules();
        await this.storage.loadConfig();
    }

    // ════════════════════════════════════════════════════════
    //  4. СИНОНИМЫ КОМПАНИЙ
    // ════════════════════════════════════════════════════════
    async loadAliases() {
        const data = await this.apiGet('/api/company-aliases');
        this._aliasesData = data;
        const tbody = document.getElementById('aliasesTableBody');
        if (!tbody) return;
        if (!data.length) {
            const term = this.searchTerm;
            tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:20px;color:#999;">${term ? 'Ничего не найдено' : 'Нет синонимов компаний'}</td></tr>`;
            return;
        }
        let html = '';
        data.forEach(a => {
            const rowText = (a.pattern + ' ' + a.canonical + ' ' + (a.match_type === 'exact' ? 'Точное' : a.match_type === 'regex' ? 'Regex' : 'Содержит')).toLowerCase();
            const hidden = this.searchTerm && !rowText.includes(this.searchTerm) ? ' style="display:none;"' : '';
            html += `
                <tr data-id="${a.id}" data-type="aliases"${hidden}>
                    <td><code>${this.escapeHtml(a.pattern)}</code></td>
                    <td><strong>${this.escapeHtml(a.canonical)}</strong></td>
                    <td>${a.match_type === 'exact' ? 'Точное' : a.match_type === 'regex' ? 'Regex' : 'Содержит'}</td>
                    <td><button class="btn btn-sm btn-danger" onclick="window.settingsManager.deleteAlias(${a.id})">Удалить</button></td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
        if (this.selectedRowId) {
            const row = tbody.querySelector(`tr[data-id="${this.selectedRowId}"]`);
            if (row) row.classList.add('selected-settings');
            else this.selectedRowId = null;
        }
    }

    async addAlias() {
        const pattern = document.getElementById('aliasPattern').value.trim();
        const canonical = document.getElementById('aliasCanonical').value.trim();
        const matchType = document.getElementById('aliasMatchType').value;

        if (!pattern || !canonical) {
            if (window.app) window.app.showNotification('Заполните все поля', 'error');
            return;
        }

        const result = await this.apiPost('/api/company-aliases', {
            pattern, canonical_name: canonical, match_type: matchType
        });

        if (result.success) {
            if (window.app) window.app.showNotification('Синоним добавлен', 'success');
            document.getElementById('aliasPattern').value = '';
            document.getElementById('aliasCanonical').value = '';
            await this.loadAliases();
            await this.storage.loadConfig();
        } else {
            if (window.app) window.app.showNotification(result.error || 'Ошибка', 'error');
        }
    }

    async deleteAlias(id) {
        if (!confirm('Удалить синоним?')) return;
        await this.apiDelete(`/api/company-aliases/${id}`);
        if (window.app) window.app.showNotification('Синоним удалён', 'info');
        if (this.selectedRowId === String(id)) this.selectedRowId = null;
        await this.loadAliases();
        await this.storage.loadConfig();
    }

    // ── Импорт счетов из Excel ─────────────────────────
    async importAccountsFromExcel(file) {
        if (!file) return;
        try {
            const workbook = await this.readExcel(file);
            const rows = this.parseAccountExcel(workbook);
            if (!rows.length) {
                if (window.app) window.app.showNotification('Не найдено данных для импорта. Ожидаются колонки: Юрлицо, Счёт, Банк', 'warning');
                return;
            }
            let imported = 0;
            for (const row of rows) {
                const result = await this.apiPost('/api/account-mapping', {
                    account_number: row.account,
                    company_name: row.company,
                    bank_name: row.bank
                });
                if (result.success) imported++;
            }
            if (window.app) window.app.showNotification(`Импортировано ${imported} из ${rows.length} счетов`, 'success');
            await this.loadAccounts();
            await this.storage.loadConfig();
        } catch (e) {
            console.error(e);
            if (window.app) window.app.showNotification('Ошибка чтения Excel', 'error');
        }
    }

    readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    resolve(XLSX.read(new Uint8Array(e.target.result), { type: 'array' }));
                } catch (err) { reject(err); }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseAccountExcel(workbook) {
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        if (!rows.length) return [];

        let companyCol = -1, accountCol = -1, bankCol = -1;
        rows[0].forEach((cell, i) => {
            const c = String(cell || '').toLowerCase();
            if (c.includes('юр') || c.includes('лиц') || c.includes('компан') || c.includes('назван') || c.includes('организац')) companyCol = i;
            if (c.includes('счет') || c.includes('счёт') || c.includes('расч')) accountCol = i;
            if (c.includes('банк') || c.includes('кред')) bankCol = i;
        });
        if (accountCol === -1) accountCol = 1;
        if (companyCol === -1) companyCol = 0;
        if (bankCol === -1) bankCol = 2;

        const result = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || !row.length) continue;
            const company = String(row[companyCol] || '').trim();
            const accountRaw = String(row[accountCol] || '').trim();
            const bank = String(row[bankCol] || '').trim();
            const accountMatch = accountRaw.match(/\d{20}/);
            if (accountMatch && company) {
                result.push({ account: accountMatch[0], company, bank });
            }
        }
        return result;
    }

    // ── Утилиты ───────────────────────────────────────────
    // ── Тултипы (подсказки при фокусе на полях) ──────────
    setupTooltips() {
        const tooltip = document.getElementById('tooltipPopup');
        if (!tooltip) return;
        document.querySelectorAll('.has-tooltip').forEach(el => {
            el.addEventListener('focus', () => {
                const text = el.dataset.tooltip;
                if (!text) return;
                const rect = el.getBoundingClientRect();
                const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
                const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
                tooltip.textContent = text;
                tooltip.style.display = 'block';
                tooltip.style.left = (rect.left + scrollLeft) + 'px';
                tooltip.style.top = (rect.bottom + scrollTop + 8) + 'px';
            });
            el.addEventListener('blur', () => {
                tooltip.style.display = 'none';
            });
        });
    }

    escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }
}