// app.js - Основной файл приложения
class App {
    constructor() {
        this.currentPage = 'upload';
        this.storage = new StorageManager();
        this.receiptsManager = new ReceiptsManager(this.storage);
        this.balancesManager = new BalancesManager(this.storage);
        this.debtManager = new DebtReconciliationManager(this.storage);
        this.contractorsLibrary = new ContractorsLibrary();
        this.supplierPayments = new SupplierPaymentsManager(this.contractorsLibrary);
        this.reportsManager = new ReportsManager();
        this.originalWorkbook = null; // Для хранения оригинального файла поставщиков
        this.init();
    }

    init() {
        console.log('Initializing Financial Analysis App...');
        this.setupNavigation();
        this.setupSidebarToggle();
        this.setupEventListeners();
        this.setupContractorsManager();
        this.reportsManager.init();
        this.loadCurrentPage();
        this.updateStats();
        this.setHeaderDate();

        // Устанавливаем сегодняшнюю дату по умолчанию
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('calculationDate').value = today;

        console.log('App initialized successfully');
    }

    setupNavigation() {
        const navButtons = document.querySelectorAll('.sidebar-nav-btn');
        const pages = document.querySelectorAll('.page');

        navButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                e.preventDefault();
                const targetPage = button.getAttribute('data-page');
                this.switchPage(targetPage);
                // Закрываем мобильное меню после выбора
                document.getElementById('sidebar')?.classList.remove('open');
            });
        });
    }

    setupSidebarToggle() {
        const toggle = document.getElementById('sidebarToggle');
        const sidebar = document.getElementById('sidebar');
        if (toggle && sidebar) {
            toggle.addEventListener('click', () => {
                sidebar.classList.toggle('collapsed');
            });
        }

        const mobileBtn = document.getElementById('mobileMenuBtn');
        if (mobileBtn && sidebar) {
            mobileBtn.addEventListener('click', () => {
                sidebar.classList.toggle('open');
            });
        }
    }

    setHeaderDate() {
        const el = document.getElementById('headerDate');
        if (el) {
            const now = new Date();
            el.textContent = now.toLocaleDateString('ru-RU', {
                day: 'numeric', month: 'long', year: 'numeric'
            });
        }
    }

    switchPage(targetPage) {
        console.log('Switching to page:', targetPage);
        this.currentPage = targetPage;

        const navButtons = document.querySelectorAll('.sidebar-nav-btn');
        const pages = document.querySelectorAll('.page');

        // Убираем активный класс у всех кнопок
        navButtons.forEach(btn => btn.classList.remove('active'));

        // Добавляем активный класс текущей кнопке
        const activeButton = Array.from(navButtons).find(btn =>
            btn.getAttribute('data-page') === targetPage
        );
        if (activeButton) {
            activeButton.classList.add('active');
        }

        // Скрываем все страницы
        pages.forEach(page => {
            page.style.display = 'none';
            page.classList.remove('active');
        });

        // Показываем целевую страницу
        const targetPageElement = document.getElementById(`${targetPage}-page`);
        if (targetPageElement) {
            targetPageElement.style.display = 'block';
            targetPageElement.classList.add('active');

            // Обновляем breadcrumb
            this.updateBreadcrumb(targetPage);

            // Обновляем данные на странице
            this.updatePageData(targetPage);
        }
    }

    updateBreadcrumb(pageName) {
        const el = document.getElementById('breadcrumbCurrent');
        if (!el) return;
        const names = {
            upload: 'Загрузка выписок',
            receipts: 'Поступления',
            balances: 'Остатки',
            debt: 'Дебиторка',
            suppliers: 'Оплаты поставщикам',
            library: 'Библиотека',
            reports: 'Отчёты'
        };
        el.textContent = names[pageName] || pageName;
    }

    updatePageData(pageName) {
        switch(pageName) {
            case 'receipts':
                this.receiptsManager.updateTable();
                break;
            case 'balances':
                this.balancesManager.updateTable();
                break;
            case 'debt':
                this.updateReconciliationUI();
                break;
            case 'suppliers':
                this.updateSuppliersUI();
                break;
            case 'library':
                this.renderLibraryTable();
                break;
        }
    }

    updateSuppliersUI() {
        // Обновляем статистику на странице поставщиков
        const stats = document.getElementById('suppliersStats');
        const pivotTables = this.supplierPayments.getPivotTables();
        const libraryStats = this.contractorsLibrary.getStats();

        if (pivotTables.length > 0) {
            stats.style.display = 'grid';
            document.getElementById('suppliersRegistriesCount').textContent = pivotTables.length;
        }

        // Обновляем статистику на странице библиотеки
        this.updateLibraryStats();
    }

    updateLibraryStats() {
        const libraryStats = this.contractorsLibrary.getStats();
        document.getElementById('libTotal').textContent = libraryStats.total;
        document.getElementById('libWithExplanation').textContent = libraryStats.withExplanation;
    }

    updateReconciliationUI() {
        const stats = this.debtManager.getStats();
        if (stats.debtRows > 0 || stats.receiptsWithDates > 0) {
            this.showReconciliationStats(stats);
        }
    }

    setupEventListeners() {
        // Загрузка файлов (основная)
        const selectFilesBtn = document.getElementById('selectFilesBtn');
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.txt,.1c,.csv';
        fileInput.multiple = true;

        selectFilesBtn.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', (e) => {
            this.handleFileSelect(e.target.files);
            fileInput.value = '';
        });

        // Drag and drop для основной загрузки
        const dropArea = document.getElementById('dropArea');
        dropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = '#2563eb';
            dropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        dropArea.addEventListener('dragleave', () => {
            dropArea.style.borderColor = '';
            dropArea.style.background = '';
        });

        dropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = '';
            dropArea.style.background = '';

            if (e.dataTransfer.files.length) {
                this.handleFileSelect(e.dataTransfer.files);
            }
        });

        // Обработка файлов
        const processFilesBtn = document.getElementById('processFilesBtn');
        processFilesBtn.addEventListener('click', () => {
            this.processFiles();
        });

        // Очистка файлов
        const clearFilesBtn = document.getElementById('clearFilesBtn');
        clearFilesBtn.addEventListener('click', () => {
            this.clearFiles();
        });

        // Экспорт данных
        document.getElementById('exportReceiptsBtn').addEventListener('click', () => {
            this.receiptsManager.exportToExcel();
        });

        document.getElementById('exportBalancesBtn').addEventListener('click', () => {
            this.balancesManager.exportToExcel();
        });

        // Обновление данных
        document.getElementById('refreshReceiptsBtn').addEventListener('click', () => {
            this.receiptsManager.updateTable();
        });

        document.getElementById('calculateInterestsBtn').addEventListener('click', () => {
            this.balancesManager.calculateInterests();
        });

        // Загрузка ИНН
        document.getElementById('innUploadBtn').addEventListener('click', () => {
            document.getElementById('innFileUpload').click();
        });

        document.getElementById('innFileUpload').addEventListener('change', (e) => {
            this.receiptsManager.loadINNData(e.target.files[0]);
            e.target.value = '';
        });

        // Загрузка данных по депозитам
        document.getElementById('depositUploadBtn').addEventListener('click', () => {
            document.getElementById('depositFileUpload').click();
        });

        document.getElementById('depositFileUpload').addEventListener('change', (e) => {
            this.balancesManager.loadDepositData(e.target.files[0]);
            e.target.value = '';
        });

        // Очистка поиска
        document.getElementById('clearSearchBtn').addEventListener('click', () => {
            this.receiptsManager.clearSearch();
        });

        // Поиск
        document.getElementById('searchReceipts').addEventListener('input', (e) => {
            this.receiptsManager.searchTransactions(e.target.value);
        });

        // Модальное окно для депозитов
        const modal = document.getElementById('depositModal');
        const closeModalButtons = document.querySelectorAll('.close-modal');

        closeModalButtons.forEach(button => {
            button.addEventListener('click', () => {
                modal.classList.remove('active');
            });
        });

        document.getElementById('saveDepositBtn').addEventListener('click', () => {
            this.balancesManager.saveDeposit();
        });

        // Клик вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });

        // ===== Обработчики для страницы сверки долгов =====
        this.setupDebtReconciliationListeners();

        // ===== Обработчики для страницы оплат поставщикам =====
        this.setupSupplierPaymentsListeners();

        // ===== Обработчики для страницы библиотеки =====
        this.setupLibraryListeners();
    }

    setupDebtReconciliationListeners() {
        // Выбор файла реестра ДЗ
        const selectDebtRegistryBtn = document.getElementById('selectDebtRegistryBtn');
        const debtRegistryFile = document.getElementById('debtRegistryFile');
        const debtRegistryDropArea = document.getElementById('debtRegistryDropArea');
        const debtRegistryFileInfo = document.getElementById('debtRegistryFileInfo');

        selectDebtRegistryBtn.addEventListener('click', () => {
            debtRegistryFile.click();
        });

        debtRegistryFile.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadDebtRegistryFile(e.target.files[0]);
            }
        });

        // Drag and drop для реестра ДЗ
        debtRegistryDropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            debtRegistryDropArea.style.borderColor = '#2563eb';
            debtRegistryDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        debtRegistryDropArea.addEventListener('dragleave', () => {
            debtRegistryDropArea.style.borderColor = '';
            debtRegistryDropArea.style.background = '';
        });

        debtRegistryDropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            debtRegistryDropArea.style.borderColor = '';
            debtRegistryDropArea.style.background = '';

            if (e.dataTransfer.files.length > 0) {
                this.loadDebtRegistryFile(e.dataTransfer.files[0]);
            }
        });

        // Выбор файла поступлений
        const selectReceiptsRegistryBtn = document.getElementById('selectReceiptsRegistryBtn');
        const receiptsRegistryFile = document.getElementById('receiptsRegistryFile');
        const receiptsRegistryDropArea = document.getElementById('receiptsRegistryDropArea');
        const receiptsRegistryFileInfo = document.getElementById('receiptsRegistryFileInfo');

        selectReceiptsRegistryBtn.addEventListener('click', () => {
            receiptsRegistryFile.click();
        });

        receiptsRegistryFile.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadReceiptsRegistryFile(e.target.files[0]);
            }
        });

        // Drag and drop для поступлений
        receiptsRegistryDropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            receiptsRegistryDropArea.style.borderColor = '#2563eb';
            receiptsRegistryDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        receiptsRegistryDropArea.addEventListener('dragleave', () => {
            receiptsRegistryDropArea.style.borderColor = '';
            receiptsRegistryDropArea.style.background = '';
        });

        receiptsRegistryDropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            receiptsRegistryDropArea.style.borderColor = '';
            receiptsRegistryDropArea.style.background = '';

            if (e.dataTransfer.files.length > 0) {
                this.loadReceiptsRegistryFile(e.dataTransfer.files[0]);
            }
        });

        // Кнопка сверки
        document.getElementById('reconcileBtn').addEventListener('click', () => {
            this.performReconciliation();
        });

        // Кнопка экспорта
        document.getElementById('exportReconciledBtn').addEventListener('click', () => {
            this.exportReconciledFile();
        });

        // Кнопка очистки
        document.getElementById('clearReconciliationBtn').addEventListener('click', () => {
            this.clearReconciliationData();
        });

        // Кнопка настроек сводных таблиц
        document.getElementById('summarySettingsBtn').addEventListener('click', () => {
            this.openSummarySettingsModal();
        });

        // Закрытие модального окна настроек
        const closeSummaryModal = document.getElementById('closeSummarySettingsModal');
        closeSummaryModal.addEventListener('click', () => {
            document.getElementById('summarySettingsModal').classList.remove('active');
        });

        // Сохранение настроек сводных
        document.getElementById('saveSummarySettingsBtn').addEventListener('click', () => {
            this.saveSummarySettings();
        });

        // Кнопка сохранения данных текущего дня
        document.getElementById('saveCurrentDayBtn').addEventListener('click', () => {
            this.saveCurrentDayData();
        });

        // Кнопка очистки данных предыдущего дня
        document.getElementById('clearPreviousDayDataBtn').addEventListener('click', () => {
            this.clearPreviousDayData();
        });

        // Загрузка данных предыдущего дня из Excel файла
        const loadPreviousDayFileBtn = document.getElementById('loadPreviousDayFileBtn');
        const previousDayFileInput = document.getElementById('previousDayFileInput');
        const previousDayFileInfo = document.getElementById('previousDayFileInfo');

        loadPreviousDayFileBtn.addEventListener('click', () => {
            previousDayFileInput.click();
        });

        previousDayFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                const result = await this.loadPreviousDayDataFromFile(file);

                if (result.success) {
                    // Обновляем информацию о файле
                    previousDayFileInfo.innerHTML = `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.count} ДТ, ${this.formatCurrency(result.total)})`;

                    // Обновляем таблицу
                    this.renderPreviousDayDataTable();

                    // Обновляем индикатор
                    this.updatePreviousDayIndicator();

                    this.showNotification(result.message, 'success');
                } else {
                    this.showNotification(result.message, 'error');
                }

                // Сбрасываем input, чтобы можно было загрузить тот же файл повторно
                e.target.value = '';
            }
        });

        // Загрузка файла СИ УАТ
        const selectSiUatBtn = document.getElementById('selectSiUatBtn');
        const siUatFileInput = document.getElementById('siUatFile');
        const siUatDropArea = document.getElementById('siUatDropArea');
        const siUatFileInfo = document.getElementById('siUatFileInfo');

        selectSiUatBtn.addEventListener('click', () => {
            siUatFileInput.click();
        });

        siUatFileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadSiUatFile(e.target.files[0]);
            }
        });

        siUatDropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            siUatDropArea.style.borderColor = '#2563eb';
            siUatDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        siUatDropArea.addEventListener('dragleave', () => {
            siUatDropArea.style.borderColor = '';
            siUatDropArea.style.background = '';
        });

        siUatDropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            siUatDropArea.style.borderColor = '';
            siUatDropArea.style.background = '';

            if (e.dataTransfer.files.length > 0) {
                this.loadSiUatFile(e.dataTransfer.files[0]);
            }
        });

        // Клик вне модального окна настроек
        window.addEventListener('click', (e) => {
            if (e.target === document.getElementById('summarySettingsModal')) {
                document.getElementById('summarySettingsModal').classList.remove('active');
            }
        });
    }

    // Настройка управления контрагентами
    setupContractorsManager() {
        const manageContractorsBtn = document.getElementById('manageContractorsBtn');
        const contractorsModal = document.getElementById('contractorsModal');
        const closeModal = document.getElementById('closeContractorsModal');
        const cancelBtn = document.getElementById('cancelContractorsBtn');
        const saveBtn = document.getElementById('saveContractorsBtn');
        const addBtn = document.getElementById('addContractorBtn');
        const contractorsList = document.getElementById('contractorsList');

        // Открытие модального окна
        manageContractorsBtn.addEventListener('click', () => {
            this.renderContractorsList();
            contractorsModal.classList.add('active');
        });

        // Закрытие модального окна
        const closeModalFn = () => {
            contractorsModal.classList.remove('active');
        };

        closeModal.addEventListener('click', closeModalFn);
        cancelBtn.addEventListener('click', closeModalFn);

        // Клик вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target === contractorsModal) {
                contractorsModal.classList.remove('active');
            }
        });

        // Добавление нового контрагента
        addBtn.addEventListener('click', () => {
            const contractor = prompt('Введите наименование контрагента:');
            if (contractor && contractor.trim() !== '') {
                if (this.debtManager.addTargetContractor(contractor)) {
                    this.renderContractorsList();
                    this.showNotification('Контрагент добавлен', 'success');
                } else {
                    this.showNotification('Контрагент уже существует или некорректное название', 'error');
                }
            }
        });

        // Сохранение изменений (фактически уже сохранено, но обновляем интерфейс)
        saveBtn.addEventListener('click', () => {
            contractorsModal.classList.remove('active');
            this.showNotification('Список контрагентов сохранен', 'success');
        });
    }

    // Отрисовка списка контрагентов в модальном окне
    renderContractorsList() {
        const contractorsList = document.getElementById('contractorsList');
        const contractors = this.debtManager.getTargetContractors();

        if (contractors.length === 0) {
            contractorsList.innerHTML = '<p class="empty-message">Список контрагентов пуст. Добавьте хотя бы одного.</p>';
            return;
        }

        let html = '';
        contractors.forEach(contractor => {
            html += `
                <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px; border-bottom: 1px solid var(--border);">
                    <span>${contractor}</span>
                    <button class="btn btn-danger btn-sm remove-contractor" data-contractor="${contractor}" style="padding: 4px 8px;">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            `;
        });

        contractorsList.innerHTML = html;

        // Добавляем обработчики для кнопок удаления
        contractorsList.querySelectorAll('.remove-contractor').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                const contractor = btn.dataset.contractor;
                if (confirm(`Удалить контрагента "${contractor}"?`)) {
                    if (this.debtManager.removeTargetContractor(contractor)) {
                        this.renderContractorsList();
                        this.showNotification('Контрагент удален', 'success');
                    }
                }
            });
        });
    }

    async loadDebtRegistryFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadDebtRegistryFile(file);
            document.getElementById('debtRegistryFileInfo').innerHTML =
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.message})`;

            // Активируем кнопку сверки, если оба файла загружены
            this.updateReconcileButtonState();

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    async loadReceiptsRegistryFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadReceiptsFile(file);
            document.getElementById('receiptsRegistryFileInfo').innerHTML =
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.message})`;

            // Активируем кнопку сверки, если оба файла загружены
            this.updateReconcileButtonState();

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    updateReconcileButtonState() {
        const stats = this.debtManager.getStats();
        const reconcileBtn = document.getElementById('reconcileBtn');
        reconcileBtn.disabled = !(stats.debtRows > 0 && stats.receiptsWithDates > 0);
    }

    performReconciliation() {
        this.showLoading();
        try {
            const result = this.debtManager.reconcile();
            if (result.success) {
                this.showReconciliationStats(this.debtManager.getStats());
                this.showReconciliationLog(this.debtManager.getProcessedLog());
                document.getElementById('exportReconciledBtn').disabled = false;
                document.getElementById('saveCurrentDayBtn').disabled = false;
                // Обновляем индикатор данных предыдущего дня
                this.updatePreviousDayIndicator();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка сверки:', error);
            this.showNotification('Ошибка при выполнении сверки', 'error');
        } finally {
            this.hideLoading();
        }
    }

    showReconciliationStats(stats) {
        document.getElementById('statTotalDocuments').textContent = stats.totalDocuments || 0;
        document.getElementById('statFoundDocuments').textContent = stats.foundDocuments || 0;
        document.getElementById('statUpdatedDocuments').textContent = stats.updatedDocuments || 0;
        document.getElementById('statReceiptsWithDates').textContent = stats.receiptsWithDates || 0;
        document.getElementById('reconciliationStats').style.display = 'block';
    }

    showReconciliationLog(log) {
        const logContainer = document.getElementById('logContainer');
        const logSection = document.getElementById('reconciliationLog');

        if (!log || log.length === 0) {
            logContainer.innerHTML = '<p class="empty-message">Нет обработанных документов</p>';
        } else {
            let html = '<table class="log-table"><tr><th>Документ</th><th>Действие</th><th>Дата</th><th>Сумма</th></tr>';
            log.slice(0, 50).forEach(item => {
                const dateStr = item.date ? new Date(item.date).toLocaleDateString('ru-RU') : '';
                html += `<tr>
                    <td>${item.documentName}</td>
                    <td><span class="badge badge-success">${item.action}</span></td>
                    <td>${dateStr}</td>
                    <td class="number-cell">${this.storage.formatNumber(item.amount || 0)}</td>
                </tr>`;
            });
            if (log.length > 50) {
                html += `<tr><td colspan="4" class="text-center">... и еще ${log.length - 50} записей</td></tr>`;
            }
            html += '</table>';
            logContainer.innerHTML = html;
        }

        logSection.style.display = 'block';
    }

    async exportReconciledFile() {
        const progressBar = document.getElementById('saveProgressBar');
        const progressFill = document.getElementById('saveProgressFill');
        const progressText = document.getElementById('saveProgressText');
        const exportBtn = document.getElementById('exportReconciledBtn');

        try {
            // Показываем прогресс-бар
            progressBar.style.display = 'block';
            exportBtn.disabled = true;
            progressFill.style.width = '10%';
            progressText.textContent = 'Подготовка данных...';

            // Анимация прогресса
            let progress = 10;
            const progressInterval = setInterval(() => {
                if (progress < 80) {
                    progress += Math.random() * 15;
                    if (progress > 80) progress = 80;
                    progressFill.style.width = progress + '%';
                    if (progress < 40) {
                        progressText.textContent = 'Обработка документов...';
                    } else if (progress < 60) {
                        progressText.textContent = 'Формирование отчёта...';
                    } else {
                        progressText.textContent = 'Создание файлов...';
                    }
                }
            }, 500);

            const result = await this.debtManager.exportToExcel();

            // Останавливаем анимацию
            clearInterval(progressInterval);
            progressFill.style.width = '100%';
            progressText.textContent = 'Готово!';

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else if (result.message) {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка при экспорте:', error);
            this.showNotification('Ошибка при сохранении файла', 'error');
        } finally {
            // Скрываем прогресс-бар через небольшую задержку
            setTimeout(() => {
                progressBar.style.display = 'none';
                progressFill.style.width = '0%';
                exportBtn.disabled = false;
            }, 1500);
        }
    }

    clearReconciliationData() {
        this.debtManager.clearData();
        document.getElementById('debtRegistryFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('receiptsRegistryFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('siUatFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('reconciliationStats').style.display = 'none';
        document.getElementById('reconciliationLog').style.display = 'none';
        document.getElementById('exportReconciledBtn').disabled = true;
        document.getElementById('reconcileBtn').disabled = true;
        document.getElementById('previousDayIndicator').style.display = 'none';
        this.showNotification('Данные очищены', 'info');
    }

    // Загрузка файла СИ УАТ
    async loadSiUatFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadSiUatFile(file);
            document.getElementById('siUatFileInfo').innerHTML =
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name}`;

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка загрузки файла СИ УАТ', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Загрузка данных предыдущего дня из Excel файла
    async loadPreviousDayDataFromFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadPreviousDayDataFromFile(file);
            const fileInfoEl = document.getElementById('previousDayFileInfo');

            if (result.success) {
                fileInfoEl.innerHTML =
                    `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.count} ДТ, ${this.formatCurrency(result.total)})`;
                this.showNotification(result.message, 'success');
            } else {
                fileInfoEl.innerHTML =
                    `<i class="fas fa-exclamation-circle" style="color: var(--danger);"></i> ${result.message}`;
                this.showNotification(result.message, 'error');
            }

            return result;
        } catch (error) {
            this.showNotification('Ошибка загрузки файла: ' + error.message, 'error');
            return { success: false, message: error.message };
        } finally {
            this.hideLoading();
        }
    }

    // Открытие модального окна настроек сводных
    openSummarySettingsModal() {
        // Заполняем поля сводных данных
        document.getElementById('summaryDtLegal').value = this.debtManager.summaryDT.legal || '';
        document.getElementById('summaryDtNotRecoverable').value = this.debtManager.summaryDT.notRecoverable || '';
        document.getElementById('summaryDtRecoverable').value = this.debtManager.summaryDT.recoverable || '';

        document.getElementById('summarySiuatTotalDebt').value = this.debtManager.summarySIUAT.totalDebt || '';
        document.getElementById('summarySiuatTotalOverdue').value = this.debtManager.summarySIUAT.totalOverdue || '';
        document.getElementById('summarySiuatLegal').value = this.debtManager.summarySIUAT.legal || '';
        document.getElementById('summarySiuatNotRecoverable').value = this.debtManager.summarySIUAT.notRecoverable || '';
        document.getElementById('summarySiuatRecoverable').value = this.debtManager.summarySIUAT.recoverable || '';

        // Заполняем дату предыдущего дня
        const previousInfo = this.debtManager.getPreviousDayData();
        document.getElementById('previousDayDateInput').value = previousInfo.date || '';

        // Заполняем таблицу данных предыдущего дня из localStorage
        this.renderPreviousDayDataTable();

        document.getElementById('summarySettingsModal').classList.add('active');
    }

    // Отрисовка таблицы данных предыдущего дня
    renderPreviousDayDataTable() {
        const tbody = document.getElementById('previousDayDataBody');
        const statsEl = document.getElementById('previousDayStats');
        const previousInfo = this.debtManager.getPreviousDayData();
        const previousData = previousInfo.data || previousInfo; // совместимость со старым форматом

        if (Object.keys(previousData).length === 0) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="2">Данные не загружены. Загрузите Excel файл.</td></tr>';
            statsEl.style.display = 'none';
            return;
        }

        // Показываем статистику
        statsEl.style.display = 'block';
        const dtCount = Object.keys(previousData).length;
        const totalAmount = Object.values(previousData).reduce((sum, val) => sum + val, 0);

        document.getElementById('previousDayDtCount').textContent = dtCount;
        document.getElementById('previousDayDtTotal').textContent = this.formatCurrency(totalAmount);

        // Таблица — только для просмотра (данные из файла, не редактируемые)
        let html = '';
        const sortedFilials = Object.keys(previousData).sort();

        sortedFilials.forEach(filial => {
            const amount = previousData[filial] || 0;
            html += `<tr>
                <td>${this.escapeHtml(filial)}</td>
                <td class="number-cell">${this.formatNumber(amount)}</td>
            </tr>`;
        });

        tbody.innerHTML = html;
    }

    // Форматирование валюты для отображения
    formatCurrency(amount) {
        return new Intl.NumberFormat('ru-RU', {
            style: 'currency',
            currency: 'RUB',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount);
    }

    // Форматирование числа
    formatNumber(num) {
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }

    // Очистка данных предыдущего дня
    clearPreviousDayData() {
        if (!confirm('Вы уверены, что хотите удалить все данные предыдущего дня?')) {
            return;
        }

        const result = this.debtManager.clearPreviousDayData();

        if (result.success) {
            // Сбрасываем информацию о файле
            document.getElementById('previousDayFileInfo').innerHTML = '<i class="fas fa-info-circle"></i> Файл не загружен';
            document.getElementById('previousDayFileInput').value = '';
            document.getElementById('previousDayDateInput').value = '';

            this.renderPreviousDayDataTable();
            this.updatePreviousDayIndicator();

            this.showNotification('Данные предыдущего дня очищены', 'success');
        } else {
            this.showNotification(result.message, 'error');
        }
    }

    // Сохранение данных текущего дня в localStorage для сравнения завтра
    saveCurrentDayData() {
        this.debtManager.saveCurrentDayData();
        this.updatePreviousDayIndicator();
        this.showNotification('Данные текущего дня сохранены для сравнения завтра', 'success');
    }

    // Сохранение настроек сводных
    saveSummarySettings() {
        // Сохраняем сводные данные ДТ
        this.debtManager.summaryDT = {
            legal: parseFloat(document.getElementById('summaryDtLegal').value) || 0,
            notRecoverable: parseFloat(document.getElementById('summaryDtNotRecoverable').value) || 0,
            recoverable: parseFloat(document.getElementById('summaryDtRecoverable').value) || 0
        };

        // Сохраняем сводные данные СИ УАТ
        this.debtManager.summarySIUAT = {
            totalDebt: parseFloat(document.getElementById('summarySiuatTotalDebt').value) || 0,
            totalOverdue: parseFloat(document.getElementById('summarySiuatTotalOverdue').value) || 0,
            legal: parseFloat(document.getElementById('summarySiuatLegal').value) || 0,
            notRecoverable: parseFloat(document.getElementById('summarySiuatNotRecoverable').value) || 0,
            recoverable: parseFloat(document.getElementById('summarySiuatRecoverable').value) || 0
        };

        // Сохраняем дату предыдущего дня
        const previousDate = document.getElementById('previousDayDateInput').value;
        const previousInfo = this.debtManager.getPreviousDayData();
        if (Object.keys(previousInfo.data || previousInfo).length > 0 || previousDate) {
            localStorage.setItem('previousDayDebt_manual', JSON.stringify({
                data: previousInfo.data || previousInfo,
                date: previousDate
            }));
        }

        // Данные предыдущего дня сохраняются автоматически при загрузке из файла
        // Здесь просто сохраняем сводные данные
        this.debtManager.saveSummaryData();

        // Закрываем модалку
        document.getElementById('summarySettingsModal').classList.remove('active');

        // Обновляем индикатор
        this.updatePreviousDayIndicator();

        this.showNotification('Настройки сводных таблиц сохранены', 'success');
    }

    // Обновление индикатора данных предыдущего дня
    updatePreviousDayIndicator() {
        const indicator = document.getElementById('previousDayIndicator');
        const indicatorText = document.getElementById('previousDayIndicatorText');
        const previousInfo = this.debtManager.getPreviousDayData();
        const previousData = previousInfo.data || previousInfo; // совместимость со старым форматом

        if (Object.keys(previousData).length > 0) {
            indicator.style.display = 'flex';
            indicator.className = 'previous-day-indicator loaded';
            const dateStr = previousInfo.date ? ` (${previousInfo.date})` : '';
            indicatorText.textContent = `Данные за предыдущий день загружены${dateStr} (${Object.keys(previousData).length} подразделений)`;
        } else {
            indicator.style.display = 'flex';
            indicator.className = 'previous-day-indicator warning';
            indicatorText.textContent = 'Данные за предыдущий день не заполнены. Откройте Настройки сводных.';
        }
    }

    // Экранирование HTML
    escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    handleFileSelect(files) {
        if (!files || files.length === 0) return;

        this.storage.addFiles(Array.from(files));
        this.updateFileList();
        this.updateStats();
    }

    updateFileList() {
        const fileListContent = document.getElementById('fileListContent');
        const fileListSection = document.getElementById('fileListSection');
        const statsSection = document.getElementById('statsSection');
        const files = this.storage.getFiles();

        // Показываем/скрываем секции в зависимости от наличия файлов
        if (files.length === 0) {
            fileListSection.style.display = 'none';
            statsSection.style.display = 'none';
            return;
        }

        // Показываем секции при наличии файлов
        fileListSection.style.display = 'block';
        if (this.storage.getTransactions().length > 0) {
            statsSection.style.display = 'block';
        }

        if (files.length === 0) {
            fileListContent.innerHTML = '<p class="empty-message">Файлы не загружены</p>';
            return;
        }

        let html = '';
        files.forEach((file) => {
            const size = file.size > 1024 * 1024
                ? `${(file.size / (1024 * 1024)).toFixed(2)} MB`
                : `${(file.size / 1024).toFixed(2)} KB`;

            html += `
                <div class="file-item">
                    <div class="file-name">${file.name}</div>
                    <div class="file-size">${size}</div>
                </div>
            `;
        });

        fileListContent.innerHTML = html;
    }

    async processFiles() {
        const files = this.storage.getFiles();
        if (files.length === 0) {
            alert('Пожалуйста, сначала загрузите файлы выписок');
            return;
        }

        this.showLoading();

        try {
            const parser = new BankStatementParser();
            const results = await parser.processFiles(files);

            // Сохраняем данные
            this.storage.setStatements(results.statements);
            this.storage.setTransactions(results.transactions);
            this.storage.setAccounts(results.accounts);

            // Обновляем интерфейс
            this.updateStats();
            this.receiptsManager.updateTable();
            this.balancesManager.updateTable();

            // Показываем уведомление
            this.showNotification(`Обработано ${files.length} файлов. Найдено ${results.transactions.length} операций.`, 'success');

        } catch (error) {
            console.error('Error processing files:', error);
            this.showNotification('Ошибка при обработке файлов', 'error');
        } finally {
            this.hideLoading();
        }
    }

    clearFiles() {
        if (confirm('Вы уверены, что хотите очистить список файлов?')) {
            this.storage.clearFiles();
            this.updateFileList();
            this.updateStats();
            this.receiptsManager.updateTable();
            this.balancesManager.updateTable();
        }
    }

    updateStats() {
        const files = this.storage.getFiles();
        const transactions = this.storage.getTransactions();
        const accounts = this.storage.getAccounts();

        document.getElementById('filesCount').textContent = files.length;
        document.getElementById('operationsCount').textContent = transactions.length;
        document.getElementById('accountsCount').textContent = Object.keys(accounts).length;

        // Показываем статистику только если есть обработанные данные
        const statsSection = document.getElementById('statsSection');
        if (transactions.length > 0) {
            statsSection.style.display = 'block';
        }
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

    showNotification(message, type = 'info') {
        // Создаем уведомление
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <div class="notification-content">
                <i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle'}"></i>
                <span>${message}</span>
            </div>
        `;

        document.body.appendChild(notification);

        // Удаляем уведомление через 5 секунд
        setTimeout(() => {
            notification.style.animation = 'fadeOut 0.3s ease';
            setTimeout(() => notification.remove(), 300);
        }, 5000);
    }

    loadCurrentPage() {
        this.updatePageData(this.currentPage);
    }

    // ===== ОБРАБОТЧИКИ ДЛЯ СТРАНИЦЫ ОПЛАТ ПОСТАВЩИКАМ =====
    setupSupplierPaymentsListeners() {
        // Загрузка файла
        const uploadBtn = document.getElementById('uploadSuppliersBtn');
        const fileInput = document.getElementById('suppliersFileUpload');

        uploadBtn.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                await this.handleSuppliersFileUpload(e.target.files[0]);
                e.target.value = '';
            }
        });

        // Переход в библиотеку
        document.getElementById('goToLibraryBtn').addEventListener('click', () => {
            this.switchPage('library');
        });

        // Закрытие модальных окон (только для депозитов и контрагентов дебиторки)
        document.querySelectorAll('.close-modal').forEach(btn => {
            btn.addEventListener('click', () => {
                btn.closest('.modal').classList.remove('active');
            });
        });

        // Закрытие по клику вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target.classList.contains('modal')) {
                e.target.classList.remove('active');
            }
        });
    }

    async handleSuppliersFileUpload(file) {
        this.showLoading();
        try {
            // Сохраняем оригинальный файл для экспорта
            this.originalSuppliersFile = file;

            const result = await this.supplierPayments.loadExcelFile(file);
            if (result.success) {
                this.renderSuppliersContent();
                this.updateSuppliersUI();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка загрузки файла поставщиков:', error);
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    renderSuppliersContent() {
        const container = document.getElementById('suppliersContent');
        const pivotTables = this.supplierPayments.getPivotTables();

        if (pivotTables.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <div class="empty-state-icon">
                        <i class="fas fa-cloud-upload-alt"></i>
                    </div>
                    <h3 class="empty-state-title">Загрузите Excel-файл с реестрами</h3>
                    <p class="empty-state-description">Загрузите Excel-файл с реестрами заявок для формирования сводных таблиц</p>
                </div>
            `;
            return;
        }

        let html = '';

        pivotTables.forEach((pivot, index) => {
            html += `
                <div style="background: var(--bg-card); border: 1px solid var(--border-light); border-radius: var(--radius-lg); padding: 24px; margin-bottom: 24px; box-shadow: var(--shadow-sm);">
                    <h3 style="margin-bottom: 16px; font-size: 16px; font-weight: 600;">
                        <i class="fas fa-table" style="color: var(--primary);"></i> ${pivot.sheetName}
                    </h3>
                    <div class="table-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Контрагент</th>
                                    ${pivot.pivotHeaders.map(h => `<th>${h}</th>`).join('')}
                                    <th>Итого</th>
                                    <th>Пояснение</th>
                                </tr>
                            </thead>
                            <tbody>
            `;

            let totals = {};
            pivot.pivotHeaders.forEach(h => totals[h] = 0);
            let grandTotal = 0;

            pivot.pivotData.forEach(row => {
                html += `<tr>`;
                html += `<td><strong>${row.contractor}</strong></td>`;

                pivot.pivotHeaders.forEach(h => {
                    const value = row[h] || 0;
                    totals[h] = (totals[h] || 0) + value;
                    html += `<td class="number-cell">${this.storage.formatNumber(value)}</td>`;
                });

                grandTotal += row.total;
                html += `<td class="number-cell" style="font-weight: 600;">${this.storage.formatNumber(row.total)}</td>`;
                html += `<td style="color: var(--text-secondary); font-size: 13px;">${row.explanation || ''}</td>`;
                html += `</tr>`;
            });

            // Итоговая строка
            html += `<tr style="background: var(--bg-tertiary); font-weight: 600;">`;
            html += `<td>ИТОГО</td>`;
            pivot.pivotHeaders.forEach(h => {
                html += `<td class="number-cell">${this.storage.formatNumber(totals[h] || 0)}</td>`;
            });
            html += `<td class="number-cell">${this.storage.formatNumber(grandTotal)}</td>`;
            html += `<td></td>`;
            html += `</tr>`;

            html += `
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        });

        // Кнопка экспорта
        html += `
            <div style="display: flex; justify-content: flex-end; gap: 12px; margin-top: 24px;">
                <button id="exportSuppliersBtn" class="btn btn-success">
                    <i class="fas fa-file-excel"></i> Экспорт в Excel
                </button>
            </div>
        `;

        container.innerHTML = html;

        // Обработчик экспорта
        document.getElementById('exportSuppliersBtn').addEventListener('click', () => {
            this.exportSuppliersToExcel();
        });
    }

    async exportSuppliersToExcel() {
        this.showLoading();
        try {
            const result = await this.supplierPayments.exportToExcel(this.originalSuppliersFile);
            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка экспорта:', error);
            this.showNotification('Ошибка при экспорте', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // ===== СТРАНИЦА БИБЛИОТЕКИ КОНТРАГЕНТОВ =====
    setupLibraryListeners() {
        // Инициализация фильтров
        this.libraryFilters = {
            name: new Set(),
            organization: new Set(),
            explanation: new Set()
        };
        this.librarySearchTerm = '';
        this.selectedRowId = null;

        // Добавление контрагента
        document.getElementById('libAddBtn').addEventListener('click', () => {
            this.openAddContractorModal();
        });

        // Импорт
        document.getElementById('libImportBtn').addEventListener('click', () => {
            document.getElementById('libraryImportFile').click();
        });

        document.getElementById('libraryImportFile').addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                await this.importLibraryFromExcel(e.target.files[0]);
                e.target.value = '';
            }
        });

        // Экспорт
        document.getElementById('libExportBtn').addEventListener('click', () => {
            this.contractorsLibrary.exportToExcel();
        });

        // Очистка
        document.getElementById('libClearBtn').addEventListener('click', () => {
            const result = this.contractorsLibrary.clearAll();
            if (result.success) {
                this.renderLibraryTable();
                this.updateLibraryStats();
                this.showNotification(result.message, 'info');
            }
        });

        // Поиск
        document.getElementById('librarySearch').addEventListener('input', (e) => {
            this.librarySearchTerm = e.target.value;
            this.renderLibraryTable();
        });

        // Очистка поиска
        document.getElementById('libClearSearchBtn').addEventListener('click', () => {
            document.getElementById('librarySearch').value = '';
            this.librarySearchTerm = '';
            this.renderLibraryTable();
        });

        // Кнопки фильтров в заголовках
        document.querySelectorAll('.filter-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const filterType = btn.dataset.filter;
                this.toggleFilterDropdown(filterType, btn);
            });
        });

        // Закрытие фильтров
        document.querySelectorAll('.filter-close').forEach(btn => {
            btn.addEventListener('click', () => {
                const filterType = btn.dataset.filter;
                this.closeFilterDropdown(filterType);
            });
        });

        // Поиск внутри фильтров
        document.querySelectorAll('[data-filter-search]').forEach(input => {
            input.addEventListener('input', (e) => {
                const filterType = e.target.dataset.filterSearch;
                this.filterOptionsSearch(filterType, e.target.value);
            });
        });

        // Чекбокс "Выбрать все"
        document.querySelectorAll('.select-all-cb').forEach(cb => {
            cb.addEventListener('change', (e) => {
                const filterType = e.target.dataset.filterSelect;
                this.toggleSelectAll(filterType, e.target.checked);
            });
        });

        // Применить фильтр
        document.querySelectorAll('[data-filter-apply]').forEach(btn => {
            btn.addEventListener('click', () => {
                const filterType = btn.dataset.filterApply;
                this.applyFilter(filterType);
            });
        });

        // Сбросить фильтр
        document.querySelectorAll('[data-filter-clear]').forEach(btn => {
            btn.addEventListener('click', () => {
                const filterType = btn.dataset.filterClear;
                this.clearFilter(filterType);
            });
        });

        // Сохранение контрагента
        document.getElementById('saveContractorBtn').addEventListener('click', () => {
            this.saveContractor();
        });

        // Закрытие фильтров при клике на оверлей
        document.getElementById('filterDropdowns').addEventListener('click', (e) => {
            if (e.target === document.getElementById('filterDropdowns')) {
                this.closeAllFilterDropdowns();
            }
        });

        // Закрытие фильтров при клике вне
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.filter-dropdown') && !e.target.closest('.filter-btn')) {
                this.closeAllFilterDropdowns();
            }
        });

        // Горячие клавиши
        document.addEventListener('keydown', (e) => {
            // Escape закрывает фильтры
            if (e.key === 'Escape') {
                this.closeAllFilterDropdowns();
            }
            // Enter открывает редактирование выбранной строки
            if (e.key === 'Enter' && this.selectedRowId) {
                this.openEditContractorModal(this.selectedRowId);
            }
        });
    }

    toggleFilterDropdown(filterType, btn) {
        const dropdown = document.getElementById(`filter${this.capitalizeFirst(filterType)}`);
        const isOpen = dropdown.style.display === 'block';

        // Закрываем все фильтры
        this.closeAllFilterDropdowns();

        if (!isOpen) {
            // Позиционируем фильтр под кнопкой
            const rect = btn.getBoundingClientRect();
            dropdown.style.left = rect.left + 'px';
            dropdown.style.top = (rect.bottom + 4) + 'px';
            dropdown.style.display = 'block';
            document.getElementById('filterDropdowns').style.display = 'flex';

            // Заполняем опции
            this.populateFilterOptions(filterType);

            // Подсвечиваем кнопку
            btn.classList.add('active');
        }
    }

    closeFilterDropdown(filterType) {
        const dropdown = document.getElementById(`filter${this.capitalizeFirst(filterType)}`);
        dropdown.style.display = 'none';
        document.querySelector(`.filter-btn[data-filter="${filterType}"]`)?.classList.remove('active');

        // Проверяем, есть ли открытые фильтры
        const anyOpen = document.querySelectorAll('.filter-dropdown[style*="block"]');
        if (anyOpen.length === 0) {
            document.getElementById('filterDropdowns').style.display = 'none';
        }
    }

    toggleSelectAll(filterType, checked) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const checkboxes = container.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(cb => {
            cb.checked = checked;
        });
    }

    updateSelectAllState(filterType) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const checkboxes = container.querySelectorAll('input[type="checkbox"]');
        const selectAllCb = document.querySelector(`.select-all-cb[data-filter-select="${filterType}"]`);

        if (checkboxes.length === 0) return;

        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        selectAllCb.checked = allChecked;
    }

    closeAllFilterDropdowns() {
        ['name', 'organization', 'explanation'].forEach(type => {
            this.closeFilterDropdown(type);
        });
    }

    populateFilterOptions(filterType) {
        const contractors = this.contractorsLibrary.getAll();
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);

        // Собираем уникальные значения
        const values = new Map();
        contractors.forEach(c => {
            const val = c[filterType] || '(пусто)';
            values.set(val, (values.get(val) || 0) + 1);
        });

        // Сортируем
        const sorted = Array.from(values.entries()).sort((a, b) => a[0].localeCompare(b[0]));

        // Получаем текущие выбранные
        const selected = this.libraryFilters[filterType];

        let html = '';
        sorted.forEach(([value, count]) => {
            const isChecked = selected.size === 0 || selected.has(value);
            html += `
                <label class="filter-option">
                    <input type="checkbox" value="${this.escapeHtml(value)}" ${isChecked ? 'checked' : ''}>
                    <span class="filter-option-label" title="${this.escapeHtml(value)}">${this.escapeHtml(value)}</span>
                    <span class="filter-option-count">${count}</span>
                </label>
            `;
        });

        container.innerHTML = html;

        // Обработчики чекбоксов
        container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.addEventListener('change', () => {
                this.updateSelectAllState(filterType);
            });
            cb.addEventListener('click', (e) => {
                e.stopPropagation();
            });
        });
    }

    filterOptionsSearch(filterType, searchTerm) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const options = container.querySelectorAll('.filter-option');
        const term = searchTerm.toLowerCase();

        options.forEach(option => {
            const label = option.querySelector('.filter-option-label').textContent.toLowerCase();
            option.style.display = label.includes(term) ? '' : 'none';
        });
    }

    applyFilter(filterType) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const checkboxes = container.querySelectorAll('input[type="checkbox"]');
        const selected = new Set();

        checkboxes.forEach(cb => {
            if (cb.checked) {
                selected.add(cb.value);
            }
        });

        // Если выбраны все — очищаем фильтр (показываем всё)
        const totalOptions = container.querySelectorAll('.filter-option').length;
        if (selected.size === totalOptions || selected.size === 0) {
            this.libraryFilters[filterType] = new Set();
        } else {
            this.libraryFilters[filterType] = selected;
        }

        this.closeFilterDropdown(filterType);
        this.renderLibraryTable();
    }

    clearFilter(filterType) {
        this.libraryFilters[filterType] = new Set();
        this.closeFilterDropdown(filterType);
        this.renderLibraryTable();
    }

    capitalizeFirst(str) {
        return str.charAt(0).toUpperCase() + str.slice(1);
    }

    escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    renderLibraryTable() {
        const tbody = document.getElementById('libraryTableBody');
        let contractors = this.contractorsLibrary.getAll();

        // Фильтрация по поиску
        if (this.librarySearchTerm) {
            const term = this.librarySearchTerm.toLowerCase();
            contractors = contractors.filter(c =>
                c.name.toLowerCase().includes(term) ||
                (c.organization && c.organization.toLowerCase().includes(term)) ||
                (c.explanation && c.explanation.toLowerCase().includes(term))
            );
        }

        // Фильтрация по чекбоксам
        for (const [filterType, selectedValues] of Object.entries(this.libraryFilters)) {
            if (selectedValues.size > 0) {
                contractors = contractors.filter(c => {
                    const val = c[filterType] || '(пусто)';
                    return selectedValues.has(val);
                });
            }
        }

        this.updateLibraryStats();

        if (contractors.length === 0) {
            tbody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="4">${this.librarySearchTerm || Object.values(this.libraryFilters).some(s => s.size > 0) ? 'Ничего не найдено' : 'Библиотека пуста'}</td>
                </tr>
            `;
            this.selectedRowId = null;
            return;
        }

        let html = '';
        contractors.forEach(contractor => {
            const isSelected = contractor.id === this.selectedRowId;
            html += `
                <tr data-id="${contractor.id}" ${isSelected ? 'class="selected"' : ''}>
                    <td><strong>${contractor.name}</strong></td>
                    <td>${contractor.organization || '<span style="color: var(--text-tertiary);">—</span>'}</td>
                    <td style="color: var(--text-secondary); font-size: 12px;">${contractor.explanation || '<span style="color: var(--text-tertiary);">—</span>'}</td>
                    <td class="actions-cell">
                        <button class="btn btn-sm btn-secondary edit-contractor" data-id="${contractor.id}" title="Редактировать">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-sm btn-danger delete-contractor" data-id="${contractor.id}" title="Удалить">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                </tr>
            `;
        });

        tbody.innerHTML = html;

        // Выделение строк по клику
        tbody.querySelectorAll('tr[data-id]').forEach(row => {
            // Одинарный клик — выделение
            row.addEventListener('click', (e) => {
                // Не выделяем если клик по кнопкам
                if (e.target.closest('.btn')) return;
                this.selectRow(row.dataset.id);
            });

            // Двойной клик — редактирование
            row.addEventListener('dblclick', (e) => {
                if (e.target.closest('.btn')) return;
                this.openEditContractorModal(row.dataset.id);
            });
        });

        // Обработчики редактирования
        tbody.querySelectorAll('.edit-contractor').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openEditContractorModal(btn.dataset.id);
            });
        });

        // Обработчики удаления
        tbody.querySelectorAll('.delete-contractor').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const contractor = this.contractorsLibrary.contractors.find(c => c.id === btn.dataset.id);
                if (contractor && confirm(`Удалить "${contractor.name}"?`)) {
                    this.contractorsLibrary.removeContractor(btn.dataset.id);
                    if (this.selectedRowId === btn.dataset.id) {
                        this.selectedRowId = null;
                    }
                    this.renderLibraryTable();
                    this.updateLibraryStats();
                    this.showNotification('Контрагент удалён', 'success');
                }
            });
        });
    }

    selectRow(id) {
        this.selectedRowId = id;

        // Обновляем визуальное выделение — только строки с data-id
        const tbody = document.getElementById('libraryTableBody');
        tbody.querySelectorAll('tr[data-id]').forEach(row => {
            row.classList.toggle('selected', row.dataset.id === id);
        });
    }

    openAddContractorModal() {
        document.getElementById('contractorModalTitle').textContent = 'Добавить контрагента';
        document.getElementById('contractorId').value = '';
        document.getElementById('contractorName').value = '';
        document.getElementById('contractorOrganization').value = '';
        document.getElementById('contractorExplanation').value = '';
        document.getElementById('addContractorModal').classList.add('active');

        // Фокус на первое поле
        setTimeout(() => document.getElementById('contractorName').focus(), 100);

        // Enter для сохранения
        document.getElementById('contractorForm').addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                this.saveContractor();
            }
        });
    }

    openEditContractorModal(id) {
        const contractor = this.contractorsLibrary.contractors.find(c => c.id === id);
        if (!contractor) return;

        document.getElementById('contractorModalTitle').textContent = 'Редактировать контрагента';
        document.getElementById('contractorId').value = contractor.id;
        document.getElementById('contractorName').value = contractor.name;
        document.getElementById('contractorOrganization').value = contractor.organization || '';
        document.getElementById('contractorExplanation').value = contractor.explanation || '';
        document.getElementById('addContractorModal').classList.add('active');

        // Фокус на первое поле
        setTimeout(() => document.getElementById('contractorName').focus(), 100);

        // Enter для сохранения
        document.getElementById('contractorForm').addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                this.saveContractor();
            }
        });
    }

    saveContractor() {
        const id = document.getElementById('contractorId').value;
        const name = document.getElementById('contractorName').value.trim();
        const organization = document.getElementById('contractorOrganization').value.trim();
        const explanation = document.getElementById('contractorExplanation').value.trim();

        if (!name) {
            this.showNotification('Введите наименование контрагента', 'error');
            return;
        }

        let result;
        if (id) {
            result = this.contractorsLibrary.updateContractor(id, { name, organization, explanation });
        } else {
            result = this.contractorsLibrary.addContractor({ name, organization, explanation });
        }

        if (result.success || result.updated) {
            document.getElementById('addContractorModal').classList.remove('active');
            this.renderLibraryTable();
            this.updateSuppliersUI();
            this.showNotification(result.message, 'success');
        } else {
            this.showNotification(result.message, 'error');
        }
    }

    async importLibraryFromExcel(file) {
        this.showLoading();
        try {
            const result = await this.contractorsLibrary.importFromExcel(file);
            if (result.success) {
                this.renderLibraryTable();
                this.updateSuppliersUI();
                // Показываем расширенное уведомление
                const message = `${result.message}`;
                this.showNotification(message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка импорта библиотеки:', error);
            this.showNotification('Ошибка импорта: ' + error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }
}

// Инициализация приложения после загрузки страницы
document.addEventListener('DOMContentLoaded', () => {
    window.app = new App();
});
