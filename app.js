// app.js - Основной файл приложения
class App {
    constructor() {
        this.currentPage = 'upload';
        this.storage = new StorageManager();
        this.receiptsManager = new ReceiptsManager(this.storage);
        this.balancesManager = new BalancesManager(this.storage);
        this.debtManager = new DebtReconciliationManager(this.storage);
        this.init();
    }

    init() {
        console.log('Initializing Financial Analysis App...');
        this.setupNavigation();
        this.setupEventListeners();
        this.setupContractorsManager();
        this.loadCurrentPage();
        this.updateStats();

        // Устанавливаем сегодняшнюю дату по умолчанию
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('calculationDate').value = today;

        console.log('App initialized successfully');
    }

    setupNavigation() {
        const navButtons = document.querySelectorAll('.nav-btn');
        const pages = document.querySelectorAll('.page');

        navButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                e.preventDefault();
                const targetPage = button.getAttribute('data-page');
                this.switchPage(targetPage);
            });
        });
    }

    switchPage(targetPage) {
        console.log('Switching to page:', targetPage);
        this.currentPage = targetPage;

        const navButtons = document.querySelectorAll('.nav-btn');
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

            // Обновляем данные на странице
            this.updatePageData(targetPage);
        }
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
                // При переходе на страницу сверки обновляем статистику, если есть данные
                this.updateReconciliationUI();
                break;
        }
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

    exportReconciledFile() {
        try {
            const result = this.debtManager.exportToExcel();
            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка при сохранении файла', 'error');
        }
    }

    clearReconciliationData() {
        this.debtManager.clearData();
        document.getElementById('debtRegistryFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('receiptsRegistryFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('reconciliationStats').style.display = 'none';
        document.getElementById('reconciliationLog').style.display = 'none';
        document.getElementById('exportReconciledBtn').disabled = true;
        document.getElementById('reconcileBtn').disabled = true;
        this.showNotification('Данные очищены', 'info');
    }

    handleFileSelect(files) {
        if (!files || files.length === 0) return;

        this.storage.addFiles(Array.from(files));
        this.updateFileList();
        this.updateStats();
    }

    updateFileList() {
        const fileListContent = document.getElementById('fileListContent');
        const files = this.storage.getFiles();

        if (files.length === 0) {
            fileListContent.innerHTML = '<p class="empty-message">Файлы не загружены</p>';
            return;
        }

        let html = '';
        files.forEach((file, index) => {
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
}

// Инициализация приложения после загрузки страницы
document.addEventListener('DOMContentLoaded', () => {
    window.app = new App();
});
