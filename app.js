// app.js - Основной файл приложения
class App {
    constructor() {
        this.currentPage = 'upload';
        this.storage = new StorageManager();
        this.receiptsManager = new ReceiptsManager(this.storage);
        this.balancesManager = new BalancesManager(this.storage);
        this.init();
    }

    init() {
        console.log('Initializing Financial Analysis App...');
        this.setupNavigation();
        this.setupEventListeners();
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
        }
    }

    setupEventListeners() {
        // Загрузка файлов
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

        // Drag and drop
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

        // Модальное окно
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
                <i class="fas fa-${type === 'success' ? 'check-circle' : 'exclamation-circle'}"></i>
                <span>${message}</span>
            </div>
        `;

        // Добавляем стили для уведомления
        if (!document.querySelector('#notification-styles')) {
            const styles = document.createElement('style');
            styles.id = 'notification-styles';
            styles.textContent = `
                .notification {
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    padding: 1rem 1.5rem;
                    border-radius: var(--radius);
                    background: white;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                    z-index: 10000;
                    animation: slideIn 0.3s ease;
                    max-width: 400px;
                }

                .notification-success {
                    border-left: 4px solid var(--success);
                }

                .notification-error {
                    border-left: 4px solid var(--error);
                }

                .notification-content {
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }

                .notification-content i {
                    font-size: 1.2rem;
                }

                .notification-success .notification-content i {
                    color: var(--success);
                }

                .notification-error .notification-content i {
                    color: var(--error);
                }

                @keyframes slideIn {
                    from { transform: translateX(100%); opacity: 0; }
                    to { transform: translateX(0); opacity: 1; }
                }
            `;
            document.head.appendChild(styles);
        }

        document.body.appendChild(notification);

        // Удаляем уведомление через 5 секунд
        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease';
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
