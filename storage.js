// storage.js - Управление хранением данных (упрощенная версия для теста)
class StorageManager {
    constructor() {
        this.files = [];
        this.statements = [];
        this.transactions = [];
        this.accounts = {};
        this.innData = {};
        this.depositData = {};
    }

    // Управление файлами
    addFiles(newFiles) {
        const existingNames = this.files.map(f => f.name);
        const uniqueFiles = newFiles.filter(f => !existingNames.includes(f.name));
        this.files = [...this.files, ...uniqueFiles];
    }

    getFiles() { return [...this.files]; }
    clearFiles() { 
        this.files = [];
        this.statements = [];
        this.transactions = [];
        this.accounts = {};
    }

    // Управление выписками
    setStatements(statements) { this.statements = statements; }
    getStatements() { return [...this.statements]; }

    // Управление транзакциями
    setTransactions(transactions) { this.transactions = transactions; }
    getTransactions() { return [...this.transactions]; }
    getIncomingTransactions() { return this.transactions.filter(t => t.direction === 'incoming'); }
    getOutgoingTransactions() { return this.transactions.filter(t => t.direction === 'outgoing'); }

    // Управление счетами
    setAccounts(accounts) { this.accounts = accounts; }
    getAccounts() { return { ...this.accounts }; }
    updateAccount(account, data) { this.accounts[account] = { ...this.accounts[account], ...data }; }

    // Управление ИНН
    setINNData(data) { this.innData = data; }
    getINNData() { return { ...this.innData }; }
    getCompanyByINN(inn) { return this.innData[inn]; }

    // Управление депозитами
    setDepositData(data) { this.depositData = data; }
    getDepositData() { return { ...this.depositData }; }
    getDepositForAccount(account) { return this.depositData[account]; }
    setDepositForAccount(account, data) { this.depositData[account] = data; }

    // Вспомогательные методы
    formatCurrency(amount) {
        return new Intl.NumberFormat('ru-RU', {
            style: 'currency',
            currency: 'RUB',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount);
    }

    formatNumber(num) {
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }

    parseDate(dateStr) {
        console.log('Парсинг даты:', dateStr);
        if (!dateStr) return new Date();
        
        // Формат DD.MM.YYYY
        const parts = dateStr.split('.');
        if (parts.length === 3) {
            const date = new Date(parts[2], parts[1] - 1, parts[0]);
            console.log('Парсинг DD.MM.YYYY результат:', date);
            return date;
        }
        
        // Формат YYYY-MM-DD
        const isoMatch = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (isoMatch) {
            const date = new Date(isoMatch[1], isoMatch[2] - 1, isoMatch[3]);
            console.log('Парсинг YYYY-MM-DD результат:', date);
            return date;
        }
        
        // Формат MM/DD/YYYY или другие
        const date = new Date(dateStr);
        console.log('Парсинг через new Date результат:', date);
        
        // Проверяем, что дата валидна
        if (isNaN(date.getTime())) {
            console.log('Неверный формат даты, возвращаем текущую дату');
            return new Date();
        }
        
        return date;
    }

    getDaysBetween(startDate, endDate) {
        console.log('Расчет дней между:', { startDate, endDate });
        const start = this.parseDate(startDate);
        const end = this.parseDate(endDate);
        console.log('Даты после парсинга:', { start, end });
        const diffTime = Math.abs(end - start);
        const days = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        console.log('Результат дней:', days);
        return days;
    }

    calculateInterest(amount, rate, days) {
        console.log('Расчет процентов:', { amount, rate, days });
        if (!amount || !rate || !days) return 0;
        
        // Проверяем типы данных
        amount = parseFloat(amount);
        rate = parseFloat(rate);
        days = parseInt(days);
        
        console.log('После преобразования:', { amount, rate, days });
        
        const result = (amount * rate * days) / (100 * 365);
        console.log('Результат расчета:', result);
        
        return result;
    }
}
