// parser.js - Парсер банковских выписок
// Конфигурация загружается из storage (БД), без хардкода в нормализации имён
class BankStatementParser {
    constructor() {
        // Вся конфигурация загружается с сервера/API через storage
        this.accountMapping = this.loadAccountMapping();
        this.companyPatterns = [];   // [{pattern, canonical, match_type}] из БД
        this._ready = false;
    }

    async init(storage) {
        // Берём конфигурацию из storage (уже загружена с сервера)
        const allAccounts = storage.getAccounts();
        this.companyPatterns = storage.getCompanyAliases();
        this.storage = storage;
        this._ready = true;
        console.log('✅ BankStatementParser: конфигурация загружена из storage');
        console.log(`   companyAliases: ${this.companyPatterns.length}, accounts: ${Object.keys(allAccounts).length}`);
    }

    loadAccountMapping() {
        console.log("Загрузка маппинга счетов из localStorage");
        const stored = localStorage.getItem('accountMapping');
        if (stored) {
            try {
                return JSON.parse(stored);
            } catch (e) {
                console.error('Ошибка загрузки маппинга из localStorage', e);
            }
        }
        // Стандартный маппинг (полный список) — fallback если ни localStorage, ни сервер не доступны
        return {
            "40702810900000004317": { company: "Сервис-Интегратор ООО", bank: "ВБРР" },
            "40702810300000011971": { company: "Сервис-Интегратор ООО", bank: "МКБ" },
            "40702810907700000421": { company: "Сервис-Интегратор ООО", bank: "БКС" },
            "40702810400000204768": { company: "Сервис-Интегратор ООО", bank: "ПСБ" },
            "40702810040000071672": { company: "Сервис-Интегратор ООО", bank: "Сбер" },
            "40702810404800000145": { company: "Сервис-Интегратор ООО", bank: "ВТБ" },
            "40702810040000022168": { company: "Сервис-Интегратор ООО", bank: "Сбер" },
            "40702810900000189310": { company: "Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810800000189601": { company: "Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810500000211743": { company: "Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810200000223730": { company: "Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810700990012381": { company: "Сервис-Интегратор ООО", bank: "МИБ" },
            "40702810240000080065": { company: "Сервис-Интегратор ООО", bank: "Сбер" },
            "40702810612010866225": { company: "Сервис-Интегратор ООО", bank: "Совкомбанк" },
            "40702810701300050818": { company: "Сервис-Интегратор ООО", bank: "Альфа" },
            "40702810001360001709": { company: "Сервис-Интегратор ООО", bank: "Ингосстрах" },
            "40702810000000011018": { company: "Сервис-Интегратор ООО", bank: "СДМ" },
            "40702810014900002747": { company: "Сервис-Интегратор ООО", bank: "Синара" },
            "40702810777700083889": { company: "Сервис-Интегратор ООО", bank: "Дело" },
            "40702810800000084832": { company: "Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810000000147197": { company: "Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810600000009460": { company: "Сервис-Интегратор ООО", bank: "РЕАЛИСТ" },
            "40702810400000199295": { company: "СИ УАТ ООО", bank: "ГПБ" },
            "40702810805010002132": { company: "СИ УАТ ООО", bank: "МКБ" },
            "40702810612010694918": { company: "СИ УАТ ООО", bank: "Совкомбанк" },
            "40702810200790000026": { company: "СИ УАТ ООО", bank: "Аверс" },
            "40702810003000156608": { company: "СИ УАТ ООО", bank: "ПСБ" },
            "40702810900000102708": { company: "СИ УАТ ООО", bank: "ГПБ" },
            "40702810500249213086": { company: "СИ УАТ ООО", bank: "ВТБ" },
            "40702810740000405629": { company: "СИ УАТ ООО", bank: "Сбер" },
            "40702810800000300877": { company: "Сервис-Интегратор Логистика ООО", bank: "ПСБ" },
            "40702810340000082125": { company: "Сервис-Интегратор Логистика ООО", bank: "Сбер" },
            "40702810500000141745": { company: "Сервис-Интегратор УТ ООО", bank: "ГПБ" },
            "40702810577700204635": { company: "Сервис-Интегратор УТ ООО", bank: "Дело" },
            "40702810340000106836": { company: "Сервис-Интегратор УТ ООО", bank: "Сбер" },
            "40702810112010694913": { company: "Сервис-Интегратор УТ ООО", bank: "Совкомбанк" },
            "40702810100760006507": { company: "Сервис-Интегратор УТ ООО", bank: "МКБ" },
            "40702810125620007380": { company: "Сервис-Интегратор УТ ООО", bank: "ВТБ" },
            "40702810500000009494": { company: "Сервис-Интегратор Сахалин ООО", bank: "СДМ" },
            "40702810100190001583": { company: "Сервис-Интегратор Сахалин ООО", bank: "МКБ" },
            "40702810240000071676": { company: "Сервис-Интегратор Сахалин ООО", bank: "Сбер" },
            "40702810504800000566": { company: "Сервис-Интегратор Сахалин ООО", bank: "ВТБ" },
            "40702810100990012143": { company: "СОИР ООО", bank: "МИБ" },
            "40702810700000001892": { company: "СОИР ООО", bank: "ГПБ" },
            "40702810404800000297": { company: "СОИР ООО", bank: "ВТБ" },
            "40702810412010126770": { company: "СОИР ООО", bank: "Совкомбанк" },
            "40702810240000407651": { company: "Сервис ЦМ ООО", bank: "Сбер" },
            "40702810024840001102": { company: "Сервис ЦМ ООО", bank: "ВТБ" },
            "40702810240000097197": { company: "Управляющая компания Сервис-Интегратор ООО", bank: "Сбер" },
            "40702810924840000960": { company: "Управляющая компания Сервис-Интегратор ООО", bank: "ВТБ" },
            "40702810100000125365": { company: "Управляющая компания Сервис-Интегратор ООО", bank: "ГПБ" },
            "40702810040000409079": { company: "Управляющая компания Сервис-Интегратор ООО", bank: "Сбер" },
            "40702810124840002315": { company: "Сервис-Интегратор Арктика ООО", bank: "ВТБ" },
            "40701810540000401219": { company: "Сервис-Интегратор АО", bank: "Сбер" },
            "40702810000000157491": { company: "Сервис-Интегратор АО", bank: "ГПБ" },
            "40701810424841000004": { company: "Сервис-Интегратор АО", bank: "ВТБ" },
            "40702810014900002734": { company: "Сервис-Интегратор АО", bank: "Синара" },
            "40701810212010391926": { company: "Сервис-Интегратор АО", bank: "Совкомбанк" }
        };
    }

    saveAccountMapping(mapping) {
        this.accountMapping = mapping;
        localStorage.setItem('accountMapping', JSON.stringify(mapping));
    }

    async processFiles(files) {
        const statements = [];
        const allTransactions = [];
        const accounts = {};

        for (const file of files) {
            try {
                const content = await this.readFile(file);
                const parsed = this.parseStatement(content, file.name);

                statements.push({
                    filename: file.name,
                    content: content,
                    account: parsed.account,
                    bank: parsed.bank,
                    date: parsed.date,
                    transactions: parsed.transactions
                });

                allTransactions.push(...parsed.transactions.map(t => ({
                    ...t,
                    sourceFile: file.name
                })));

                if (parsed.account) {
                    accounts[parsed.account] = {
                        company: parsed.company,
                        bank: parsed.bank,
                        balance: parsed.balance,
                        date: parsed.date
                    };
                }

            } catch (error) {
                console.error(`Error processing file ${file.name}:`, error);
                throw new Error(`Ошибка обработки файла ${file.name}: ${error.message}`);
            }
        }

        return {
            statements,
            transactions: allTransactions,
            accounts
        };
    }

    readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error('Ошибка чтения файла'));
            reader.readAsText(file, 'Windows-1251');
        });
    }

    parseStatement(content, filename) {
        const lines = content.split('\n');
        let account = '';
        let bank = '';
        let date = '';
        let balance = null;
        let company = '';
        const transactions = [];

        let currentTransaction = null;
        let inDocumentSection = false;
        let inAccountSection = false;

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();

            // Определение банка
            if (!bank) {
                bank = this.detectBank(line);
            }

            // Секция счета
            if (line === 'СекцияРасчСчет') {
                inAccountSection = true;
                continue;
            }

            if (inAccountSection) {
                if (line.startsWith('РасчСчет=')) {
                    account = line.split('=')[1]?.trim();
                    company = this.getCompanyByAccount(account);
                    const mappedBank = this.getBankByAccount(account);
                    if (mappedBank) {
                        bank = mappedBank;
                    }
                } else if (line.startsWith('ДатаКонца=')) {
                    date = line.split('=')[1]?.trim();
                } else if (line.startsWith('КонечныйОстаток=')) {
                    balance = parseFloat(line.split('=')[1]?.replace(',', '.') || 0);
                } else if (line === 'КонецРасчСчет') {
                    inAccountSection = false;
                }
                if (line.startsWith('ДатаНачала=') && !date) {
                    date = line.split('=')[1]?.trim();
                }
            }

            // Секция документа
            if (line.startsWith('СекцияДокумент=')) {
                inDocumentSection = true;
                currentTransaction = {
                    date: '',
                    number: '',
                    amount: 0,
                    payer: '',
                    payerINN: '',
                    payerAccount: '',
                    payerBank: '',
                    recipient: '',
                    recipientAccount: '',
                    recipientBank: '',
                    purpose: '',
                    direction: ''
                };
                continue;
            }

            if (line === 'КонецДокумента' && currentTransaction) {
                inDocumentSection = false;

                if (!currentTransaction.direction) {
                    this.determineTransactionDirection(currentTransaction, account);
                }

                this.processTransaction(currentTransaction, account, company);
                transactions.push(currentTransaction);
                currentTransaction = null;
                continue;
            }

            if (inDocumentSection && currentTransaction) {
                this.parseDocumentLine(line, currentTransaction);
            }
        }

        // Если счет не определен из секции, пробуем из имени файла
        if (!account) {
            const accountMatch = filename.match(/\d{20}/);
            if (accountMatch) {
                account = accountMatch[0];
                company = this.getCompanyByAccount(account);
                const mappedBank = this.getBankByAccount(account);
                if (mappedBank) {
                    bank = mappedBank;
                }
            }
        }

        // Если банк не определен, пробуем из имени файла
        if (!bank) {
            bank = this.detectBankFromFilename(filename);
        }

        // Если банк все еще не определен, пробуем получить из маппинга
        if (!bank && account) {
            bank = this.getBankByAccount(account);
        }

        return {
            account,
            bank,
            date,
            balance,
            company,
            transactions
        };
    }

    parseDocumentLine(line, transaction) {
        const [key, ...valueParts] = line.split('=');
        if (!key || valueParts.length === 0) return;

        const value = valueParts.join('=').trim();

        switch(key) {
            case 'Дата': transaction.date = value; break;
            case 'Номер': transaction.number = value; break;
            case 'Сумма': transaction.amount = parseFloat(value.replace(',', '.')) || 0; break;
            case 'НазначениеПлатежа': transaction.purpose = value; break;

            case 'Плательщик':
            case 'Плательщик1':
                transaction.payer = this.cleanCompanyName(value);
                break;

            case 'ПлательщикИНН':
                transaction.payerINN = value;
                break;

            case 'ПлательщикСчет':
            case 'ПлательщикРасчСчет':
                transaction.payerAccount = value;
                break;

            case 'ПлательщикБанк':
            case 'ПлательщикБанк1':
            case 'БанкПлательщика':
                transaction.payerBank = value;
                break;

            case 'Получатель':
            case 'Получатель1':
                transaction.recipient = this.cleanCompanyName(value);
                break;

            case 'ПолучательСчет':
            case 'ПолучательРасчСчет':
                transaction.recipientAccount = value;
                break;

            case 'ПолучательБанк':
            case 'ПолучательБанк1':
            case 'БанкПолучателя':
                transaction.recipientBank = value;
                break;

            case 'ДатаПоступило':
            case 'Дебит':
                if (value && value.trim()) transaction.direction = 'incoming';
                break;

            case 'ДатаСписано':
            case 'Кредит':
                if (value && value.trim()) transaction.direction = 'outgoing';
                break;
        }
    }

    cleanCompanyName(name) {
        if (!name) return '';
        name = name.replace(/^ИНН\s+\d+\s+/, '');
        name = name.replace(/^["']+|["']+$/g, '');
        return name.trim();
    }

    determineTransactionDirection(transaction, account) {
        if (transaction.recipientAccount === account ||
            this.isOurCompany(transaction.recipient)) {
            transaction.direction = 'incoming';
        } else if (transaction.payerAccount === account ||
                   this.isOurCompany(transaction.payer)) {
            transaction.direction = 'outgoing';
        } else if (transaction.recipientAccount &&
                   this.accountMapping[transaction.recipientAccount]) {
            transaction.direction = 'incoming';
        } else if (transaction.payerAccount &&
                   this.accountMapping[transaction.payerAccount]) {
            transaction.direction = 'outgoing';
        }
    }

    processTransaction(transaction, account, company) {
        if (transaction.direction === "incoming") {
            transaction.ourAccount = account || transaction.recipientAccount;
            transaction.ourCompany = company;
            if (!transaction.ourCompany && transaction.recipient) {
                transaction.ourCompany = this.normalizeCompanyName(transaction.recipient);
            }
            transaction.ourBank = this.normalizeBankName(transaction.recipientBank);
            transaction.counterCompany = transaction.payer;
            transaction.counterAccount = transaction.payerAccount;
        } else if (transaction.direction === "outgoing") {
            transaction.ourAccount = account || transaction.payerAccount;
            transaction.ourCompany = company;
            if (!transaction.ourCompany && transaction.payer) {
                transaction.ourCompany = this.normalizeCompanyName(transaction.payer);
            }
            transaction.ourBank = this.normalizeBankName(transaction.payerBank);
            transaction.counterCompany = transaction.recipient;
            transaction.counterAccount = transaction.recipientAccount;
        }

        if (transaction.ourCompany) {
            transaction.ourCompany = this.normalizeCompanyName(transaction.ourCompany);
        }
        if (transaction.counterCompany) {
            transaction.counterCompany = this.normalizeCompanyName(transaction.counterCompany);
        }
    }

    // ── Нормализация имён компаний через company_aliases из БД ──
    normalizeCompanyName(name) {
        if (!name) return name;

        // Проверяем синонимы из БД (company_aliases)
        for (const alias of this.companyPatterns) {
            let matches = false;
            if (alias.match_type === 'exact') {
                matches = name.toUpperCase() === alias.pattern.toUpperCase();
            } else if (alias.match_type === 'regex') {
                try { matches = new RegExp(alias.pattern, 'i').test(name); } catch(e) {}
            } else {
                // 'contains' (по умолчанию)
                matches = name.toUpperCase().includes(alias.pattern.toUpperCase());
            }
            if (matches) return alias.canonical;
        }
        return name;
    }

    isOurCompany(companyName) {
        if (!companyName) return false;
        for (const alias of this.companyPatterns) {
            let matches = false;
            const upper = companyName.toUpperCase();
            if (alias.match_type === 'exact') {
                matches = upper === alias.pattern.toUpperCase();
            } else if (alias.match_type === 'regex') {
                try { matches = new RegExp(alias.pattern, 'i').test(companyName); } catch(e) {}
            } else {
                matches = upper.includes(alias.pattern.toUpperCase());
            }
            if (matches) return true;
        }
        // Если синонимов нет — проверяем по маппингу счетов
        const upperName = companyName.toUpperCase();
        for (const [account, info] of Object.entries(this.accountMapping)) {
            if (info.company && info.company.toUpperCase().includes(upperName)) return true;
        }
        return false;
    }

    getCompanyByAccount(account) {
        if (!account) return '';
        const cleanAccount = account.replace(/\s/g, '');
        return this.accountMapping[cleanAccount]?.company || '';
    }

    getBankByAccount(account) {
        if (!account) return '';
        const cleanAccount = account.replace(/\s/g, '');
        return this.accountMapping[cleanAccount]?.bank || '';
    }

    // ── Определение банка ──────────────────────────────────
    detectBank(line) {
        const upper = line.toUpperCase();
        const banks = [
            ['ПСБ', /ПСБ|PSBCORPORATE/],
            ['Сбер', /СБЕР|СБЕРКАЗНАЧЕЙСТВО/],
            ['СДМ', /СДМ|ИНН\s*7729395092/],
            ['МКБ', /МКБ|МОСКОВСКИЙ КРЕДИТНЫЙ БАНК/],
            ['ВТБ', /ВТБ/],
            ['ГПБ', /ГПБ/],
            ['БКС', /БКС/],
            ['Синара', /СИНАРА/],
            ['Совкомбанк', /СОВКОМБАНК/],
            ['Аверс', /АВЕРС/],
            ['Альфа', /АЛЬФА/],
            ['ВБРР', /ВБРР/],
            ['МИБ', /МИБ/],
            ['Дело', /ДЕЛО/],
            ['Ингосстрах', /ИНГОССТРАХ/]
        ];
        for (const [name, pattern] of banks) {
            if (pattern.test(upper)) return name;
        }
        return '';
    }

    detectBankFromFilename(filename) {
        const upper = filename.toUpperCase();
        const map = { 'ПСБ': 'ПСБ', 'PSB': 'ПСБ', 'СБЕР': 'Сбер', 'SBER': 'Сбер',
            'СДМ': 'СДМ', 'SDM': 'СДМ', 'МКБ': 'МКБ', 'MKB': 'МКБ',
            'ВТБ': 'ВТБ', 'VTB': 'ВТБ', 'ГПБ': 'ГПБ', 'GPB': 'ГПБ',
            'БКС': 'БКС', 'BCS': 'БКС', 'СИНАРА': 'Синара', 'СОВКОМБАНК': 'Совкомбанк',
            'АВЕРС': 'Аверс', 'АЛЬФА': 'Альфа', 'ВБРР': 'ВБРР', 'МИБ': 'МИБ', 'ДЕЛО': 'Дело' };
        for (const [key, val] of Object.entries(map)) {
            if (upper.includes(key)) return val;
        }
        return '';
    }

    normalizeBankName(bankName) {
        if (!bankName) return '';
        const upper = bankName.toUpperCase();
        const map = { 'АВЕРС': 'Аверс', 'АЛЬФА-БАНК': 'Альфа', 'АЛЬФА': 'Альфа',
            'БКС': 'БКС Банк', 'ВБРР': 'ВБРР', 'ВТБ': 'ВТБ', 'ГПБ': 'ГПБ',
            'ДЕЛО': 'Дело', 'ИНГОССТРАХ': 'Ингосстрах', 'МКБ': 'МОСКОВСКИЙ КРЕДИТНЫЙ БАНК',
            'МИБ': 'МИБ', 'ПСБ': 'ПСБ', 'СБЕРБАНК': 'Сбербанк', 'СБЕР': 'Сбербанк',
            'СДМ': 'СДМ-БАНК', 'СИНАРА': 'Синара', 'СОВКОМБАНК': 'Совкомбанк' };
        for (const [k, v] of Object.entries(map)) {
            if (upper.includes(k)) return v;
        }
        return bankName;
    }

    formatNumber(num) {
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }

    // ===== УПРАВЛЕНИЕ МАППИНГОМ СЧЕТОВ (СЕРВЕР + LOCALSTORAGE) =====

    async loadMappingFromServer() {
        try {
            console.log('🔄 Загрузка маппинга счетов с сервера...');
            const response = await fetch('/api/account-mapping', {
                method: 'GET',
                headers: { 'Content-Type': 'application/json' }
            });

            if (!response.ok) {
                console.warn('⚠️ Сервер недоступен, используем локальный маппинг');
                return this.accountMapping;
            }

            const result = await response.json();
            if (!result.success || !result.data) {
                console.warn('⚠️ Сервер вернул пустой маппинг');
                return this.accountMapping;
            }

            const serverMapping = {};
            for (const item of result.data) {
                if (item.is_active && item.account_number) {
                    serverMapping[item.account_number] = {
                        company: item.company_name,
                        bank: item.bank_name || ''
                    };
                }
            }

            const mergedMapping = { ...this.accountMapping, ...serverMapping };
            this.accountMapping = mergedMapping;
            this.saveAccountMapping(mergedMapping);
            this._serverMappingLoaded = true;

            console.log(`✅ Загружено ${Object.keys(serverMapping).length} записей с сервера, всего: ${Object.keys(mergedMapping).length}`);
            return mergedMapping;

        } catch (error) {
            console.warn('⚠️ Ошибка загрузки с сервера:', error.message);
            return this.accountMapping;
        }
    }

    async syncMappingToServer() {
        try {
            console.log('🔄 Синхронизация маппинга на сервер...');

            const mappingList = [];
            for (const [accountNumber, info] of Object.entries(this.accountMapping)) {
                mappingList.push({
                    account_number: accountNumber,
                    company_name: info.company || '',
                    bank_name: info.bank || ''
                });
            }

            const response = await fetch('/api/account-mapping/sync', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ mapping: mappingList })
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            const result = await response.json();
            console.log(`✅ Синхронизация: добавлено ${result.added}, обновлено ${result.updated}`);
            return result;

        } catch (error) {
            console.warn('⚠️ Ошибка синхронизации с сервером:', error.message);
            return { success: false, error: error.message };
        }
    }

    getMappingForUI() {
        const result = [];
        for (const [accountNumber, info] of Object.entries(this.accountMapping)) {
            result.push({
                account_number: accountNumber,
                company_name: info.company || '',
                bank_name: info.bank || ''
            });
        }
        result.sort((a, b) => a.company_name.localeCompare(b.company_name) || a.account_number.localeCompare(b.account_number));
        return result;
    }
}