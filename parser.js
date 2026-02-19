// parser.js - Парсер банковских выписок
class BankStatementParser {
    constructor() {
        // Загружаем маппинг из localStorage или используем стандартный
        this.accountMapping = this.loadAccountMapping();

        this.companyPatterns = [
            /ООО\s*["']?СЕРВИС-ИНТЕГРАТОР["']?/i,
            /СИ УАТ/i,
            /СЕРВИС ЦМ/i,
            /СЕРВИС-ИНТЕГРАТОР УТ/i,
            /СЕРВИС-ИНТЕГРАТОР САХАЛИН/i,
            /СЕРВИС-ИНТЕГРАТОР ЛОГИСТИКА/i,
            /СОИР/i,
            /УПРАВЛЯЮЩАЯ КОМПАНИЯ СЕРВИС-ИНТЕГРАТОР/i,
            /СЕРВИС-ИНТЕГРАТОР АО/i
        ];
    }

    loadAccountMapping() {
        console.log("Загрузка маппинга из localStorage");
        const stored = localStorage.getItem('accountMapping');
        if (stored) {
            try {
                return JSON.parse(stored);
            } catch (e) {
                console.error('Ошибка загрузки маппинга из localStorage', e);
            }
        }
        // Стандартный маппинг (полный список)
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

                // Добавляем транзакции
                allTransactions.push(...parsed.transactions.map(t => ({
                    ...t,
                    sourceFile: file.name
                })));

                // Обновляем информацию о счетах
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

            // Определение банка по содержимому
            if (!bank) {
                if (line.includes('ПСБ') || line.includes('PSBCorporate')) bank = 'ПСБ';
                else if (line.includes('Сбер') || line.includes('СберКазначейство')) bank = 'Сбер';
                else if (line.includes('СДМ') || line.includes('ИНН 7729395092')) bank = 'СДМ';
                else if (line.includes('МКБ') || line.includes('МОСКОВСКИЙ КРЕДИТНЫЙ БАНК')) bank = 'МКБ';
                else if (line.includes('ВТБ')) bank = 'ВТБ';
                else if (line.includes('ГПБ')) bank = 'ГПБ';
                else if (line.includes('БКС')) bank = 'БКС';
                else if (line.includes('Синара')) bank = 'Синара';
                else if (line.includes('СОВКОМБАНК')) bank = 'Совкомбанк';
                else if (line.includes('АВЕРС')) bank = 'Аверс';
                else if (line.includes('АЛЬФА')) bank = 'Альфа';
                else if (line.includes('ВБРР')) bank = 'ВБРР';
                else if (line.includes('МИБ')) bank = 'МИБ';
                else if (line.includes('ДЕЛО')) bank = 'Дело';
                else if (line.includes('ИНГОССТРАХ')) bank = 'Ингосстрах';
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
                    // Получаем банк из маппинга, если он там есть
                    const mappedBank = this.getBankByAccount(account);
                    if (mappedBank) {
                        bank = mappedBank;
                    }
                } else if (line.startsWith('ДатаКонца=')) {
                    // Берем дату конца периода как дату выписки - это самая свежая дата
                    date = line.split('=')[1]?.trim();
                } else if (line.startsWith('КонечныйОстаток=')) {
                    balance = parseFloat(line.split('=')[1]?.replace(',', '.') || 0);
                } else if (line === 'КонецРасчСчет') {
                    inAccountSection = false;
                }
                // Также проверяем ДатаНачала= на случай, если ДатаКонца= отсутствует
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
                    direction: '' // 'incoming' или 'outgoing'
                };
                continue;
            }

            if (line === 'КонецДокумента' && currentTransaction) {
                inDocumentSection = false;

                // Определяем направление платежа
                if (!currentTransaction.direction) {
                    this.determineTransactionDirection(currentTransaction, account);
                }

                // Обрабатываем транзакцию
                this.processTransaction(currentTransaction, account, company);
                transactions.push(currentTransaction);
                currentTransaction = null;
                continue;
            }

            // Парсинг полей документа
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
                // Получаем банк из маппинга для счета из имени файла
                const mappedBank = this.getBankByAccount(account);
                if (mappedBank) {
                    bank = mappedBank;
                }
            }
        }

        // Если банк не определен, пробуем из имени файла
        if (!bank) {
            if (filename.includes('ПСБ') || filename.includes('PSB')) bank = 'ПСБ';
            else if (filename.includes('Сбер') || filename.includes('SBER')) bank = 'Сбер';
            else if (filename.includes('СДМ') || filename.includes('SDM')) bank = 'СДМ';
            else if (filename.includes('МКБ') || filename.includes('MKB')) bank = 'МКБ';
            else if (filename.includes('ВТБ') || filename.includes('VTB')) bank = 'ВТБ';
            else if (filename.includes('ГПБ') || filename.includes('GPB')) bank = 'ГПБ';
            else if (filename.includes('БКС') || filename.includes('BCS')) bank = 'БКС';
            else if (filename.includes('Синара')) bank = 'Синара';
            else if (filename.includes('Совкомбанк')) bank = 'Совкомбанк';
            else if (filename.includes('Аверс')) bank = 'Аверс';
            else if (filename.includes('Альфа')) bank = 'Альфа';
            else if (filename.includes('ВБРР')) bank = 'ВБРР';
            else if (filename.includes('МИБ')) bank = 'МИБ';
            else if (filename.includes('Дело')) bank = 'Дело';
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

        // Удаление ИНН из начала
        if (name.startsWith('ИНН ')) {
            name = name.replace(/^ИНН\s+\d+\s+/, '');
        }

        // Удаление лишних кавычек
        name = name.replace(/^["']+|["']+$/g, '');

        return name.trim();
    }

    determineTransactionDirection(transaction, account) {
        console.log("Определяем направление для транзакции:", transaction);
        console.log("Счет из выписки:", account);
        // Если получатель - наша компания, это входящий платеж
        if (transaction.recipientAccount === account ||
            this.isOurCompany(transaction.recipient)) {
            transaction.direction = 'incoming';
        }
        // Если плательщик - наша компания, это исходящий платеж
        else if (transaction.payerAccount === account ||
                 this.isOurCompany(transaction.payer)) {
            transaction.direction = 'outgoing';
        }
        // Эвристика: если счет получателя есть в нашей базе
        else if (transaction.recipientAccount &&
                 this.accountMapping[transaction.recipientAccount]) {
            transaction.direction = 'incoming';
        }
        // Эвристика: если счет плательщика есть в нашей базе
        else if (transaction.payerAccount &&
                 this.accountMapping[transaction.payerAccount]) {
            transaction.direction = 'outgoing';
        }
    }

    processTransaction(transaction, account, company) {
        // Определяем нашу компанию и контрагента
        if (transaction.direction === "incoming") {
            // Используем переданные account и company (из маппинга или секции счета)
            transaction.ourAccount = account || transaction.recipientAccount;
            transaction.ourCompany = company;

            // Если компания не определена через маппинг, пытаемся определить из получателя
            if (!transaction.ourCompany && transaction.recipient) {
                transaction.ourCompany = this.normalizeCompanyName(transaction.recipient);
            }

            // Банк получателя (наш банк)
            transaction.ourBank = this.normalizeBankName(transaction.recipientBank);

            // Контрагент (плательщик)
            transaction.counterCompany = transaction.payer;
            transaction.counterAccount = transaction.payerAccount;
        } else if (transaction.direction === "outgoing") {
            transaction.ourAccount = account || transaction.payerAccount;
            transaction.ourCompany = company;

            // Если компания не определена через маппинг, пытаемся определить из плательщика
            if (!transaction.ourCompany && transaction.payer) {
                transaction.ourCompany = this.normalizeCompanyName(transaction.payer);
            }

            // Банк отправителя (наш банк)
            transaction.ourBank = this.normalizeBankName(transaction.payerBank);

            // Контрагент (получатель)
            transaction.counterCompany = transaction.recipient;
            transaction.counterAccount = transaction.recipientAccount;
        }

        // Нормализуем названия компаний
        if (transaction.ourCompany) {
            transaction.ourCompany = this.normalizeCompanyName(transaction.ourCompany);
        }
        if (transaction.counterCompany) {
            transaction.counterCompany = this.normalizeCompanyName(transaction.counterCompany);
        }
    }

    getCompanyByAccount(account) {
        console.log("getCompanyByAccount:", account);
        if (!account) return '';
        const cleanAccount = account.replace(/\s/g, '');
        return this.accountMapping[cleanAccount]?.company || '';
    }

    getBankByAccount(account) {
        if (!account) return '';
        const cleanAccount = account.replace(/\s/g, '');
        return this.accountMapping[cleanAccount]?.bank || '';
    }

    isOurCompany(companyName) {
        if (!companyName) return false;

        for (const pattern of this.companyPatterns) {
            if (pattern.test(companyName)) {
                return true;
            }
        }

        return false;
    }

    normalizeCompanyName(name) {
        if (!name) return name;

        // Приводим к верхнему регистру для унификации сравнения
        const upperName = name.toUpperCase();

        // Проверяем по шаблонам компаний
        if (/СИ УАТ|СИУАТ/.test(upperName)) return "СИ УАТ ООО";
        if (/СЕРВИС ЦМ|СЕРВИСЦМ/.test(upperName)) return "Сервис ЦМ ООО";
        if (/УПРАВЛЯЮЩАЯ КОМПАНИЯ/.test(upperName)) return "Управляющая компания Сервис-Интегратор ООО";
        if (/СОИР/.test(upperName)) return "СОИР ООО";
        if (/СЕРВИСНОЕ ОБСЛУЖИВАНИЕ И РЕМОНТ/.test(upperName)) return "СОИР ООО";
        if (/СЕРВИС-ИНТЕГРАТОР УТ/.test(upperName)) return "Сервис-Интегратор УТ ООО";
        if (/СЕРВИС-ИНТЕГРАТОР САХАЛИН/.test(upperName)) return "Сервис-Интегратор Сахалин ООО";
        if (/СЕРВИС-ИНТЕГРАТОР ЛОГИСТИКА/.test(upperName)) return "Сервис-Интегратор Логистика ООО";
        if (/СЕРВИС-ИНТЕГРАТОР АО/.test(upperName)) return "Сервис-Интегратор АО";
        if (/СЕРВИС-ИНТЕграТОР АРКТИКА/.test(upperName)) return "Сервис-Интегратор Арктика";
        if (/СЕРВИС-ИНТЕГРАТОР/.test(upperName)) return "Сервис-Интегратор ООО";

        return name;
    }

    normalizeBankName(bankName) {
        if (!bankName) return bankName;

        const upperName = bankName.toUpperCase();

        // Маппинг вариантов на стандартные названия
        if (/АВЕРС/.test(upperName)) return "Аверс";
        if (/АЛЬФА-БАНК|АЛЬФА/.test(upperName)) return "Альфа";
        if (/БКС/.test(upperName)) return "БКС Банк";
        if (/ВБРР/.test(upperName)) return "ВБРР";
        if (/ВТБ/.test(upperName)) return "ВТБ";
        if (/ГПБ/.test(upperName)) return "ГПБ";
        if (/ДЕЛО/.test(upperName)) return "Дело";
        if (/ИНГОССТРАХ/.test(upperName)) return "Ингосстрах";
        if (/МЕТАЛЛИНВЕСТБАНК/.test(upperName)) return "Металлинвестбанк";
        if (/МОСКОВСКИЙ КРЕДИТНЫЙ БАНК|МКБ/.test(upperName)) return "МОСКОВСКИЙ КРЕДИТНЫЙ БАНК";
        if (/МИБ/.test(upperName)) return "МИБ";
        if (/ПСБ/.test(upperName)) return "ПСБ";
        if (/СБЕРБАНК|СБЕР/.test(upperName)) return "Сбербанк";
        if (/СДМ/.test(upperName)) return "СДМ-БАНК";
        if (/СИНАРА/.test(upperName)) return "Синара";
        if (/СОВКОМБАНК/.test(upperName)) return "Совкомбанк";

        return bankName;
    }

    formatNumber(num) {
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }
}
