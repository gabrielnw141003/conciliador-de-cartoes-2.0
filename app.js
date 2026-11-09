'use strict';

document.addEventListener('DOMContentLoaded', () => {
    const elements = {
        process: document.getElementById('processar'),
        erpFile: document.getElementById('fileErp'),
        redeFile: document.getElementById('fileRede'),
        erpName: document.getElementById('arquivo-vendas'),
        redeName: document.getElementById('arquivo-rede'),
        message: document.getElementById('mensagem'),
        messageText: document.getElementById('mensagem-texto'),
        results: document.getElementById('resultado'),
        selectErp: document.getElementById('selecionarVendas'),
        selectRede: document.getElementById('selecionarRede'),
        badge: document.getElementById('result-badge')
    };

    const missingElement = Object.entries(elements)
        .find(([, element]) => !element);

    if (missingElement) {
        throw new Error(`Elemento obrigatório não encontrado: ${missingElement[0]}.`);
    }

    const initialResults = elements.results.innerHTML;

    elements.selectErp.addEventListener('click', () => elements.erpFile.click());
    elements.selectRede.addEventListener('click', () => elements.redeFile.click());

    elements.erpFile.addEventListener('change', () => {
        elements.erpName.value = elements.erpFile.files[0]?.name ?? '';
        resetResults();
        setMessage('Arquivo do ERP selecionado. Selecione o extrato da Rede.', '');
    });

    elements.redeFile.addEventListener('change', () => {
        elements.redeName.value = elements.redeFile.files[0]?.name ?? '';
        resetResults();
        setMessage('Extrato da Rede selecionado. Pronto para conferir.', '');
    });

    elements.process.addEventListener('click', async () => {
        const erpFile = elements.erpFile.files[0];
        const redeFile = elements.redeFile.files[0];

        if (!erpFile || !redeFile) {
            setMessage('Selecione os arquivos do ERP e da Rede antes de conferir.', 'error');
            return;
        }

        if (!window.XLSX) {
            setMessage('A biblioteca de planilhas não foi carregada. Verifique a conexão e recarregue a página.', 'error');
            return;
        }

        elements.process.disabled = true;
        setMessage('Lendo planilhas e processando a conciliação...', 'loading');

        try {
            const [erpRows, redeRows] = await Promise.all([
                readSpreadsheet(erpFile),
                readSpreadsheet(redeFile)
            ]);
            const report = reconcile(erpRows, redeRows);
            renderReport(report);
            elements.badge.textContent = 'ANÁLISE CONCLUÍDA';
            elements.badge.classList.add('complete');
            setMessage('Conferência concluída.', 'success');
        } catch (error) {
            console.error('Falha ao processar a conciliação:', error);
            setMessage(
                `Não foi possível concluir a conferência: ${error.message}`,
                'error'
            );
        } finally {
            elements.process.disabled = false;
        }
    });

    function setMessage(text, state) {
        elements.message.className = `feedback${state ? ` ${state}` : ''}`;
        elements.messageText.textContent = text;
    }

    function resetResults() {
        elements.results.innerHTML = initialResults;
        elements.badge.textContent = 'AGUARDANDO ANÁLISE';
        elements.badge.classList.remove('complete');
    }

    function normalizeText(value) {
        return String(value ?? '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase()
            .trim();
    }

    async function readSpreadsheet(file) {
        let workbook;

        try {
            workbook = window.XLSX.read(await file.arrayBuffer(), {
                type: 'array',
                cellDates: true
            });
        } catch (error) {
            const reason = error instanceof Error ? ` ${error.message}` : '';
            throw new Error(`O arquivo "${file.name}" não pôde ser lido.${reason}`);
        }

        const sheetName = workbook.SheetNames?.[0];
        const sheet = sheetName && workbook.Sheets[sheetName];

        if (!sheet) {
            throw new Error(`O arquivo "${file.name}" não contém uma planilha válida.`);
        }

        const rows = window.XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: '',
            raw: true
        });

        const headerTerms = [
            'nsu', 'cv', 'autoriz', 'comprovante', 'valor',
            'venda', 'total', 'data', 'r$ parcela',
            'vlr parcela', 'vl parcela', 'vlr venda', 'vl venda'
        ];
        let headerIndex = -1;
        let bestScore = 0;

        rows.slice(0, 20).forEach((row, index) => {
            const cells = row.map(normalizeText);
            const score = headerTerms.reduce(
                (sum, term) => sum + Number(cells.some(cell => cell.includes(term))),
                0
            );
            if (score > bestScore) {
                bestScore = score;
                headerIndex = index;
            }
        });

        if (headerIndex < 0) {
            throw new Error(`Não foi possível identificar o cabeçalho do arquivo "${file.name}".`);
        }

        const headers = rows[headerIndex].map(value => String(value ?? '').trim());
        const records = rows.slice(headerIndex + 1)
            .filter(row => row.some(value => String(value ?? '').trim() !== ''))
            .map(row => {
                const record = {};
                headers.forEach((header, index) => {
                    if (!header) return;
                    let key = header;
                    let duplicate = 2;
                    while (Object.hasOwn(record, key)) {
                        key = `${header} (${duplicate})`;
                        duplicate += 1;
                    }
                    record[key] = row[index] ?? '';
                });
                return record;
            })
            .filter(record => Object.keys(record).length > 0);

        if (records.length === 0) {
            throw new Error(`O arquivo "${file.name}" não contém vendas após o cabeçalho.`);
        }

        return records;
    }

    function findColumn(records, terms, excludedTerms = []) {
        const columns = Object.keys(records[0] ?? {});
        return columns.find(column => {
            const name = normalizeText(column);
            return terms.some(term => name.includes(term))
                && !excludedTerms.some(term => name.includes(term));
        });
    }

    function findValueColumn(records, source) {
        const priorities = source === 'ERP'
            ? [
                'r$ parcela', 'vlr. parcela', 'vlr parcela', 'vl parcela',
                'valor da parcela', 'valor parcela', 'r$ venda', 'vlr. venda',
                'vlr venda', 'vl venda', 'valor da venda', 'valor venda',
                'r$ total', 'vlr. total', 'vlr total', 'valor total', 'valor bruto',
                'valor'
            ]
            : [
                'valor da venda original', 'valor venda original', 'valor original',
                'valor da venda atualizado', 'valor venda atualizado', 'valor atualizado',
                'valor bruto', 'valor da venda', 'valor venda', 'valor'
            ];
        const excluded = ['taxa', 'desconto', 'percentual', 'liquido', 'recebido', 'receber'];
        const columns = Object.keys(records[0] ?? {});
        const normalizeColumn = column => normalizeText(column)
            .replace(/[.]/g, '')
            .replace(/\s+/g, ' ');

        for (const priority of priorities) {
            const column = columns.find(candidate => {
                const name = normalizeColumn(candidate);
                const term = normalizeColumn(priority);
                return (name === term || name.includes(term))
                    && !excluded.some(term => name.includes(term));
            });
            if (column) return column;
        }
        return undefined;
    }

    function parseCents(value) {
        if (typeof value === 'number') {
            return Number.isFinite(value) ? Math.round(value * 100) : NaN;
        }
        if (value === null || value === undefined || String(value).trim() === '') {
            return NaN;
        }

        let text = String(value).trim();
        const negative = /^\s*-/.test(text) || /^\(.*\)$/.test(text);
        text = text.replace(/[^\d,.]/g, '');

        const comma = text.lastIndexOf(',');
        const dot = text.lastIndexOf('.');
        if (comma >= 0 && dot >= 0) {
            text = comma > dot
                ? text.replace(/\./g, '').replace(',', '.')
                : text.replace(/,/g, '');
        } else if (comma >= 0 || dot >= 0) {
            const separator = comma >= 0 ? ',' : '.';
            const digitsAfterSeparator = text.length - text.lastIndexOf(separator) - 1;
            if (digitsAfterSeparator === 1 || digitsAfterSeparator === 2) {
                text = text.replace(separator, '.');
            } else {
                text = text.replace(/[,.]/g, '');
            }
        }

        const amount = Number(text);
        if (!Number.isFinite(amount)) return NaN;
        return Math.round((negative ? -amount : amount) * 100);
    }

    function normalizeNsu(value) {
        if (value === null || value === undefined) return '';
        return String(value).trim().replace(/\.0$/, '').replace(/\D/g, '').replace(/^0+/, '');
    }

    function extractNsu(record) {
        const columns = Object.keys(record).filter(column => {
            const name = normalizeText(column).replace(/[^a-z0-9]/g, '');
            if (name.includes('preautorizado')) return false;
            return ['nsu', 'cv', 'autoriz', 'comprovante'].some(term => name.includes(term));
        });

        for (const column of columns) {
            const nsu = normalizeNsu(record[column]);
            if (nsu) return nsu;
        }
        return '';
    }

    function formatType(value) {
        const text = String(value ?? '').trim();
        const normalized = normalizeText(text);
        if (normalized.includes('deb')) return 'Débito';
        if (normalized.includes('cred')) return 'Crédito';
        return text;
    }

    function reconcile(erpRecords, redeRecords) {
        const erpValueColumn = findValueColumn(erpRecords, 'ERP');
        const redeValueColumn = findValueColumn(redeRecords, 'Rede');
        if (!erpValueColumn) {
            const columns = Object.keys(erpRecords[0] ?? {}).join(', ');
            throw new Error(
                `Não encontrei a coluna de valor no relatório do ERP. ` +
                `Cabeçalhos detectados: ${columns || 'nenhum'}.`
            );
        }
        if (!redeValueColumn) {
            const columns = Object.keys(redeRecords[0] ?? {}).join(', ');
            throw new Error(
                `Não encontrei a coluna de valor no extrato da Rede. ` +
                `Cabeçalhos detectados: ${columns || 'nenhum'}.`
            );
        }

        const statusColumn = findColumn(redeRecords, ['status', 'situacao']);
        const redeValidRecords = redeRecords.filter(record => {
            if (!statusColumn || !record[statusColumn]) return true;
            const status = normalizeText(record[statusColumn]);
            return !['cancel', 'rejeit', 'negad', 'estorn']
                .some(term => status.includes(term));
        });

        const toSale = (record, valueColumn, source) => {
            const cents = parseCents(record[valueColumn]);
            if (!Number.isFinite(cents)) {
                throw new Error(`Valor inválido na coluna "${valueColumn}" do arquivo ${source}.`);
            }
            return {
                record,
                nsu: extractNsu(record),
                cents
            };
        };

        const sales = erpRecords.map(record => ({
            ...toSale(record, erpValueColumn, 'ERP'),
            status: '',
            redeCents: null,
            type: ''
        }));
        const redeSales = redeValidRecords.map(record => ({
            ...toSale(record, redeValueColumn, 'Rede'),
            matched: false
        }));

        const erpTotals = new Map();
        sales.forEach(sale => {
            if (sale.nsu) {
                erpTotals.set(sale.nsu, (erpTotals.get(sale.nsu) ?? 0) + sale.cents);
            }
        });

        const redeGroups = new Map();
        redeSales.forEach(sale => {
            if (!sale.nsu) return;
            const group = redeGroups.get(sale.nsu) ?? { cents: 0, sales: [] };
            group.cents += sale.cents;
            group.sales.push(sale);
            redeGroups.set(sale.nsu, group);
        });

        const docColumn = findColumn(erpRecords, ['doc', 'documento', 'nota', 'nf']);
        const erpDateColumn = findColumn(erpRecords, ['data', 'emissao']);
        const erpTypeColumn = findColumn(erpRecords, ['forma', 'pgto', 'pagamento']);
        const redeDateColumn = findColumn(redeRecords, ['data']);
        const redeTypeColumn = findColumn(redeRecords, ['modalidade', 'tipo']);

        sales.forEach(sale => {
            if (!sale.nsu) {
                sale.status = 'NSU não informado';
                return;
            }

            const group = redeGroups.get(sale.nsu);
            if (!group) {
                sale.status = 'Não encontrado na Rede';
                return;
            }

            sale.redeCents = group.cents;
            sale.type = formatType(
                group.sales.map(item => redeTypeColumn && item.record[redeTypeColumn]).find(Boolean)
                ?? ''
            );
            sale.status = Math.abs(erpTotals.get(sale.nsu) - group.cents) <= 5
                ? 'Conciliada'
                : 'Divergência de valor';
            group.sales.forEach(item => {
                item.matched = true;
            });
        });

        return {
            sales,
            extras: redeSales.filter(sale => !sale.matched),
            columns: {
                doc: docColumn,
                erpDate: erpDateColumn,
                erpType: erpTypeColumn,
                redeDate: redeDateColumn,
                redeType: redeTypeColumn
            }
        };
    }

    function renderReport(report) {
        const groups = [
            {
                title: 'Vendas conciliadas',
                kind: 'good',
                items: report.sales.filter(sale => sale.status === 'Conciliada')
            },
            {
                title: 'Divergências de valor',
                kind: 'warning',
                items: report.sales.filter(sale => sale.status === 'Divergência de valor')
            },
            {
                title: 'Vendas não encontradas na Rede',
                kind: 'bad',
                items: report.sales.filter(sale => sale.status === 'Não encontrado na Rede')
            },
            {
                title: 'Vendas sem NSU',
                kind: 'warning',
                items: report.sales.filter(sale => sale.status === 'NSU não informado')
            },
            {
                title: 'Transações sobrando na Rede',
                kind: 'warning',
                items: report.extras,
                isRede: true
            }
        ];

        const reconciledCount = groups[0].items.length;
        const discrepancyCount = groups[1].items.length;
        const pendingCount = groups.slice(2).reduce((sum, group) => sum + group.items.length, 0);
        const summary = document.createElement('div');
        summary.className = 'summary-grid';
        summary.append(
            createSummaryCard('Vendas no ERP', report.sales.length, ''),
            createSummaryCard('Conciliadas', reconciledCount, 'good'),
            createSummaryCard('Divergências', discrepancyCount, 'warning'),
            createSummaryCard('Pendências', pendingCount, 'bad')
        );

        elements.results.replaceChildren(summary);
        groups.forEach(group => {
            elements.results.append(createResultSection(group, report.columns));
        });

        const footnote = document.createElement('p');
        footnote.className = 'audit-footnote';
        footnote.textContent = 'A conciliação compara os totais por NSU, com tolerância de R$ 0,05. Transações canceladas, rejeitadas ou estornadas são desconsideradas.';
        elements.results.append(footnote);
    }

    function createSummaryCard(title, value, kind) {
        const card = document.createElement('article');
        card.className = `summary-card${kind ? ` ${kind}` : ''}`;

        const top = document.createElement('div');
        top.className = 'summary-top';
        top.append(document.createTextNode(title));

        const icon = document.createElement('span');
        icon.className = 'summary-icon';
        icon.setAttribute('aria-hidden', 'true');
        icon.textContent = kind === 'good' ? '✓' : kind === 'warning' ? '!' : kind === 'bad' ? '×' : '#';
        top.append(icon);

        const number = document.createElement('strong');
        number.className = 'summary-value';
        number.textContent = String(value);

        card.append(top, number);
        return card;
    }

    function createResultSection(group, columns) {
        const section = document.createElement('details');
        section.className = `result-section ${group.kind}`;
        section.open = group.items.length > 0;

        const summary = document.createElement('summary');
        const titleWrap = document.createElement('span');
        titleWrap.className = 'section-title-wrap';

        const indicator = document.createElement('span');
        indicator.className = 'section-indicator';
        indicator.setAttribute('aria-hidden', 'true');
        titleWrap.append(indicator, document.createTextNode(group.title));

        const count = document.createElement('span');
        count.className = 'section-count';
        count.textContent = String(group.items.length);

        const chevron = document.createElement('span');
        chevron.className = 'chevron';
        chevron.setAttribute('aria-hidden', 'true');
        chevron.textContent = '⌄';
        summary.append(titleWrap, count, chevron);

        const list = document.createElement('div');
        list.className = 'result-list';
        if (group.items.length === 0) {
            const empty = document.createElement('p');
            empty.className = 'no-items';
            empty.textContent = 'Nenhum registro nesta categoria.';
            list.append(empty);
        } else {
            group.items.forEach(item => {
                list.append(createResultItem(item, group.isRede, columns));
            });
        }

        section.append(summary, list);
        return section;
    }

    function createResultItem(item, isRede, columns) {
        const card = document.createElement('article');
        card.className = 'result-item';

        const top = document.createElement('div');
        top.className = 'item-top';

        const title = document.createElement('span');
        title.className = 'item-title';
        const documentNumber = !isRede && columns.doc
            ? item.record[columns.doc]
            : '';
        title.textContent = documentNumber
            ? `Documento ${displayValue(documentNumber)}`
            : isRede ? 'Transação da Rede' : `NSU ${item.nsu || 'não informado'}`;

        const status = document.createElement('span');
        status.className = 'item-status';
        status.textContent = isRede ? 'Sem correspondência no ERP' : item.status;
        top.append(title, status);

        const fields = document.createElement('div');
        fields.className = 'item-fields';
        let note;
        if (!isRede && item.nsu) addField(fields, 'NSU', item.nsu);

        const dateColumn = isRede ? columns.redeDate : columns.erpDate;
        if (dateColumn && item.record[dateColumn] !== '') {
            addField(fields, 'Data', displayValue(item.record[dateColumn]));
        }

        if (isRede) {
            addField(fields, 'Valor na Rede', formatMoney(item.cents));
            const type = columns.redeType && formatType(item.record[columns.redeType]);
            if (type) addField(fields, 'Modalidade', type);
        } else {
            addField(fields, 'Valor ERP', formatMoney(item.cents));
            if (item.redeCents !== null) {
                addField(fields, 'Total na Rede', formatMoney(item.redeCents));
            }
            if (item.type) addField(fields, 'Modalidade', item.type);
            if (item.status === 'Divergência de valor') {
                note = document.createElement('p');
                note.className = 'item-note';
                note.textContent = 'Os totais das vendas com este NSU diferem entre o ERP e a Rede.';
            }
            if (item.status === 'Não encontrado na Rede' && columns.erpType) {
                const type = formatType(item.record[columns.erpType]);
                if (type) addField(fields, 'Forma de pagamento', type);
            }
        }

        card.append(top, fields);
        if (note) card.append(note);
        return card;
    }

    function addField(container, label, value) {
        const field = document.createElement('span');
        field.className = 'item-field';
        const strong = document.createElement('strong');
        strong.textContent = value;
        field.append(document.createTextNode(label), strong);
        container.append(field);
    }

    function displayValue(value) {
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
            return new Intl.DateTimeFormat('pt-BR').format(value);
        }
        return String(value ?? '').trim();
    }

    function formatMoney(cents) {
        return new Intl.NumberFormat('pt-BR', {
            style: 'currency',
            currency: 'BRL'
        }).format(cents / 100);
    }
});
