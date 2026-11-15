import * as XLSX from './vendor/xlsx.mjs';
const normalize = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
export const columns = ['Status', 'Account', 'Sequence', 'Core amount (as supplied)', 'Core date', 'Core sheet', 'Core row', 'Reason', 'D Navi action code', 'Direction (as supplied)', 'D Navi sheet', 'D Navi rows'];
const key = (r) => `${r.account}:${r.sequence}`;
export function extract(data, kind) {
    if (data.byteLength > 25 * 1024 * 1024)
        throw new Error('Each workbook must be smaller than 25 MB.');
    let book;
    try {
        book = XLSX.read(data, { type: 'array', cellFormula: false, cellHTML: false, cellText: false });
    }
    catch {
        throw new Error('Unable to read workbook. Select a valid, unencrypted .xlsx file.');
    }
    const required = kind === 'dn'
        ? { account: 'pan', sequence: 'systemtracenumber', code: 'actioncode' }
        : { account: 'accountno', sequence: 'seqno', amount: 'amount', date: 'date' };
    const candidates = [];
    for (const name of book.SheetNames) {
        const sheet = book.Sheets[name];
        if (!sheet['!ref'])
            continue;
        const range = XLSX.utils.decode_range(sheet['!ref']);
        if (range.e.r > 250000 || range.e.c > 200)
            throw new Error(`${name}: worksheet exceeds 250,000 rows or 201 columns.`);
        const text = (r, c, identifier = false) => {
            const cell = sheet[XLSX.utils.encode_cell({ r, c })];
            if (!cell || cell.v == null)
                return '';
            if (cell.t === 'e')
                return '#ERROR';
            if (identifier && typeof cell.v === 'number' && !Number.isSafeInteger(cell.v)) {
                throw new Error(`${name}, row ${r + 1}: account or sequence is not a safe integer. Store identifiers as text in Excel.`);
            }
            return String(cell.v).trim();
        };
        for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 29); r++) {
            const headers = new Map();
            for (let c = range.s.c; c <= range.e.c; c++)
                headers.set(normalize(text(r, c)), c);
            if (!Object.values(required).every(h => headers.has(h)))
                continue;
            const sourceHeaders = Array.from({ length: range.e.c - range.s.c + 1 }, (_, i) => text(r, range.s.c + i) || `Column ${XLSX.utils.encode_col(range.s.c + i)}`);
            const records = [];
            const field = (row, header, identifier = false) => {
                const col = headers.get(header);
                return col === undefined ? '' : text(row, col, identifier);
            };
            for (let row = r + 1; row <= range.e.r; row++) {
                const account = field(row, required.account, true);
                const sequence = field(row, required.sequence, true);
                if (!account && !sequence)
                    continue; // Ignore prefilled helpers and serial numbers.
                if (!/^\d+$/.test(account) || !/^\d+$/.test(sequence)) {
                    throw new Error(`${name}, row ${row + 1}: account and sequence must contain digits. Preserve leading zeros as text in Excel.`);
                }
                records.push({ account, sequence: sequence.padStart(6, '0'), amount: field(row, 'amount'),
                    date: field(row, kind === 'dn' ? 'trandatetime' : 'date'), code: field(row, 'actioncode'), sheet: name, row: row + 1,
                    source: { headers: sourceHeaders,
                        values: sourceHeaders.map((_, i) => text(row, range.s.c + i)) } });
            }
            candidates.push(records);
            break;
        }
    }
    if (candidates.length !== 1)
        throw new Error(`Expected one ${kind === 'dn' ? 'D Navi' : 'Core'} sheet with the required headers; found ${candidates.length}.`);
    if (!candidates[0].length)
        throw new Error(`No ${kind === 'dn' ? 'D Navi' : 'Core'} transactions found.`);
    return candidates[0];
}
export function reconcile(dn, core, direction = {}) {
    const configured = (direction.header ?? '').trim();
    const source = direction.file === 'core' ? core : dn;
    let directionIndex = -1;
    if (configured) {
        if (!['core', 'dn'].includes(direction.file)) throw new Error('Select the file containing the direction column, or leave the column header blank.');
        directionIndex = (source[0]?.source?.headers ?? []).findIndex(h => normalize(h) === normalize(configured));
        if (directionIndex < 0) throw new Error(`Direction column “${configured}” was not found in ${direction.file === 'core' ? 'Core' : 'D Navi'}. Check the header or leave it blank.`);
    }
    const directionValue = (c, ds) => {
        if (directionIndex < 0) return '';
        const records = direction.file === 'core' ? (c ? [c] : []) : ds;
        return [...new Set(records.map(r => r.source?.values[directionIndex] ?? '').filter(Boolean))].join(' | ');
    };
    const requiresReversal = r => /^(910|911)(?:\D|$)/.test((r.code ?? '').trim());
    const responseReason = ds => {
        const codes = [...new Set(ds.filter(requiresReversal).map(r => r.code.trim().slice(0, 3)))];
        return `D Navi reversal response (${codes.join(', ')})`;
    };
    const grouped = new Map();
    for (const d of dn) {
        if (!grouped.has(key(d))) grouped.set(key(d), []);
        grouped.get(key(d)).push(d);
    }
    const rows = [];
    const add = (c, ds, reason) => {
        const d = ds[0];
        rows.push(['Should be reversed', c?.account ?? d.account, c?.sequence ?? d.sequence,
            c?.amount ?? '', c?.date ?? '', c?.sheet ?? '', c ? String(c.row) : '',
            reason, [...new Set(ds.map(r => r.code ?? ''))].join(' | '), directionValue(c, ds),
            d?.sheet ?? '', ds.map(r => String(r.row)).join(' | ')]);
    };
    for (const c of core) {
        const matches = grouped.get(key(c)) ?? [];
        if (!matches.length) add(c, [], 'Present in Core; absent from D Navi');
        else if (matches.some(requiresReversal)) add(c, matches, responseReason(matches));
    }
    const coreKeys = new Set(core.map(key));
    for (const d of dn) if (!coreKeys.has(key(d)) && requiresReversal(d)) add(null, [d], `${responseReason([d])}; absent from Core`);
    return { columns, rows, core_count: core.length, dn_count: dn.length, reversal_count: rows.length };
}
export function exportReport(report, core = [], dn = []) {
    const book = XLSX.utils.book_new();
    const append = (name, headers, rows) => {
        // Explicit string values preserve identifiers and prevent formula execution.
        const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        sheet['!cols'] = headers.map(h => ({ wch: Math.min(40, Math.max(18, h.length + 2)) }));
        sheet['!autofilter'] = { ref: sheet['!ref'] };
        XLSX.utils.book_append_sheet(book, sheet, name);
    };
    const coreHeaders = core[0]?.source?.headers ?? [];
    const dnHeaders = dn[0]?.source?.headers ?? [];
    const byRow = new Map(core.map(r => [`${r.sheet}:${r.row}`, r]));
    const dnByRow = new Map(dn.map(r => [`${r.sheet}:${r.row}`, r]));
    append('Should be reversed', [
        ...report.columns, ...coreHeaders.map(h => `Core · ${h}`),
        'D Navi match', ...dnHeaders.map(h => `D Navi · ${h}`),
    ], report.rows.map(row => {
        const matches = row[11] ? row[11].split(' | ').map(n => dnByRow.get(`${row[10]}:${n}`)).filter(Boolean) : [];
        return [
            ...row, ...(byRow.get(`${row[5]}:${row[6]}`)?.source?.values ?? coreHeaders.map(() => '')),
            matches.length ? 'Present' : 'Not present',
            ...dnHeaders.map((_, i) => matches.map(r => r.source?.values[i] ?? '').join(' | ')),
        ];
    }));
    const sourceSheet = (name, records, headers) => {
        append(name, ['Source sheet', 'Source row', ...headers], records.map(r => [
            r.sheet, String(r.row), ...(r.source?.values ?? headers.map(() => '')),
        ]));
    };
    sourceSheet('Core source', core, coreHeaders);
    sourceSheet('D Navi source', dn, dnHeaders);
    return XLSX.write(book, { type: 'array', bookType: 'xlsx', compression: true });
}
