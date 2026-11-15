const element = (id) => document.getElementById(id);
const form = element('form');
const run = element('run');
const message = element('message');
const results = element('results');
const search = element('search');
let report = null;
let url = null;
let activeFilter = 'all';
const filterLabels = { all: 'All', '911': '911', '910': '910', present: 'Present in D Navi', absent: 'Absent from D Navi' };
function matchesFilter(row, filter) {
    if (filter === 'present') return Boolean(row[11]);
    if (filter === 'absent') return !row[11];
    if (filter === '911' || filter === '910') {
        return row[8].split(' | ').some(code => new RegExp(`^${filter}(?:\\D|$)`).test(code.trim()));
    }
    return true;
}
function cell(tag, value) {
    const node = document.createElement(tag);
    node.textContent = String(value);
    return node;
}
function render() {
    if (!report)
        return;
    const query = search.value.trim();
    const searched = report.rows.filter(r => !query || r[1].includes(query) || r[2].includes(query));
    const rows = searched.filter(r => matchesFilter(r, activeFilter));
    for (const button of element('table-filters').querySelectorAll('button')) {
        const filter = button.dataset.filter;
        button.setAttribute('aria-pressed', String(filter === activeFilter));
        button.textContent = `${filterLabels[filter]} (${searched.filter(r => matchesFilter(r, filter)).length.toLocaleString()})`;
    }
    const body = element('body');
    body.replaceChildren();
    for (const row of rows.slice(0, 500)) {
        const tr = document.createElement('tr');
        row.forEach(v => tr.append(cell('td', v)));
        body.append(tr);
    }
    element('shown').textContent = rows.length
        ? `${rows.length.toLocaleString()} transactions · Showing up to 500 rows. Excel contains all records.`
        : 'No transactions found. There are no reversal transactions matching this view.';
}
async function readFile(id) {
    const file = element(id).files?.[0];
    if (!file || !file.name.toLowerCase().endsWith('.xlsx'))
        throw new Error('Select an .xlsx file for each input.');
    if (file.size > 25 * 1024 * 1024)
        throw new Error('Each workbook must be smaller than 25 MB.');
    return file.arrayBuffer();
}
function processFiles(dn, core, direction) {
    return new Promise((resolve, reject) => {
        const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
        worker.onmessage = event => {
            worker.terminate();
            if (event.data.error)
                reject(new Error(event.data.error));
            else
                resolve(event.data);
        };
        worker.onerror = () => { worker.terminate(); reject(new Error('Unable to process workbooks. Try smaller files.')); };
        worker.postMessage({ dn, core, direction }, [dn, core]);
    });
}
form.onsubmit = async (event) => {
    event.preventDefault();
    run.disabled = true;
    results.hidden = true;
    report = null;
    if (url) {
        URL.revokeObjectURL(url);
        url = null;
    }
    message.className = '';
    message.textContent = 'Reading workbooks and matching transactions in your browser…';
    try {
        const [dn, core] = await Promise.all([readFile('dn'), readFile('core')]);
        const data = await processFiles(dn, core, { file: element('direction-file').value, header: element('direction-header').value });
        report = data.report;
        url = URL.createObjectURL(new Blob([data.file], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        const cards = element('cards');
        cards.replaceChildren();
        for (const [label, count] of [['Core transactions', report.core_count], ['D Navi transactions', report.dn_count], ['Should be reversed', report.reversal_count]]) {
            const card = document.createElement('div');
            card.className = 'card';
            card.append(cell('b', count.toLocaleString()), cell('span', label));
            cards.append(card);
        }
        const tr = document.createElement('tr');
        report.columns.forEach(v => tr.append(cell('th', v)));
        element('head').replaceChildren(tr);
        search.value = '';
        activeFilter = 'all';
        results.hidden = false;
        message.textContent = report.reversal_count ? 'List ready. Includes Core-only transactions and D Navi 910/911 responses.' : 'No reversals found under the Core-only and 910/911 response rules.';
        render();
    }
    catch (error) {
        message.className = 'error';
        message.textContent = error instanceof Error ? error.message : 'Unable to process workbooks.';
    }
    finally {
        run.disabled = false;
    }
};
search.oninput = render;
element('table-filters').onclick = event => {
    const button = event.target.closest('button[data-filter]');
    if (!button) return;
    activeFilter = button.dataset.filter;
    render();
};
element('download').onclick = () => {
    if (!url)
        return;
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Transactions-to-reverse.xlsx';
    link.click();
};
export {};

function updateDirectionOptions() {
    element('direction-options').hidden = !(element('dn').files?.length && element('core').files?.length);
    results.hidden = true;
    report = null;
    if (url) { URL.revokeObjectURL(url); url = null; }
    message.textContent = '';
}
element('dn').onchange = updateDirectionOptions;
element('core').onchange = updateDirectionOptions;
for (const id of ['direction-file', 'direction-header']) element(id).oninput = () => {
    results.hidden = true;
    message.textContent = 'Settings changed. Generate the report again.';
};
