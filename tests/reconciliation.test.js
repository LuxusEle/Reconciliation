import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from '../vendor/xlsx.mjs';
import { extract, reconcile, exportReport } from '../reconciliation.js';
const tx = (account, sequence, row = 2) => ({ account, sequence, amount: '1000', date: '60325', sheet: 'Core', row });
function excel(headers, rows) {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([headers, ...rows]), 'Data');
    return XLSX.write(book, { type: 'array', bookType: 'xlsx' });
}
test('only missing Core rows, including duplicates, are marked for reversal', () => {
    const result = reconcile([tx('001', '000001')], [tx('001', '000001'), tx('002', '000002'), tx('002', '000002', 3)]);
    assert.equal(result.reversal_count, 2);
    assert.ok(result.rows.every(r => r[0] === 'Should be reversed' && r[1] === '002'));
});
test('911 is reversed even with masked amount', () => {
    const dn = extract(excel(['PAN', 'System Trace Number', 'Action Code', 'Ntwk Recon Amt'], [['001', '1', '911', 'XXXXX']]), 'dn');
    const core = extract(excel(['Account No', 'SeqNo', 'Amount', 'Date'], [['001', '000001', '1000', '60325']]), 'core');
    assert.equal(reconcile(dn, core).reversal_count, 1);
    assert.equal(dn[0].account, '001');
    assert.equal(dn[0].sequence, '000001');
});
test('sample includes missing Core rows and all 911 timeouts', () => {
    const bytes = readFileSync('resources/Reconcilation.xlsx');
    const data = Uint8Array.from(bytes).buffer;
    const dn = extract(data, 'dn'), core = extract(data, 'core');
    const report = reconcile(dn, core);
    assert.equal(core.length, 10725);
    assert.equal(dn.length, 2843);
    assert.equal(report.rows.filter(r => r[7] === 'Present in Core; absent from D Navi').length, 8787);
    assert.ok(report.reversal_count > 8787);
    const book = XLSX.read(exportReport(report, core, dn), { type: 'array' });
    assert.deepEqual(book.SheetNames, ['Should be reversed', 'Core source', 'D Navi source']);
    const rows = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header: 1 });
    assert.equal(rows.length, report.reversal_count + 1);
    assert.deepEqual(rows[1].slice(0, report.columns.length), report.rows[0]);
    assert.ok(rows[0].includes('Core · Trn Success "Y"'));
    assert.ok(rows[0].includes('D Navi · Action Code'));
    assert.equal(rows[1][report.columns.length + core[0].source.headers.length], 'Not present');
    const dnRows = XLSX.utils.sheet_to_json(book.Sheets['D Navi source'], { header: 1, defval: '' });
    assert.equal(dnRows.length, 2844);
    assert.ok(dnRows[0].includes('Ntwk Recon Amt'));
    assert.deepEqual(dnRows[1].slice(2), dn[0].source.values);
    const coreRows = XLSX.utils.sheet_to_json(book.Sheets['Core source'], { header: 1, defval: '' });
    assert.equal(coreRows.length, 10726);
    assert.deepEqual(coreRows[1].slice(2), core[0].source.values);
});
test('identifiers stay text and formula-like values are not formulas', () => {
    const report = reconcile([], [{ ...tx('001', '000001'), amount: '=1+1' }]);
    const book = XLSX.read(exportReport(report), { type: 'array' });
    const sheet = book.Sheets[book.SheetNames[0]];
    assert.equal(sheet.B2.v, '001');
    assert.equal(sheet.D2.v, '=1+1');
    assert.equal(sheet.D2.f, undefined);
});
test('invalid headers and identifiers give actionable errors', () => {
    assert.throws(() => extract(excel(['wrong'], [['x']]), 'dn'), /required headers/);
    assert.throws(() => extract(excel(['PAN', 'System Trace Number', 'Action Code'], [['001', 'abc', '000']]), 'dn'), /row 2/);
});

test('910 and 911 without Core are included, approvals are excluded', () => {
    const dn = [{...tx('001','000001'), code:'911-Timeout'}, {...tx('002','000002'), code:'000-Approved'}, {...tx('003','000003'), code:'910-Denied'}];
    const report = reconcile(dn, []);
    assert.equal(report.rows.length,2);
    assert.equal(report.rows[0][1],'001');
    assert.equal(report.rows[1][1],'003');
    assert.equal(report.rows[0][6],'');
    assert.match(report.rows[0][7],/absent from Core/);
});
test('direction can come from either file and invalid headers are rejected', () => {
    const core = extract(excel(['Account No','SeqNo','Amount','Date','Direction'],[['001','1','100','d','Outgoing']]),'core');
    const dn = extract(excel(['PAN','System Trace Number','Action Code','Role'],[['001','1','911-Timeout','Acquirer']]),'dn');
    assert.equal(reconcile(dn,core).rows[0][9],'');
    assert.equal(reconcile(dn,core,{file:'core',header:' direction '}).rows[0][9],'Outgoing');
    const report = reconcile(dn,core,{file:'dn',header:'Role'});
    assert.equal(report.rows[0][9],'Acquirer');
    assert.throws(()=>reconcile(dn,core,{file:'core',header:'Missing'}),/not found/);
    const book = XLSX.read(exportReport(report,core,dn),{type:'array'});
    const rows = XLSX.utils.sheet_to_json(book.Sheets['Should be reversed'],{header:1,defval:''});
    const action = rows[0].indexOf('D Navi · Action Code');
    assert.equal(rows[1][action],'911-Timeout');
    assert.equal(rows[1][report.columns.length + core[0].source.headers.length],'Present');
});
test('duplicate D Navi matches retain timeout evidence even alongside approval', () => {
    const dn = [{...tx('001','000001'),code:'000-Approved'},{...tx('001','000001',3),code:'911-Timeout'}];
    const report = reconcile(dn,[tx('001','000001')]);
    assert.equal(report.reversal_count,1);
    assert.equal(report.rows[0][11],'2 | 3');
    assert.match(report.rows[0][8],/911/);
});

test('910 with Core is reversed and 9100 is not treated as 910', () => {
    const dn = [{...tx('001','000001'),code:'910-Card Issuer Signed Off'}, {...tx('002','000002'),code:'9100'}];
    const report = reconcile(dn,[tx('001','000001'),tx('002','000002')]);
    assert.equal(report.reversal_count,1);
    assert.equal(report.rows[0][1],'001');
    assert.match(report.rows[0][7],/910/);
});
