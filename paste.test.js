// 貼表自我檢查：用範本裡 114供參 的實際填法做兩份回傳檔，確認順序、合計／主管／基金列與說明列位置
const assert = require('assert'), fs = require('fs'), vm = require('vm');
global.ExcelJS = require('exceljs');
global.XLSX = require('xlsx');
Object.assign(global, require('./form-builder.js'));
global.document = { getElementById: () => ({ classList: { toggle() {} } }), addEventListener() {} };
vm.runInThisContext(fs.readFileSync('./paste.js', 'utf8') + ';global.ps_copyRow=ps_copyRow;global.ps_t={ps_addFiles,ps_build,set:(b,o)=>{ps_base=b;ps_orgs=o}}');

const name = '(空白)02-非營業-115協商彙總表-委員會審查結果';
const buf = fs.readFileSync(`forms/${name}.xlsx`);
// 回傳檔：空白範本的「115請查填」合計列以下換成 114供參 的 rows 列
async function filled(rows) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.getWorksheet('115請查填'), old = wb.getWorksheet('114供參');
    (ws.model.merges || []).forEach(m => { if (+m.split(':')[0].replace(/\D/g, '') > 7) ws.unMergeCells(m); });
    ws.spliceRows(8, ws.rowCount);
    // 照實際檔案保留標題列的合併（A:D、F:H）
    rows.forEach((r, i) => ps_copyRow(old, r, ws, 8 + i, 21));
    wb.worksheets.filter(w => w !== ws).forEach(w => wb.removeWorksheet(w.id)); // 往年工作表有共用公式，ExcelJS 存不回去
    return wb.xlsx.writeBuffer();
}
(async () => {
    ps_t.set({ name, buf }, JSON.parse(fs.readFileSync('surveys/orgs.json')));
    const file = async (n, rows) => { const b = await filled(rows); return { name: n, arrayBuffer: async () => b }; };
    // 故意先拖通傳會，輸出應依標準名單把法務部排前面
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13, 14, 15]), await file('法務部矯正機關作業基金.xlsx', [8, 9, 10])]);
    const out = new ExcelJS.Workbook();
    await out.xlsx.load(await ps_t.ps_build());
    assert.deepStrictEqual(out.worksheets.map(w => w.name), ['115請查填']);
    const ws = out.worksheets[0], a = r => String(ws.getRow(r).getCell(1).text).trim();
    assert.deepStrictEqual([7, 8, 9, 11, 12].map(a), ['合計', '法務部主管', '法務部矯正機關作業基金', '通傳會主管', '通訊傳播監督管理基金']);
    assert.strictEqual(ws.getCell('F7').value, '共4案');
    assert.strictEqual(ws.getCell('F12').value, '共3案');
    assert.strictEqual(ws.getCell('P7').value.formula, 'COUNTIF(P8:P15,"v")');
    assert.strictEqual(ws.getCell('P11').value.formula, 'COUNTIF(P12:P15,"v")');
    assert.strictEqual(ws.getCell('P11').value.result, 2); // 委保1、委保2 打 V
    assert.ok(a(16).startsWith('註'), '說明列接在最後一列資料後面');
    assert.ok(ws.model.merges.includes('A8:D8'), '標題列的合併儲存格照搬');

    // 沒有存好的範本：版面取自第一個拖入的檔案（通傳會），結果一樣依標準順序、合計重算
    ps_t.set(null, JSON.parse(fs.readFileSync('surveys/orgs.json')));
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13, 14, 15]), await file('法務部矯正機關作業基金.xlsx', [8, 9, 10])]);
    const out2 = new ExcelJS.Workbook();
    await out2.xlsx.load(await ps_t.ps_build());
    const w2 = out2.worksheets[0], a2 = r => String(w2.getRow(r).getCell(1).text).trim();
    assert.deepStrictEqual([7, 8, 11, 12].map(a2), ['合計', '法務部主管', '通傳會主管', '通訊傳播監督管理基金'], '回傳檔自己的主管、基金列不能被當成佔位列略過');
    assert.strictEqual(w2.getCell('F7').value, '共4案');
    console.log('paste: all checks passed');
})().catch(e => { console.error(e); process.exit(1); });
