// 貼表自我檢查：用範本裡 114供參 的實際填法做兩份回傳檔，確認順序、合計／主管／基金列與說明列位置
const assert = require('assert'), fs = require('fs'), vm = require('vm');
global.ExcelJS = require('exceljs');
global.XLSX = require('xlsx');
Object.assign(global, require('./form-builder.js'));
global.document = { getElementById: () => ({ classList: { toggle() {} } }), addEventListener() {} };
vm.runInThisContext(fs.readFileSync('./paste.js', 'utf8') + ';global.ps_copyRow=ps_copyRow;global.ps_t={ps_addFiles,ps_build,ps_effMode,ps_fillPlan,files:()=>ps_files,set:async(b,o)=>{ps_orgs=o;ps_files=[];ps_stored=b;await ps_rebase()}}');

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
    const orgs = JSON.parse(fs.readFileSync('surveys/orgs.json'));
    await ps_t.set({ name, buf }, orgs);
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
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13, 14, 15]), await file('法務部矯正機關作業基金.xlsx', [8, 9, 10])]);
    const out2 = new ExcelJS.Workbook();
    await out2.xlsx.load(await ps_t.ps_build());
    const w2 = out2.worksheets[0], a2 = r => String(w2.getRow(r).getCell(1).text).trim();
    assert.deepStrictEqual([7, 8, 11, 12].map(a2), ['合計', '法務部主管', '通傳會主管', '通訊傳播監督管理基金'], '回傳檔自己的主管、基金列不能被當成佔位列略過');
    assert.strictEqual(w2.getCell('F7').value, '共4案');

    // 手動調整順序：通傳會移到前面，之後再拖入的檔案不會把順序排回去
    const fl = ps_t.files();
    [fl[0], fl[1]] = [fl[1], fl[0]];
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13])]); // 同機關重拖：原位置換掉
    assert.deepStrictEqual(ps_t.files().map(f => f.org), ['通訊傳播監督管理基金', '法務部矯正機關作業基金']);
    const out3 = new ExcelJS.Workbook();
    await out3.xlsx.load(await ps_t.ps_build());
    assert.strictEqual(String(out3.worksheets[0].getCell('A8').text).trim(), '通傳會主管');

    // 空白範本跟回傳檔一起拖入：檔名有「空白」自動勾成範本，不貼入
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13, 14, 15]), { name: name + '.xlsx', arrayBuffer: async () => buf }, await file('法務部矯正機關作業基金.xlsx', [8, 9, 10])]);
    assert.deepStrictEqual(ps_t.files().map(f => f.tpl), [true, false, false]);
    const out4 = new ExcelJS.Workbook();
    await out4.xlsx.load(await ps_t.ps_build());
    const w4 = out4.worksheets[0], a4 = r => String(w4.getRow(r).getCell(1).text).trim();
    assert.strictEqual(w4.getCell('F7').value, '共4案');
    assert.deepStrictEqual([8, 11].map(a4), ['法務部主管', '通傳會主管']);
    assert.ok(a4(16).startsWith('註'), '範本的說明列接在最後');

    // 依項目填入：每份都有全部基金、各填自己那一列
    const F = orgs.sortOrder.slice(0, 6);
    const book = async (cols, rows) => {
        const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('表');
        ws.addRow(['測試表']); ws.addRow(cols);
        F.forEach((n, i) => { const row = ws.addRow([n, ...(rows[i] || [])]); const last = cols.length; row.getCell(last).value = { formula: `${ws.getColumn(last - 1).letter}${row.number}/B${row.number}` }; });
        const b = await wb.xlsx.writeBuffer(); return b;
    };
    const head = ['基金名稱', '預算數', '決算數', '執行率'];
    const asFile = (name, b) => ({ name, arrayBuffer: async () => b });
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([
        asFile('空白範本.xlsx', await book(head, [])),
        asFile(F[0] + '.xlsx', await book(head, [[100, 90]])),
        // 自己插了「調整數」欄，決算數 777 往下拖到別的基金
        asFile(F[1] + '.xlsx', await book(['基金名稱', '預算數', '調整數', '決算數', '執行率'], [[], [200, 5, 777], [null, null, 777], [null, null, 777]])),
        // 順手填了別的基金（F[4]）
        asFile(F[2] + '.xlsx', await book(head, [[], [], [300, 280], [], [50]])),
    ]);
    assert.strictEqual(ps_t.ps_effMode(), 'fill');
    const plan = ps_t.ps_fillPlan(), cat = k => plan.checks.filter(c => c[0] === k).map(c => c[1]);
    assert.deepStrictEqual(cat('沒有對應到的基金'), [F[3], F[5]].map(n => n.replace(/\s/g, '')));
    assert.deepStrictEqual(cat('由別的機關的檔案填入'), [F[4]].map(n => n.replace(/\s/g, '')));
    assert.deepStrictEqual(cat('範本沒有的欄'), ['調整數']);
    assert.strictEqual(cat('重複數字（疑似往下拖曳）').length, 1);
    const o5 = new ExcelJS.Workbook();
    await o5.xlsx.load(await ps_t.ps_build());
    const w5 = o5.worksheets[0], v = (r, c) => w5.getRow(r).getCell(c).value;
    assert.deepStrictEqual([v(3, 2), v(3, 3), v(4, 2), v(4, 3), v(5, 3), v(6, 3)], [100, 90, 200, 777, 280, null], '欄依表頭對位、拖曳只留最上面一格');
    assert.strictEqual(v(7, 2), 50);
    assert.strictEqual(w5.getRow(7).getCell(2).fill?.fgColor?.argb, 'FFF8CBAD', '別的機關填的標橘');
    assert.notStrictEqual(w5.getRow(3).getCell(2).fill?.fgColor?.argb, 'FFF8CBAD', '其他格不能跟著變色');
    assert.deepStrictEqual(o5.worksheets.map(w => w.name), ['表', '查對']);

    // 有人在最左邊插一欄（序號）：兩種貼法都依表頭文字找到名稱欄
    const lead = async (cols, rows) => {
        const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('表');
        ws.addRow(['測試表']); ws.addRow(cols); rows.forEach(r => ws.addRow(r));
        return wb.xlsx.writeBuffer();
    };
    const all = F.map(n => [n]);
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([
        asFile('空白範本.xlsx', await lead(['基金名稱', '預算數', '決算數'], all)),
        asFile(F[0] + '.xlsx', await lead(['基金名稱', '預算數', '決算數'], all.map((r, i) => i === 0 ? [r[0], 100, 90] : r))),
        asFile(F[1] + '.xlsx', await lead(['序號', '基金名稱', '預算數', '決算數'], all.map((r, i) => [i + 1, r[0], ...(i === 1 ? [200, 150] : [])]))),
        asFile(F[2] + '.xlsx', await lead(['序號', '名稱', '預算數', '決算數'], all.map((r, i) => [i + 1, r[0], ...(i === 2 ? [300, 280] : [])]))),
    ]);
    const p6 = ps_t.ps_fillPlan(), c6 = k => p6.checks.filter(c => c[0] === k);
    assert.deepStrictEqual(p6.fills.map(x => [x.r, x.data]), [[3, { 2: 100, 3: 90 }], [4, { 2: 200, 3: 150 }]], '最左邊插欄的那份照樣填進自己那列');
    assert.deepStrictEqual(c6('找不到名稱欄').map(c => c[2]), [F[2] + '.xlsx'], '名稱欄改了表頭的那份列在查對');
    assert.deepStrictEqual(c6('範本沒有的欄').map(c => c[1]), ['序號', '序號', '名稱']);
    const sh = ['提案序號', '提案委員', '內容'];
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([
        asFile(F[0] + '.xlsx', await lead(sh, [['委保1', '王', '甲案']])),
        asFile(F[1] + '.xlsx', await lead(['序號', ...sh], [[1, '委保1', '李', '乙案']])),
    ]);
    const o7 = new ExcelJS.Workbook();
    await o7.xlsx.load(await ps_t.ps_build());
    const r7 = r => [1, 2, 3].map(c => o7.worksheets[0].getRow(r).getCell(c).text);
    assert.deepStrictEqual([r7(3), r7(4)], [['委保1', '王', '甲案'], ['委保1', '李', '乙案']], '接續貼上也依表頭對位，不會整列右移');
    console.log('paste: all checks passed');
})().catch(e => { console.error(e); process.exit(1); });
