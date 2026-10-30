// 貼表自我檢查：用範本裡 114供參 的實際填法做兩份回傳檔，確認順序、主管／基金列與說明列位置
const assert = require('assert'), fs = require('fs'), vm = require('vm');
global.ExcelJS = require('exceljs');
global.XLSX = require('xlsx');
Object.assign(global, require('./form-builder.js'));
global.document = { getElementById: () => ({ classList: { toggle() {} } }), addEventListener() {} };
vm.runInThisContext(fs.readFileSync('./paste.js', 'utf8') + ';global.ps_copyRow=ps_copyRow;global.ps_t={ps_addFiles,applyDel:f=>ps_applyDel(f),ps_resort,ps_setOrder,ps_build,mode:m=>{ps_mode=m},ps_fillPlan,stackChecks:()=>ps_stackChecks(),files:()=>ps_files,set:async(b,o)=>{ps_orgs=o;ps_files=[];ps_mode="stack";await ps_rebase();if(b)await ps_addFiles([{name:b.name+".xlsx",arrayBuffer:async()=>b.buf}])}}');

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
    // 合計、小計不由程式產生：範本的合計列不帶，資料緊接在表頭下面
    assert.deepStrictEqual([7, 8, 10, 11].map(a), ['法務部主管', '法務部矯正機關作業基金', '通傳會主管', '通訊傳播監督管理基金']);
    const formulas = []; ws.eachRow(row => row.eachCell(c => { if (ps_isFormula(c.value)) formulas.push(c.address); }));
    assert.deepStrictEqual(formulas, [], '不產生 COUNTIF 等小計公式');
    assert.ok(!ws.getColumn(1).values.some(v => String(v).trim() === '合計'), '沒有合計列');
    assert.ok(a(15).startsWith('註'), '說明列接在最後一列資料後面');
    assert.ok(ws.model.merges.includes('A7:D7'), '標題列的合併儲存格照搬');

    // 沒有存好的範本：版面取自第一個拖入的檔案（通傳會），結果一樣依標準順序
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13, 14, 15]), await file('法務部矯正機關作業基金.xlsx', [8, 9, 10])]);
    const out2 = new ExcelJS.Workbook();
    await out2.xlsx.load(await ps_t.ps_build());
    const w2 = out2.worksheets[0], a2 = r => String(w2.getRow(r).getCell(1).text).trim();
    assert.deepStrictEqual([7, 10, 11].map(a2), ['法務部主管', '通傳會主管', '通訊傳播監督管理基金'], '回傳檔自己的主管、基金列不能被當成佔位列略過');

    // 手動調整順序：通傳會移到前面，之後再拖入的檔案不會把順序排回去
    const fl = ps_t.files();
    [fl[0], fl[1]] = [fl[1], fl[0]];
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13])]); // 同機關重拖：原位置換掉
    assert.deepStrictEqual(ps_t.files().map(f => f.org), ['通訊傳播監督管理基金', '法務部矯正機關作業基金']);
    const out3 = new ExcelJS.Workbook();
    await out3.xlsx.load(await ps_t.ps_build());
    assert.strictEqual(String(out3.worksheets[0].getCell('A7').text).trim(), '通傳會主管');

    // 空白範本跟回傳檔一起拖入：檔名有「空白」自動勾成範本，不貼入
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([await file('通訊傳播監督管理基金.xlsx', [11, 12, 13, 14, 15]), { name: name + '.xlsx', arrayBuffer: async () => buf }, await file('法務部矯正機關作業基金.xlsx', [8, 9, 10])]);
    assert.deepStrictEqual(ps_t.files().map(f => f.tpl), [true, false, false]);
    const out4 = new ExcelJS.Workbook();
    await out4.xlsx.load(await ps_t.ps_build());
    const w4 = out4.worksheets[0], a4 = r => String(w4.getRow(r).getCell(1).text).trim();
    assert.deepStrictEqual([7, 10].map(a4), ['法務部主管', '通傳會主管']);
    assert.ok(a4(15).startsWith('註'), '範本的說明列接在最後');

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
    const plan = ps_t.ps_fillPlan(), cat = k => plan.checks.filter(c => c[0] === k).map(c => c[1]);
    assert.deepStrictEqual(cat('沒有對應到的基金'), [F[3], F[5]].map(n => n.replace(/\s/g, '')));
    assert.deepStrictEqual(cat('由別的機關的檔案填入'), [F[4]].map(n => n.replace(/\s/g, '')));
    assert.deepStrictEqual(cat('範本沒有的欄'), ['調整數']);
    assert.strictEqual(cat('重複數字（疑似往下拖曳）').length, 1);
    ps_t.mode('fill'); // 每份都有全部基金的表：手動切到依項目填入
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
    // 一份含多個機關：依最前面那個機關排，不是排最後
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([
        asFile(F[3] + '.xlsx', await lead(sh, [['委保1', '丁', '丁案']])),
        asFile(`${F[2]}、${F[4]}.xlsx`, await lead(sh, [['委保1', '丙', '丙案'], ['委保2', '戊', '戊案']])),
        asFile(F[0] + '.xlsx', await lead(sh, [['委保1', '甲', '甲案']])),
    ]);
    assert.deepStrictEqual(ps_t.files().map(f => f.org), [F[0], `${F[2]}、${F[4]}`, F[3]]);
    const fs3 = ps_t.files(); [fs3[0], fs3[2]] = [fs3[2], fs3[0]];
    ps_t.ps_resort();
    assert.deepStrictEqual(ps_t.files().map(f => f.org), [F[0], `${F[2]}、${F[4]}`, F[3]], '依基金清單重排回到標準順序');
    // 自訂順序清單：照清單排，清單外的排最後；查對列出清單上沒檔案的、檔案不在清單的
    ps_t.ps_setOrder(`${F[3]}\n${F[0]}\n${F[5]}`);
    assert.deepStrictEqual(ps_t.files().map(f => f.org), [F[3], F[0], `${F[2]}、${F[4]}`]);
    const oc = ps_t.stackChecks();
    assert.deepStrictEqual(oc.filter(c => c[0] === '清單上沒有對應的檔案').map(c => c[1]), [F[5]]);
    assert.deepStrictEqual(oc.filter(c => c[0] === '不在自訂清單').map(c => c[1]), [`${F[2]}、${F[4]}`]);
    // 認得出標準機關的行只比機關：F[2] 那行不會因為檔名含 F[2] 以外的字而誤對
    assert.ok(!ps_t.files().some(f => f.org === F[3] && f.rank === 1));
    ps_t.ps_setOrder('');
    assert.deepStrictEqual(ps_t.files().map(f => f.org), [F[0], `${F[2]}、${F[4]}`, F[3]], '清空清單回到標準順序');
    assert.ok(ps_t.stackChecks().some(c => c[0] === '檔名含多個機關' && c[3].includes(F[2])));
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([
        asFile(F[0] + '.xlsx', await lead(sh, [['委保1', '王', '甲案']])),
        asFile(F[1] + '.xlsx', await lead(['序號', ...sh], [[1, '委保1', '李', '乙案']])),
    ]);
    const o7 = new ExcelJS.Workbook();
    await o7.xlsx.load(await ps_t.ps_build());
    const r7 = r => [1, 2, 3].map(c => o7.worksheets[0].getRow(r).getCell(c).text);
    assert.deepStrictEqual([r7(3), r7(4)], [['委保1', '王', '甲案'], ['委保1', '李', '乙案']], '接續貼上也依表頭對位，不會整列右移');

    // 隱藏欄、列：照樣讀入，回傳檔自己藏的（範本沒藏）列在查對
    const hid = async (cols, rows, hideCols = [], hideRows = []) => {
        const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('表');
        ws.addRow(['測試表']); ws.addRow(cols); rows.forEach(r => ws.addRow(r));
        hideCols.forEach(c => { ws.getColumn(c).hidden = true; }); hideRows.forEach(r => { ws.getRow(r).hidden = true; });
        return wb.xlsx.writeBuffer();
    };
    const hh = ['基金名稱', '預算數', '決算數'];
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([
        asFile('空白範本.xlsx', await hid(hh, all, [2])),                                                  // 範本自己藏 B 欄：不提醒
        asFile(F[0] + '.xlsx', await hid(hh, all.map((r, i) => i === 0 ? [r[0], 100, 90] : r), [2, 3])),   // 回傳檔多藏 C 欄：提醒
        asFile(F[1] + '.xlsx', await hid(hh, all.map((r, i) => i === 1 ? [r[0], 200, 150] : r), [], [4])), // 藏了自己那列：提醒
    ]);
    const p8 = ps_t.ps_fillPlan();
    assert.deepStrictEqual(p8.fills.map(x => x.data), [{ 2: 100, 3: 90 }, { 2: 200, 3: 150 }], '隱藏的欄、列照樣讀入');
    assert.deepStrictEqual(p8.checks.filter(c => c[0].startsWith('隱藏')).map(c => [c[0], c[1]]), [['隱藏欄有資料', '決算數'], ['隱藏列有資料', F[1]]]);
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([asFile(F[0] + '.xlsx', await lead(sh, [['委保1', '王', '甲案']])), asFile(F[1] + '.xlsx', await hid(sh, [['委保1', '李', '乙案']], [3]))]);
    assert.ok(ps_t.stackChecks().some(c => c[0] === '隱藏欄有資料' && c[1] === '內容'), '接續貼上也提醒隱藏欄');

    // 一個基金填好幾列：基金底下的明細列各基金同名，連同所屬基金對位，也算那個基金的（不標橘）
    const dn = [F[0], '(一)業務計畫', F[1], '(一)業務計畫', F[2], F[3], F[4]];
    const det = async rows => { const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('表'); ws.addRow(['測試表']); ws.addRow(['名稱', '預算數', '決算數']);
        dn.forEach((n, i) => ws.addRow([n, ...(rows[i] || [])])); return wb.xlsx.writeBuffer(); };
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([asFile('空白範本.xlsx', await det([])), asFile(F[0] + '.xlsx', await det([[10, 9], [5, 4]])), asFile(F[1] + '.xlsx', await det([[], [], [100, 90], [60, 50]]))]);
    const p9 = ps_t.ps_fillPlan();
    assert.deepStrictEqual(p9.fills.map(x => [x.r, x.data[2], !!x.flag]), [[3, 10, false], [4, 5, false], [5, 100, false], [6, 60, false]]);
    assert.ok(!p9.checks.some(c => c[1] === '(一)業務計畫'));

    // 接續貼上只取自己那一段：每份都帶全部基金清單時，只貼自己那列（含底下明細）；只有別人沒有自己的，不貼
    const full = async (own, extra = {}) => { const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('表'); ws.addRow(['測試表']); ws.addRow(['名稱', '預算數', '決算數']);
        F.forEach(n => { ws.addRow([n, ...(n === own ? [10, 9] : [])]); if (n === own) ws.addRow(['(一)明細', 5, 4]); });
        Object.entries(extra).forEach(([n, v]) => { ws.getRow(3 + F.indexOf(n) + (F.indexOf(n) > F.indexOf(own) ? 1 : 0)).getCell(2).value = v; });
        return wb.xlsx.writeBuffer(); };
    await ps_t.set(null, orgs);
    await ps_t.ps_addFiles([asFile(F[2] + '.xlsx', await full(F[2], { [F[4]]: 77 })), asFile(F[0] + '.xlsx', await full(F[0])), asFile('國軍老舊眷村改建基金.xlsx', await full(null))]);
    assert.deepStrictEqual(ps_t.files().map(f => f.rows.length), [2, 2, 0], '只取自己那列與底下明細；沒有自己的列就不貼');
    const c10 = ps_t.stackChecks();
    assert.ok(c10.some(c => c[0] === '找不到自己基金的列' && c[2] === '國軍老舊眷村改建基金.xlsx'));
    assert.ok(c10.some(c => c[0] === '也填了別的基金' && c[1] === F[4]), '別的基金有填資料的列在查對');
    const o10 = new ExcelJS.Workbook();
    await o10.xlsx.load(await ps_t.ps_build());
    const r10 = r => o10.worksheets[0].getRow(r).getCell(1).text;
    assert.deepStrictEqual([3, 4, 5, 6].map(r10), [F[0], '(一)明細', F[2], '(一)明細']);
    // 逐列檢查：預設勾掉別人的列，可以改；全部檔案的每一列都是候選
    const fA = ps_t.files()[0];
    assert.strictEqual(fA.cand.length, F.length + 1);
    assert.strictEqual(fA.del.size, F.length - 1, '自己那列與明細以外都預設刪除');
    fA.del.delete(fA.cand.find(r => !fA.rows.includes(r))); ps_t.applyDel(fA);
    assert.strictEqual(fA.rows.length, 3, '取消勾選的列會貼上');
    // 自訂清單上沒檔案的，在清單位置插「❌ 未回傳」列
    fA.del = new Set(fA.del0); ps_t.applyDel(fA);
    ps_t.ps_setOrder(`${F[0]}\n${F[1]}\n${F[2]}`);
    const o11 = new ExcelJS.Workbook();
    await o11.xlsx.load(await ps_t.ps_build());
    const r11 = r => [1, 2].map(c => o11.worksheets[0].getRow(r).getCell(c).text).join('|');
    assert.deepStrictEqual([3, 4, 5, 6, 7].map(r11), [`${F[0]}|10`, '(一)明細|5', `${F[1]}|❌ 未回傳／無資料`, `${F[2]}|10`, '(一)明細|5']);
    ps_t.ps_setOrder('');

    // 沒有合計列的一般調查表：「註」說明列只在最後出現一次（有勾範本、沒勾範本都一樣）
    const plain = async rows => { const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('S'); [['調查表'], ['單位：千元'], ['項目', '金額'], ...rows, ['註：請填千元']].forEach(r => ws.addRow(r)); return wb.xlsx.writeBuffer(); };
    for (const tpl of [true, false]) {
        await ps_t.set(null, orgs);
        await ps_t.ps_addFiles([...(tpl ? [asFile('空白範本.xlsx', await plain([['ＯＯ計畫']]))] : []), asFile(F[0] + '.xlsx', await plain([['甲計畫', 100]])), asFile(F[2] + '.xlsx', await plain([['乙計畫', 200]]))]);
        const o12 = new ExcelJS.Workbook();
        await o12.xlsx.load(await ps_t.ps_build());
        const r12 = []; o12.worksheets[0].eachRow((row, i) => { if (i > 3) r12.push(row.getCell(1).text); });
        assert.deepStrictEqual(r12, ['甲計畫', '乙計畫', '註：請填千元'], `沒有合計列（${tpl ? '有' : '沒有'}範本）說明列只在最後`);
    }
    console.log('paste: all checks passed');
})().catch(e => { console.error(e); process.exit(1); });
