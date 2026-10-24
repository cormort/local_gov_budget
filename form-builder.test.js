// node form-builder.test.js — 公式轉換與求值的自我檢查
const assert = require('assert');
const { fb_translateFormula: tr, fb_evalFormula: ev, fb_parseTSV: tsv, fb_cleanNum: num } = require('./form-builder.js');

const vals = { B: 10, C: 4, D: 6, E: 2 };
const get = k => vals[k];

assert.strictEqual(tr('B2-C2', 2), 'B-C');
assert.strictEqual(tr('$B$5*2', 5), 'B*2');
assert.strictEqual(tr('SUM(B3:D3)-E3', 3), '(B+C+D)-E');
assert.strictEqual(tr('B2-C3', 2), null);          // 跨列不支援
assert.strictEqual(tr('ROUND(B2,0)', 2), null);     // 不支援的函數
assert.strictEqual(tr('Sheet2!B2', 2), null);

assert.strictEqual(ev('B-C', get), 6);
assert.strictEqual(ev('(B+C+D)-E', get), 18);
assert.strictEqual(ev('B+C*D', get), 34);
assert.strictEqual(ev('-(B-C)/E', get), -3);
assert.strictEqual(ev('B/0', get), 0);
assert.throws(() => ev('B+', get));
assert.throws(() => ev('(B', get));

// Excel 貼上
assert.deepStrictEqual(tsv('a\t1\nb\t2\n'), [['a', '1'], ['b', '2']]);
assert.deepStrictEqual(tsv('a\t1\r\nb\t2\r\n'), [['a', '1'], ['b', '2']]);
assert.deepStrictEqual(tsv('"x\ny"\t"say ""hi"""\n'), [['x\ny', 'say "hi"']]);
assert.deepStrictEqual(tsv('a\t\tc'), [['a', '', 'c']]);
assert.strictEqual(num('1,234'), '1234');
assert.strictEqual(num(' (500) '), '-500');
assert.strictEqual(num('-'), '');
assert.strictEqual(num('-12.5'), '-12.5');
assert.strictEqual(num('abc'), 'abc');

// 多列表頭：第 0 列表名（A1:F1 合併）、第 1 列空白、第 2–3 列表頭、第 4 列起資料
//   計畫名稱(A3:A4) | 預算數(B3:D3)        | 執行數(E3:E4) | 說明(F3:F4)
//                   | 原編 | 增減 | 核列 |
const grid = [
    ['○○計畫經費表', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['計畫名稱', '預算數', '', '', '執行數', '說明'],
    ['', '原編', '增減', '核列', '', ''],
    ['一般行政', '1000', '-50', '950', '900', '人事費'],
];
const at = (r, c) => (grid[r] || [])[c] || '';
const M = (r1, c1, r2, c2) => ({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
const merges = [M(0, 0, 0, 5), M(2, 0, 3, 0), M(2, 1, 2, 3), M(2, 4, 3, 4), M(2, 5, 3, 5)];
const { fb_headerBlock: block, fb_columnLevels: levels, fb_theadRows: thead } = require('./form-builder.js');
assert.deepStrictEqual(block(at, merges, 0, 4, 0, 5), { hr: 2, he: 3 });
assert.deepStrictEqual(block(at, [], 2, 4, 0, 5), { hr: 2, he: 2 });              // 沒有合併 → 單列表頭
assert.deepStrictEqual(block(at, [M(2, 1, 2, 3)], 2, 4, 0, 5), { hr: 2, he: 3 }); // 只有橫向合併也能抓到子表頭
assert.deepStrictEqual([0, 1, 3, 4].map(c => levels(at, merges, 2, 3, c)), [['計畫名稱'], ['預算數', '原編'], ['預算數', '核列'], ['執行數']]);
const cols = [['計畫名稱'], ['預算數', '原編'], ['預算數', '增減'], ['預算數', '核列'], ['執行數']].map(l => ({ label: l.join('－'), levels: l }));
const h = thead(cols).map(row => row.map(x => `${x.text}:${x.colspan}x${x.rowspan}`));
assert.deepStrictEqual(h, [['計畫名稱:1x2', '預算數:3x1', '執行數:1x2'], ['原編:1x1', '增減:1x1', '核列:1x1']]);
assert.deepStrictEqual(thead([{ label: 'A' }, { label: 'B' }]).map(r => r.length), [2]); // 舊範本沒有 levels → 單列

// 發布／分享網址跟著網站所在帳號
const { fb_repoUrls: urls } = require('./form-builder.js');
assert.deepStrictEqual(urls('alice.github.io', '/local_gov_budget/form-builder.html'),
    { upload: 'https://github.com/alice/local_gov_budget/upload/main/forms', site: 'https://alice.github.io/local_gov_budget/', api: 'https://api.github.com/repos/alice/local_gov_budget/contents/forms' });
assert.deepStrictEqual(urls('Bob-Lin.github.io', '/my-forms/'),
    { upload: 'https://github.com/bob-lin/my-forms/upload/main/forms', site: 'https://bob-lin.github.io/my-forms/', api: 'https://api.github.com/repos/bob-lin/my-forms/contents/forms' });
assert.deepStrictEqual(urls('carol.github.io', '/form-builder.html'),
    { upload: 'https://github.com/carol/carol.github.io/upload/main/forms', site: 'https://carol.github.io/', api: 'https://api.github.com/repos/carol/carol.github.io/contents/forms' });
assert.ok(urls('localhost', '/form-builder.html').upload.includes('/cormort/local_gov_budget/'));



// 項目層級與小計
const { fb_level: lvl, fb_parents: parents, fb_refsOtherRow: other, fb_headerMismatch: mismatch } = require('./form-builder.js');
assert.deepStrictEqual(['一、人事費', '（一）薪資', '1.本俸', '(1)主管', '說明', '二、業務費', '1.加班', '1.5億元'].map(lvl), [1, 2, 3, 4, 0, 1, 3, 0]);
assert.deepStrictEqual(parents(['一、人事費', '（一）薪資', '1.本俸', '2.加給', '說明', '（二）獎金', '二、業務費', '合計']), [-1, 0, 1, 1, -1, 0, -1, -1]);
// 真實查填表：沒編號的基金屬於上方「一、作業基金」；縮排的分基金屬於上方沒縮排的基金
assert.deepStrictEqual(parents(['合計', '一、作業基金', '行政院國家發展基金', '營建建設基金', '二、特別收入基金', '離島建設基金']), [-1, -1, 1, 1, -1, 4]);
assert.deepStrictEqual(parents(['交通作業基金', '  民航事業作業基金', '  國道公路建設管理基金', '經濟作業基金']), [-1, 0, 0, -1]);
assert.deepStrictEqual(parents(['甲', '乙', '丙']), [-1, -1, -1]); // 沒編號、沒縮排的清單都是同一層
assert.deepStrictEqual(parents(['1.甲', '(1)a', '2.乙']), [-1, 0, -1]); // 最上層不是「一、」也可以
assert.strictEqual(other('SUM(B3:B9)', 10), true);
assert.strictEqual(other('B10-C10', 10), false);
assert.strictEqual(other('[1]大基金!$M$41', 24), false); // 外部連結取值，不是合計列
assert.strictEqual(mismatch(['項目', '預算數'], ['項目', '預算數']), null);
assert.ok(mismatch(['項目', '預算數'], ['項目']).includes('欄位數量不符'));
assert.ok(mismatch(['項目', '預算數'], ['項目', '決算數']).includes('第 2 欄'));
assert.strictEqual(mismatch(['說明\r\n(執行率低於80%者)'], ['說明\n(執行率低於80%者)']), null); // 換行不同不算
const { fb_rollup: rollup } = require('./form-builder.js');
const rcols = [{ key: 'A', type: 'text' }, { key: 'B', type: 'num' }, { key: 'C', type: 'num', formula: 'B*2' }];
assert.deepStrictEqual(rollup([{ A: '一、甲', B: '' }, { A: '（一）a', B: 1 }, { A: '1.x', B: 2 }, { A: '2.y', B: '3' }, { A: '二、乙', B: 9 }], rcols).map(r => r.B), [5, 5, 2, '3', 9]);
const { fb_libraryEntries: lib } = require('./form-builder.js');
assert.deepStrictEqual(lib([{ name: 'README.md', type: 'file' }, { name: '乙表.xlsx', type: 'file' }, { name: '甲表.xlsx', type: 'file' }, { name: '甲表.json', type: 'file' }, { name: 'old', type: 'dir' }]),
    [{ name: '乙表', file: '乙表.xlsx' }, { name: '甲表', file: '甲表.json' }].sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant')));
// 彙整檢查：清單型（各填各的）才抓重複；經費表型（大家填同一批項目）不算
const { fb_itemOwners: owners, fb_sameNums: same } = require('./form-builder.js');
const oc = [{ key: 'A', type: 'text' }, { key: 'B', type: 'num' }, { key: 'C', type: 'num', formula: 'B*2' }];
const sub = (...rows) => ({ rows: rows.map(([A, B]) => ({ A, B })) });
let o = owners([sub(['一、作業基金', 9], ['甲基金', 5], ['乙基金', '']), sub(['一、作業基金', 0], ['甲基金', ''], ['乙基金', 7]), sub(['甲基金', 5], ['丙基金', 1], ['丁基金', 2])], oc);
assert.strictEqual(o.oneOwner, true);
assert.deepStrictEqual([...o.owners].map(([k, v]) => [k, v.length]), [['甲基金', 2], ['乙基金', 1], ['丙基金', 1], ['丁基金', 1]]); // 上層「一、」不算
o = owners([sub(['人事費', 1], ['業務費', 2]), sub(['人事費', 3], ['業務費', 4])], oc);
assert.strictEqual(o.oneOwner, false);
assert.strictEqual(same({ B: 5, C: 1 }, { B: '5', C: 2 }, oc), true);  // 公式欄不比
assert.strictEqual(same({ B: 5 }, { B: 6 }, oc), false);
// 多工作表格式；舊的單一工作表 JSON 轉成新格式
const { fb_normalize: norm } = require('./form-builder.js');
const oldJson = { template: { title: '經費表', sheet: '工作表1', columns: [{ key: 'A', label: '項目' }], caption: ['單位：千元'] }, rows: [{ A: '甲' }] };
assert.deepStrictEqual(norm(oldJson), { title: '經費表', sheets: [{ name: '工作表1', columns: [{ key: 'A', label: '項目' }], caption: ['單位：千元'] }], tables: { 工作表1: [{ A: '甲' }] } });
assert.strictEqual(norm({ template: { title: 'x', columns: [] } }).sheets[0].name, 'x'); // 更舊的沒有 sheet
const newJson = { template: { title: 't', sheets: [{ name: 'a', columns: [] }, { name: 'b', columns: [] }] }, tables: { a: [], b: [{}] } };
assert.deepStrictEqual(norm(newJson).sheets.map(s => s.name), ['a', 'b']);
assert.throws(() => norm({ template: {} }));
// 三層表頭：預算數 → 經常門／資本門 → 原編／核列，合計只到第二層
const g3 = [
    ['計畫名稱', '預算數', '', '', '', '', '執行數'],
    ['', '經常門', '', '資本門', '', '合計', ''],
    ['', '原編', '核列', '原編', '核列', '', ''],
    ['人事費', '10', '9', '5', '4', '', '8'],
];
const at3 = (r, c) => (g3[r] || [])[c] || '';
const m3 = [M(0, 0, 2, 0), M(0, 1, 0, 5), M(0, 6, 2, 6), M(1, 1, 1, 2), M(1, 3, 1, 4), M(1, 5, 2, 5)];
assert.deepStrictEqual(block(at3, m3, 0, 3, 0, 6), { hr: 0, he: 2 });
const lv3 = [0, 1, 2, 3, 4, 5, 6].map(c => levels(at3, m3, 0, 2, c));
assert.deepStrictEqual(lv3[2], ['預算數', '經常門', '核列']);
assert.deepStrictEqual(lv3[5], ['預算數', '合計']);
const h3 = thead(lv3.map(l => ({ label: l.join('－'), levels: l }))).map(row => row.map(x => `${x.text}:${x.colspan}x${x.rowspan}`));
assert.deepStrictEqual(h3, [['計畫名稱:1x3', '預算數:5x1', '執行數:1x3'], ['經常門:2x1', '資本門:2x1', '合計:1x2'], ['原編:1x1', '核列:1x1', '原編:1x1', '核列:1x1']]);
// 表名、單位列（第 1、2 列）有兩格以上的字也要略過；表頭從第 3 列開始
const capGrid = (r0, r1) => [r0, r1, ['計畫名稱', '預算數', '執行數'], ['甲', '1', '2']];
const capBlock = (g, mg = []) => block((r, c) => (g[r] || [])[c] || '', mg, 0, 3, 0, 2);
assert.deepStrictEqual(capBlock(capGrid(['○○表', '', ''], ['填表機關：', '', '單位：千元'])), { hr: 2, he: 2 });
assert.deepStrictEqual(capBlock(capGrid(['○○表', '', ''], ['機關：', '填表人：', '單位：元'])), { hr: 2, he: 2 });
assert.deepStrictEqual(capBlock(capGrid(['○○計畫經費表', '', '115.10.24'], ['', '', '單位：千元']), [M(0, 0, 0, 1)]), { hr: 2, he: 2 }); // 表名合併＋日期
assert.deepStrictEqual(capBlock(capGrid(['○○表', '', ''], ['', '', '單位：千元'])), { hr: 2, he: 2 });
// 範本庫分組：index.json 的分類、隱藏；沒寫的歸「未分類」排最後；index.json 本身不列
const { fb_libraryGroups: groups } = require('./form-builder.js');
const libEntries = lib([{ name: 'index.json', type: 'file' }, { name: '甲.json', type: 'file' }, { name: '乙.xlsx', type: 'file' }, { name: '丙.json', type: 'file' }, { name: '丁.json', type: 'file' }]);
assert.ok(!libEntries.some(e => e.name === 'index'));
const idx = { templates: { 甲: { category: '預算', description: 'd' }, 乙: { category: '決算' }, 丙: { category: '預算', hidden: true } } };
assert.deepStrictEqual(groups(libEntries, idx).map(g => [g.category, g.items.map(i => i.name)]), [['決算', ['乙']], ['預算', ['甲']], ['未分類', ['丁']]]);
assert.deepStrictEqual(groups(libEntries, idx, true).find(g => g.category === '預算').items.map(i => i.name).sort(), ['丙', '甲']);
assert.deepStrictEqual(groups(libEntries).map(g => g.category), ['未分類']);
// 範本庫查詢
const { fb_libraryMatch: match } = require('./form-builder.js');
const it = { name: '114決算購建固定資產執行情形', meta: { category: '決算', description: '各基金填自己那一列', year: '114' } };
assert.ok(match(it, ''));
assert.ok(match(it, '購建 114'));            // 多個字都要符合
assert.ok(match(it, '基金'));                // 比對說明
assert.ok(!match(it, '預算'));
assert.ok(match(it, '', '決算') && !match(it, '', '預算'));
assert.ok(match({ name: 'x', meta: {} }, '', '未分類'));
assert.ok(match({ name: 'ABC表' }, 'abc'));  // 不分大小寫，沒有 meta 也行
// 全形、文字格式的數字
assert.strictEqual(num('（５００）'), '-500');
assert.strictEqual(num('１２３'), '123');
assert.strictEqual(num('１，２３４．５'), '1234.5');
assert.strictEqual(num('－５００'), '-500');
assert.strictEqual(num('無'), '無');
assert.strictEqual(num(1234), '1234');
const { fb_num: toNum, fb_guessYear: guessYear } = require('./form-builder.js');
assert.strictEqual(toNum('1,234'), 1234);   // parseFloat 會變成 1
assert.strictEqual(toNum('(300)'), -300);
assert.strictEqual(toNum('無'), 0);
assert.strictEqual(toNum(''), 0);
assert.strictEqual(toNum(12.5), 12.5);
const { fb_cleanRows: cleanRows } = require('./form-builder.js');
{
    const rows = [{ A: '甲', B: '1,234', C: '(300)', D: 'x', _r: 2 }, { A: '乙', B: '無', C: 5, D: '', _r: 3 }];
    const w = cleanRows(rows, [{ key: 'A', type: 'text' }, { key: 'B', type: 'num' }, { key: 'C', type: 'num' }, { key: 'D', type: 'text' }]);
    assert.deepStrictEqual(rows, [{ A: '甲', B: 1234, C: -300, D: 'x' }, { A: '乙', B: '', C: 5, D: '' }]);
    assert.strictEqual(w, '1 格不是數字，已當空白：B3「無」');
    assert.strictEqual(cleanRows([{ A: '丙', B: '無' }], [{ key: 'A', type: 'text' }, { key: 'B', label: '原編', type: 'num' }]), '1 格不是數字，已當空白：「丙」原編「無」');
    assert.strictEqual(cleanRows([{ A: 'a', B: 1 }], [{ key: 'A', type: 'text' }, { key: 'B', type: 'num' }]), null);
}
// 年度預設
const oct2026 = new Date(2026, 9, 1);
assert.strictEqual(guessYear('114決算購建固定資產執行情形', '', oct2026), '114');
assert.strictEqual(guessYear('114決算', '113', oct2026), '113');   // 範本庫設定優先
assert.strictEqual(guessYear('業務計畫範例', undefined, oct2026), '115');
assert.strictEqual(guessYear('2026表', undefined, oct2026), '115'); // 四位數不是民國年
// 單一合併標題列（A 欄空、B1:C1「金額」）是上層表頭，不是表名
{
    const g = [['', '金額', ''], ['項目名稱', '原編', '增編'], ['甲', '1', '2']];
    const at = (r, c) => g[r]?.[c] ?? '', mg = [M(0, 1, 0, 2)];
    assert.deepStrictEqual(block(at, mg, 0, 2, 0, 2), { hr: 0, he: 1 });
    assert.deepStrictEqual([0, 1, 2].map(c => levels(at, mg, 0, 1, c)), [['項目名稱'], ['金額', '原編'], ['金額', '增編']]);
    // 從 A 欄開始合併的仍是表名
    const t = [['○○表', '', ''], ['項目名稱', '原編', '增編']];
    assert.deepStrictEqual(block((r, c) => t[r]?.[c] ?? '', [M(0, 0, 0, 2)], 0, 1, 0, 2), { hr: 1, he: 1 });
}
// API 讀不到時以 index.json 備援
const { fb_libraryFromIndex: fromIndex } = require('./form-builder.js');
assert.deepStrictEqual(fromIndex({ templates: { 乙: {}, 甲: { hidden: true } } }).map(e => [e.name, e.file]), [['乙', null], ['甲', null]]);
assert.deepStrictEqual(fromIndex({}), []);
// 調查狀態與新欄位查詢
const { fb_libraryStatus: status } = require('./form-builder.js');
const D = '2026-10-24';
assert.strictEqual(status({}, D), '');
assert.strictEqual(status({ start: '2026-11-01' }, D), 'upcoming');
assert.strictEqual(status({ start: '2026-10-01', deadline: '2026-11-30' }, D), 'open');
assert.strictEqual(status({ deadline: '2026-10-24' }, D), 'open');      // 截止日當天還在進行
assert.strictEqual(status({ deadline: '2026-10-23' }, D), 'closed');
const it2 = { name: '基金調查', meta: { unit: '主計總處基金預算處', purpose: '彙編決算書表', deadline: '2026-10-01' } };
assert.ok(match(it2, '基金預算處 決算書'));
assert.ok(match(it2, '', '', 'closed', D) && !match(it2, '', '', 'open', D));
// 調查專案：代號檢查、清單排序、繳交對照
const { fb_badName: bad, fb_surveyEntries: surveys, fb_submissions: subsOf } = require('./form-builder.js');
assert.strictEqual(bad('115決算購建'), null);
assert.strictEqual(bad('114.1調查'), null);   // 單一個點可以
for (const x of ['', ' ', 'a/b', 'a b', '../x', '.x', 'a?b', 'a#b', 'index']) assert.ok(bad(x), x);
const sv = surveys({ surveys: { 舊: { start: '2025-01-01' }, 新: { start: '2026-10-01' }, 藏: { hidden: true, start: '2026-12-01' } } });
assert.deepStrictEqual(sv.map(s => s.name), ['新', '舊']);
assert.ok(match({ name: '115決算購建', meta: { title: '115年度決算購建固定資產', form: '114決算購建固定資產執行情形' } }, '年度決算'));
const so = subsOf(['臺北市', '新北市', '桃園市'], ['台北市', '新 北 市', '高雄市']);
assert.deepStrictEqual(so.rows.map(r => r.org + ':' + r.status), ['臺北市:done', '新北市:done', '桃園市:missing', '高雄市:extra']);
assert.deepStrictEqual([so.done, so.total], [2, 3]);
assert.deepStrictEqual(subsOf([], ['甲']).rows.map(r => r.status), ['extra']);
// 今天用本地日期：台灣 10/25 凌晨 3 點（UTC 還是 10/24）也算 10/25
const { fb_today: today } = require('./form-builder.js');
assert.strictEqual(today(new Date(2026, 9, 25, 3, 0)), '2026-10-25');
assert.strictEqual(today(new Date(2026, 0, 5, 23, 59)), '2026-01-05');
assert.strictEqual(status({ start: '2026-10-25' }, today(new Date(2026, 9, 25, 3, 0))), 'open');
assert.strictEqual(status({ deadline: '2026-10-24' }, today(new Date(2026, 9, 25, 3, 0))), 'closed');
// 數字解析（借鏡 multi-xlsx_merge 的 toNumber）
assert.strictEqual(num('1,5'), '1,5');              // 不是千分位，不能變成 15
assert.strictEqual(num('1,234,567'), '1234567');
assert.strictEqual(num('1,234元'), '1234');
assert.strictEqual(num('NT$1,234'), '1234');
assert.strictEqual(num('3千元'), '3');
assert.strictEqual(num('1,234*'), '1234');          // 附註星號
assert.strictEqual(num('１、２３４'), '1234');        // 「、」當千分位
assert.strictEqual(num('−1,234'), '-1234');         // 數學減號
assert.strictEqual(num('—5'), '-5');
assert.strictEqual(num('-(5)'), '-(5)');            // 怪寫法不猜，原樣讓使用者看到
assert.strictEqual(num('+12'), '12');
assert.strictEqual(num('55%'), '55%');              // 百分比不當數字
assert.strictEqual(num('2024/01/01'), '2024/01/01');
assert.strictEqual(num('1e5'), '1e5');
assert.strictEqual(num('.5'), '0.5');
assert.strictEqual(num('(0)'), '0');                // 不要 -0
// 小數檢查：公式欄、文字欄、空白不算
const { fb_decimalCells: decimals } = require('./form-builder.js');
{
    const sheets = [{ name: 'S', columns: [{ key: 'A', type: 'text' }, { key: 'B', type: 'num' }, { key: 'C', type: 'num', formula: 'B/2' }, { key: 'D', type: 'num' }] }];
    assert.deepStrictEqual(decimals(sheets, { S: [{ A: '1.5', B: 100, C: 50.5, D: '' }, { A: 'x', B: '1,234.5', D: 0.25 }] }), { count: 2, sample: 1234.5 });
    assert.deepStrictEqual(decimals(sheets, { S: [{ A: 'x', B: 100, C: 0.5 }] }), { count: 0, sample: null });
    assert.deepStrictEqual(decimals(sheets, {}), { count: 0, sample: null });
}
// 機關名稱比對（借鏡 excel_merge）：別名、去後綴、前 4 字＋字數、包含；模糊比對只在唯一時成立
const { fb_orgMatcher: matcher } = require('./form-builder.js');
{
    const list = ['國立臺灣大學校務基金', '國立臺灣大學附設醫院作業基金', '台灣電力股份有限公司', '國立臺灣師範大學校務基金', '住宅基金'];
    const aliases = { 台電: '台灣電力股份有限公司', 臺大: '國立臺灣大學校務基金' };
    const m = matcher(list, aliases);
    assert.deepStrictEqual(m('國立台灣大學校務基金'), { i: 0, level: 1 });       // 台→臺
    assert.deepStrictEqual(m(' 國立臺灣大學（校務基金）'), { i: 0, level: 1 });   // 空白、括號
    assert.deepStrictEqual(m('台電'), { i: 2, level: 1 });                      // 別名
    assert.deepStrictEqual(m('臺大'), { i: 0, level: 1 });
    assert.deepStrictEqual(m('國立臺灣大學'), { i: 0, level: 2 });              // 去「校務基金」
    assert.deepStrictEqual(m('國立臺灣師範大學校務基今'), { i: 3, level: 3 });   // 前 4 字＋字數（錯字）
    assert.deepStrictEqual(m('住宅基金（內政部國土管理署）'), { i: 4, level: 4 }); // 名單名稱包含在內
    assert.deepStrictEqual(m('國立臺灣大學醫學院'), { i: -1 });                 // 不硬配
    assert.deepStrictEqual(m(''), { i: -1 });
    assert.deepStrictEqual(matcher(['國立甲大學校務基金', '國立甲大學附設醫院作業基金'])('國立甲大學'), { i: 0, level: 2 });
    assert.deepStrictEqual(matcher(['國立甲大學校務基金', '國立甲大學乙務基金'])('國立甲大學丙務基金'), { i: -1 }); // 前 4 字＋字數有兩個候選 → 不猜
}
{
    // 完全相同的先配，模糊的不能搶名額；同一機關兩份 → 後面那份不在名單
    const r = subsOf(['國立臺灣大學校務基金', '住宅基金'], ['國立臺灣大學', '國立臺灣大學校務基金', '住宅基金', '住宅基金'], {});
    assert.deepStrictEqual(r.rows.map(x => [x.org, x.status, x.index, !!x.fuzzy]), [
        ['國立臺灣大學校務基金', 'done', 1, false], ['住宅基金', 'done', 2, false], ['國立臺灣大學', 'extra', 0, false], ['住宅基金', 'extra', 3, false]]);
    const f = subsOf(['國立臺灣大學校務基金'], ['臺大'], { 臺大: '國立臺灣大學校務基金' });
    assert.deepStrictEqual([f.done, f.rows[0].fuzzy], [1, undefined]);       // 別名算完全相同
    const g = subsOf(['國立臺灣大學校務基金'], ['國立臺灣大學']);
    assert.deepStrictEqual([g.done, g.rows[0].fuzzy], [1, true]);
}
console.log('form-builder: all checks passed');
