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
console.log('form-builder: all checks passed');
