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
    { upload: 'https://github.com/alice/local_gov_budget/upload/main/forms', site: 'https://alice.github.io/local_gov_budget/' });
assert.deepStrictEqual(urls('Bob-Lin.github.io', '/my-forms/'),
    { upload: 'https://github.com/bob-lin/my-forms/upload/main/forms', site: 'https://bob-lin.github.io/my-forms/' });
assert.deepStrictEqual(urls('carol.github.io', '/form-builder.html'),
    { upload: 'https://github.com/carol/carol.github.io/upload/main/forms', site: 'https://carol.github.io/' });
assert.ok(urls('localhost', '/form-builder.html').upload.includes('/cormort/local_gov_budget/'));

console.log('form-builder: all checks passed');
