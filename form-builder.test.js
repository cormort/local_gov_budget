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

console.log('form-builder: all checks passed');
