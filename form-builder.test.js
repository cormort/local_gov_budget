// node form-builder.test.js — 公式轉換與求值的自我檢查
const assert = require('assert');
const { fb_translateFormula: tr, fb_evalFormula: ev } = require('./form-builder.js');

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

console.log('form-builder: all checks passed');
