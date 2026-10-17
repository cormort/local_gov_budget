#!/usr/bin/env node
/**
 * verify-xlsx-mapping.mjs — 「官方 xlsx → 系統表格」映射的回歸測試（profile 來自 vendor/forms/，由 forms repo 產生）
 *
 * 跑法：npm run test:mapping
 *
 * 驗的是「值不會錯」這件事，不是 UI：
 *   1. 官方檔 6 個基金都對得上、欄位分區正確
 *   2. 以現行 sample_fund_*.json 為基準，映射結果為 75 命中 / 0 未匹配 / 0 加總警告
 *   3. 已知的那一筆資料缺口（農損基金「綠色環境給付計畫」）會被補成 12,759,005，
 *      且補完後 4 個計畫確實等於官方「基金用途」總額
 *   4. 檔案改一個數字，差異偵測要抓到且只抓到那一筆
 *   5. 「其他」「購建固定資產」這種人工分類列永不被映射
 *   6. applyOps 不改動傳入的資料（純函式）
 */
import { createRequire } from 'node:module';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import assert from 'node:assert/strict';

const require = createRequire(import.meta.url);
const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');

let XLSX;
try {
  XLSX = require('xlsx');
} catch (e) {
  console.error('找不到 xlsx 套件，請先執行 npm install');
  process.exit(2);
}
const core = require(join(root, 'xlsx-import.js'));
require(join(root, 'vendor/forms/xlsx-profiles.js')); // 產生檔會設定 globalThis.LGB_XLSX_PROFILES
const profile = globalThis.LGB_XLSX_PROFILES.FNGBRB5300;

const XLSX_FILE = join(root, 'FNGBRB5300_20260520173956.xlsx');

// ---------- 讀基準資料（等同瀏覽器執行時 fetch sample_fund_*.json） ----------
function loadBaseTables() {
  const tables = {};
  for (const f of readdirSync(root).filter(n => /^sample_fund_.*\.json$/.test(n)).sort()) {
    const d = JSON.parse(readFileSync(join(root, f), 'utf8'));
    for (const [tid, rows] of Object.entries(d.tables || {})) {
      if (!tid.startsWith('plan_')) continue;
      assert.ok(!tables[tid], `表格 ${tid} 重複出現於多個 sample 檔`);
      tables[tid] = rows.map(r => ({ ...r }));
    }
  }
  return tables;
}

function readAoa(file) {
  const wb = XLSX.read(readFileSync(file), { type: 'buffer' });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
}

const checks = [];
function check(name, fn) {
  try { fn(); checks.push([name, true]); console.log(`  ✓ ${name}`); }
  catch (e) { checks.push([name, false]); console.log(`  ✗ ${name}\n      ${e.message}`); }
}

console.log('FNGBRB5300 → 系統表格 映射回歸測試\n');

const aoa = readAoa(XLSX_FILE);
const parsed = core.parseSheet(aoa, profile);
const baseTables = loadBaseTables();
const base = core.mapData({ tables: baseTables }, parsed, profile);

check('6 個基金全部對到官方檔的列', () => {
  assert.deepEqual(parsed.missingFunds, []);
  assert.equal(Object.keys(parsed.funds).length, 6);
});

check('欄位分區：來源 77 欄、用途 23 欄', () => {
  assert.equal(parsed.columns.src.length, 77);
  assert.equal(parsed.columns.use.length, 23);
});

check('12 張 business plan 表都在且都有命中', () => {
  const expect = [];
  for (const code of Object.keys(profile.funds)) {
    expect.push(`plan_${code}_src`, `plan_${code}_use`);
  }
  assert.deepEqual(base.missingTables, []);
  const hit = new Set(base.ops.map(o => o.tableId));
  for (const tid of expect) assert.ok(hit.has(tid), `${tid} 沒有任何命中`);
});

check('基準資料：75 命中、0 未匹配、0 加總警告、1 筆待修正', () => {
  assert.equal(base.ops.length, 75);
  assert.equal(base.unmatched.length, 0, JSON.stringify(base.unmatched));
  assert.equal(base.warnings.length, 0, JSON.stringify(base.warnings));
  assert.equal(base.stats.confirmRequired, 0);
  assert.equal(base.stats.changed, 1);
});

check('人工分類列（其他／購建固定資產）永不被映射', () => {
  for (const o of base.ops) {
    assert.notEqual(o.name, '其他');
    assert.notEqual(o.name, '購建固定資產');
  }
  assert.ok(base.ignored.length > 0, '應該要有被忽略的列');
});

check('農損基金「綠色環境給付計畫」＝ 12,759,005（別名對到官方改名後的欄位）', () => {
  const op = base.ops.find(o => o.tableId === 'plan_loss_use' && o.name === '綠色環境給付計畫');
  assert.ok(op, '找不到該列');
  assert.equal(op.how, 'alias');
  assert.equal(op.value, 12759005);
  assert.equal(op.oldRaw, '');
});

check('補完後農損基金 4 個計畫加總 = 官方基金用途 16,695,143', () => {
  const loss = baseTables.plan_loss_use;
  const official = parsed.funds.loss.use['基金用途'];
  const sum = loss
    .filter(r => (parseInt(r.level, 10) || 0) === 1)
    .reduce((acc, r) => {
      const op = base.ops.find(o => o.name === r.name && o.tableId === 'plan_loss_use');
      const v = op ? op.value : Number(String(r.bud114 || '').replace(/,/g, ''));
      return acc + (isFinite(v) ? v : 0);
    }, 0);
  assert.equal(sum, official);
});

check('母公司列加總與官方值一致（無加總警告）', () => {
  assert.equal(base.warnings.length, 0);
});

check('短名稱不做寬鬆比對：「其他」→ too-short', () => {
  const r = core.resolveColumn('其他', parsed.columns.use, profile);
  assert.equal(r.column, null);
  assert.equal(r.how, 'too-short');
});

check('完全不存在的名稱 → unmatched，不會硬套最接近的欄位', () => {
  const r = core.resolveColumn('不存在的計畫名稱XYZ', parsed.columns.use, profile);
  assert.equal(r.column, null);
  assert.equal(r.how, 'unmatched');
});

check('檔案改一個數字：只抓到那一筆差異', () => {
  const mutated = aoa.map(row => row.slice());
  const headers = (aoa[0] || []).map(core.normalizeHeader);
  const col = headers.indexOf('提升農業經營及發展計畫(本年度預算數)') >= 0
    ? headers.indexOf('提升農業經營及發展計畫(本年度預算數)')
    : headers.findIndex(h => h && h.indexOf('提升農業經營及發展計畫') === 0);
  assert.ok(col > 0, '找不到目標欄位');
  const r = mutated.findIndex(row => core.unify(String(row[0] || '').replace(/\u3000/g, '')) === core.unify('農業發展基金'));
  assert.ok(r > 0, '找不到農發基金列');
  mutated[r][col] = 19000000000; // 190 億 → 1,900 萬千元

  const res = core.mapData({ tables: baseTables }, core.parseSheet(mutated, profile), profile);
  const changed = res.ops.filter(o => o.changed);
  assert.equal(changed.length, 2, `應為 2 筆（原資料缺口 + 本次修改），實際 ${changed.length}`);
  const op = changed.find(o => o.name.indexOf('提升農業經營') === 0);
  assert.equal(op.value, 19000000);
  assert.equal(op.tableId, 'plan_agri_use');
});

check('applyOps 是純函式：不改動傳入資料，只改選取的欄位', () => {
  const before = JSON.parse(JSON.stringify(baseTables));
  const selected = base.ops.filter(o => o.how !== 'contains' && o.how !== 'fuzzy');
  const next = core.applyOps({ tables: baseTables, meta: { fund: 'x' } }, selected, profile);
  assert.deepEqual(baseTables, before, '傳入的 tables 被改動了');
  const after = next.tables.plan_loss_use.find(r => r.name === '綠色環境給付計畫');
  assert.equal(after.bud114, '12759005');
  // 未選取的列不應被改寫
  const touched = new Set(selected.map(o => `${o.tableId}#${o.index}`));
  for (const tid of Object.keys(baseTables)) {
    baseTables[tid].forEach((row, i) => {
      if (touched.has(`${tid}#${i}`)) return;
      assert.deepEqual(next.tables[tid][i], row, `${tid}[${i}] 不該被改`);
    });
  }
});

check('單位換算：元 → 千元', () => {
  assert.equal(core.scaleValue(630000000, 1000), 630000);
  assert.equal(core.scaleValue(1234, 1000), 1.234);
});

check('報告 JSON 可序列化且含關鍵欄位', () => {
  const rep = core.buildReport(parsed, base, profile, 'FNGBRB5300.xlsx');
  const round = JSON.parse(JSON.stringify(rep));
  assert.equal(round.profile, 'FNGBRB5300');
  assert.equal(round.scale, 1000);
  assert.equal(round.ops.length, 75);
  assert.ok(round.generatedAt);
  assert.equal(round.stats.changed, 1);
});

const failed = checks.filter(([, ok]) => !ok);
console.log(`\n${checks.length - failed.length} / ${checks.length} 通過`);
if (failed.length) {
  console.error(`失敗：${failed.map(([n]) => n).join('、')}`);
  process.exit(1);
}

// 附帶輸出：目前這份官方檔對現行資料的差異摘要（人看的）
console.log('\n— 現行資料 vs 官方檔（僅列有變動者）—');
for (const o of base.ops.filter(x => x.changed)) {
  console.log(`  ${o.tableId}  ${o.name}  →  ${o.column}`);
  console.log(`      ${o.oldRaw === '' ? '(空白)' : o.oldRaw} → ${o.value}  [${o.how}]`);
}
console.log(`\n官方檔：${XLSX_FILE.replace(root + '/', '')}（${Math.round(statSync(XLSX_FILE).size / 1024)} KB）`);
