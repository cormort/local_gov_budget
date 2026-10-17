#!/usr/bin/env node
/**
 * sync-forms.mjs — 把 agri_special_fund_forms 的表單定義與資料，轉成瀏覽器可直接載入的產生檔
 *
 *   npm run sync:forms     # 重新產生 vendor/forms/*.js
 *   npm run check:forms    # 只檢查有沒有漂移（CI 用，不一致就 exit 1）
 *
 * 為什麼要轉成 .js：本專案實際以 `file://` 開啟，fetch() 讀不到本地 JSON，
 * CSP 也設成 connect-src 'self'。用 <script> 載入產生檔可以同時支援 file:// 與 http。
 *
 * 產生檔（請勿手改）：
 *   vendor/forms/agri-special-fund.js  → window.LGB_FORMS
 *   vendor/forms/sample-data.js        → window.LGB_SAMPLE_DATA
 *   vendor/forms/xlsx-profiles.js      → window.LGB_XLSX_PROFILES
 */
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execSync } from 'node:child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const check = args.includes('--check');
const fromIdx = args.indexOf('--from');
const formsRepo = resolve(fromIdx >= 0 ? args[fromIdx + 1] : join(root, '..', 'agri_special_fund_forms'));

const OUT_DIR = join(root, 'vendor/forms');
const FORMS_IN = join(formsRepo, 'forms/agri-special-fund.json');
const PROFILE_IN = join(formsRepo, 'profiles/FNGBRB5300.json');

function die(msg) { console.error(`✗ ${msg}`); process.exit(2); }
function readJson(p) {
  if (!existsSync(p)) die(`找不到 ${p}（是否還沒 clone agri_special_fund_forms？可用 --from 指定路徑）`);
  return JSON.parse(readFileSync(p, 'utf8'));
}

const forms = readJson(FORMS_IN);

// 資料來源版本（方便追朔產生檔對應哪個 commit）
let rev = 'unknown';
try {
  rev = execSync(`git -C ${JSON.stringify(formsRepo)} rev-parse --short HEAD`, { encoding: 'utf8' }).trim();
} catch { /* 不是 git repo 也沒關係 */ }

// ---- 合併 6 個分基金的資料（headcount 六基金共用，只取第一份；其餘照檔串接）----
const merged = {};
const codes = forms.funds.map(f => f.code);
for (const code of codes) {
  const rel = forms.data && forms.data[code];
  if (!rel) die(`forms.data 缺少基金 ${code} 的資料檔`);
  const doc = readJson(join(formsRepo, rel));
  for (const [tid, rows] of Object.entries(doc.tables || {})) {
    if (tid === 'headcount') {
      if (!merged[tid]) merged[tid] = rows.map(r => ({ ...r }));
    } else {
      merged[tid] = (merged[tid] || []).concat(rows.map(r => ({ ...r })));
    }
  }
}

// 檢查：資料裡的表格都要在 forms 定義內
const defined = new Set(forms.tables.map(t => t.id));
for (const tid of Object.keys(merged)) {
  if (!defined.has(tid)) die(`資料含未定義表格 ${tid}`);
}

const profile = existsSync(PROFILE_IN) ? readJson(PROFILE_IN) : null;

const banner = (what) => `/* 產生檔，請勿手改：${what}
 * 來源：agri_special_fund_forms@${rev}（${formsRepo}）
 * 產生：npm run sync:forms（tools/sync-forms.mjs）
 */
`;

// 讓產生檔在瀏覽器（window）與 Node 測試（globalThis）都能載入
const wrap = (globalName, value) =>
  `(function (root) { root.${globalName} = ${JSON.stringify(value)}; })` +
  "(typeof window !== 'undefined' ? window : globalThis);\n";

const outputs = {
  'agri-special-fund.js': banner('表單定義') + wrap('LGB_FORMS', forms),
  'sample-data.js': banner('預設資料（合併 6 個分基金的 data/*.json）') + wrap('LGB_SAMPLE_DATA', merged),
  'xlsx-profiles.js': banner('官方 xlsx 映射 profile') +
    wrap('LGB_XLSX_PROFILES', profile ? { [profile.id]: profile } : {})
};

mkdirSync(OUT_DIR, { recursive: true });

let drift = [];
for (const [name, content] of Object.entries(outputs)) {
  const p = join(OUT_DIR, name);
  const prev = existsSync(p) ? readFileSync(p, 'utf8') : null;
  const normalize = (s) => (s || '').replace(/agri_special_fund_forms@[0-9a-f]+/g, '@rev');
  if (check) {
    if (normalize(prev) !== normalize(content)) drift.push(name);
    continue;
  }
  if (prev !== content) {
    writeFileSync(p, content, 'utf8');
    console.log(`  更新 ${name}（${Buffer.byteLength(content)} bytes）`);
  } else {
    console.log(`  未變更 ${name}`);
  }
}

const tableCount = Object.keys(merged).length;
const rowCount = Object.values(merged).reduce((a, rows) => a + rows.length, 0);

if (check) {
  if (drift.length) {
    console.error('✗ vendor/forms 與 forms repo 不同步：' + drift.join('、'));
    console.error('  請執行 npm run sync:forms');
    process.exit(1);
  }
  console.log(`✓ 同步（agri_special_fund_forms@${rev}）：${forms.tables.length} 張表、${tableCount} 份資料、${rowCount} 列`);
} else {
  console.log(`✓ 完成：${forms.tables.length} 張表、${tableCount} 份資料、${rowCount} 列（來源 @${rev}）`);
}
