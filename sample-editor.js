const FUNDS = [
  { short: '農發', code: 'agri',     id: 'plan_agri',     name: '農業發展基金' },
  { short: '林務', code: 'forest',   id: 'plan_forest',   name: '林務發展及造林基金' },
  { short: '天災', code: 'disaster', id: 'plan_disaster', name: '農業天然災害救助基金' },
  { short: '漁發', code: 'fish',     id: 'plan_fish',     name: '漁業發展基金' },
  { short: '農損', code: 'loss',     id: 'plan_loss',     name: '農產品受進口損害救助基金' },
  { short: '再生', code: 'renewal',  id: 'plan_renewal',  name: '農村再生基金' },
];

const COL_LABELS = {
  dec112: '113 年度決算',
  dec113: '114 年度決算',
  bud114: '115 年度預算',
};

let loadedData = null;     // Whole JSON object (preserves headcount, personnel_cost, meta, etc.)
let pendingPreview = null; // Parsed preview result

const $  = (q) => document.querySelector(q);
const $$ = (q) => document.querySelectorAll(q);

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function setStatus(id, msg, cls) {
  const el = $('#' + id);
  el.textContent = msg;
  el.className = 'status' + (cls ? ' ' + cls : '');
}

function currentFund() {
  const code = $('#sel-fund').value;
  return FUNDS.find(f => f.code === code) || FUNDS[0];
}

function targetTableId() {
  const f = currentFund();
  const target = $('#sel-target').value;
  if (target === 'plan_src') return `${f.id}_src`;
  if (target === 'plan_use') return `${f.id}_use`;
  return `control_${f.code}`;
}

function isPlanTarget() {
  return $('#sel-target').value.startsWith('plan');
}

function parseNum(s) {
  if (s === '' || s == null) return null;
  const n = parseFloat(String(s).replace(/[,\s]/g, ''));
  return isNaN(n) ? null : n;
}

function fmtNum(v) {
  if (v === '' || v == null) return '';
  const n = parseFloat(String(v).replace(/,/g, ''));
  if (isNaN(n)) return String(v);
  return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
}

// 全形空格 leading 計層級
function inferLevel(rawName) {
  let lv = 1;
  for (const ch of rawName) {
    if (ch === '　') lv++;
    else if (ch === ' ' || ch === '\t') continue; // 半形空白 / Tab 不計層
    else break;
  }
  // L0 顯式辨識
  const stripped = rawName.replace(/^[\s　]+/, '').trim();
  if (/^甲、|^乙、/.test(stripped)) return { level: 0, name: stripped };
  return { level: lv, name: stripped };
}

// 解析 textarea 內容：每行 [name] <tab/多空白> [value]
function parsePaste(text) {
  const lines = text.split(/\r?\n/);
  const rows = [];
  lines.forEach((raw, i) => {
    if (!raw.trim()) return;
    // 拆名稱與值：tab 優先，否則用 2+ 空白
    let name, value;
    if (raw.includes('\t')) {
      const parts = raw.split('\t');
      name = parts[0];
      value = parts.slice(1).join('\t').trim();
    } else {
      const m = raw.match(/^(.+?)\s{2,}(.+)$/);
      if (m) { name = m[1]; value = m[2].trim(); }
      else { name = raw; value = ''; }
    }
    const num = parseNum(value);
    rows.push({ raw, name, value, num, lineNo: i + 1 });
  });
  return rows;
}

// ===== 預覽 / 套用 =====
function buildPreviewForPlan(parsed, col) {
  // 業務計畫：覆蓋整張表
  const out = parsed.map(p => {
    const { level, name } = inferLevel(p.name);
    const row = { level, name };
    if (p.num != null) row[col] = String(p.num);
    return row;
  }).filter(r => r.name);
  return { mode: 'replace', rows: out };
}

function buildPreviewForControl(parsed, col, existing) {
  // 管制項目：依名稱比對；存在→更新欄；不存在→新增
  const map = new Map((existing || []).map((r, i) => [r.name?.trim(), i]));
  const rows = (existing || []).map(r => ({ ...r, _status: 'unchanged' }));
  let updated = 0, added = 0;
  parsed.forEach(p => {
    const key = p.name.replace(/^[\s　]+/, '').trim();
    if (!key) return;
    if (map.has(key)) {
      const idx = map.get(key);
      rows[idx] = { ...rows[idx], [col]: p.num != null ? String(p.num) : '', _status: 'updated' };
      updated++;
    } else {
      rows.push({ name: key, [col]: p.num != null ? String(p.num) : '', _status: 'added' });
      added++;
    }
  });
  return { mode: 'merge', rows, updated, added };
}

function renderPreview(preview, col) {
  const area = $('#preview-area');
  if (!preview || !preview.rows?.length) {
    area.innerHTML = '<p class="hint">尚無資料；請先載入 JSON 並貼上資料按「預覽影響」。</p>'
      + '<div class="preview-edit-toolbar"><button class="btn" id="btn-add-row" disabled>＋ 新增列</button></div>';
    return;
  }
  const showCol = COL_LABELS[col];
  const isPlan = isPlanTarget();
  const hdrs = isPlan
    ? ['Lv', '項目名稱', '113決算', '114決算', '115預算', '']
    : ['項目名稱', '113決算', '114決算', '115預算', ''];

  const lvOpts = (cur) => [0,1,2,3,4]
    .map(lv => `<option value="${lv}"${String(cur ?? 1)===String(lv)?' selected':''}>L${lv}</option>`)
    .join('');

  const rowHTML = preview.rows.map((r, idx) => {
    const lvClass = isPlan ? `lv-${r.level ?? 1}` : '';
    const stClass = r._status === 'updated' ? 'updated' : (r._status === 'added' ? 'added' : '');
    const cls = [lvClass, stClass].filter(Boolean).join(' ');
    const nameCell = `<td><input type="text" data-key="name" data-row="${idx}" value="${esc(r.name ?? '')}"></td>`;
    const numCell = (k) => `<td class="num"><input type="text" inputmode="numeric" class="cell-num" data-key="${k}" data-row="${idx}" value="${esc(fmtNum(r[k]))}"></td>`;
    const opCell = `<td class="ops"><button type="button" class="btn-del" data-action="del" data-row="${idx}" title="刪除此列">✕</button></td>`;
    const lvCell = `<td><select class="cell-lv" data-key="level" data-row="${idx}">${lvOpts(r.level)}</select></td>`;
    const cells = isPlan
      ? [lvCell, nameCell, numCell('dec112'), numCell('dec113'), numCell('bud114'), opCell]
      : [nameCell, numCell('dec112'), numCell('dec113'), numCell('bud114'), opCell];
    return `<tr class="${cls}" data-row="${idx}">${cells.join('')}</tr>`;
  }).join('');

  const summary = preview.mode === 'replace'
    ? `<p class="hint">模式：<b>覆蓋整張表</b>，共 ${preview.rows.length} 列。寫入欄：<b>${showCol}</b>。</p>`
    : `<p class="hint">模式：<b>依名稱寫值</b>，更新 ${preview.updated} 列、新增 ${preview.added} 列、保留 ${preview.rows.length - preview.updated - preview.added} 列。寫入欄：<b>${showCol}</b>。</p>`;

  area.innerHTML = summary
    + `<table class="preview-table editable"><thead><tr>${hdrs.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rowHTML}</tbody></table>`
    + `<div class="preview-edit-toolbar"><button class="btn" id="btn-add-row">＋ 新增列</button><span class="hint">編輯／新增／刪除後請按「✏️ 套用到記憶中」儲存到匯出資料。</span></div>`;

  // Wire up inline-edit handlers
  area.querySelectorAll('input[data-key], select[data-key]').forEach(el => {
    el.addEventListener('input', onCellEdit);
    if (el.tagName === 'SELECT') el.addEventListener('change', onCellEdit);
    if (el.classList.contains('cell-num')) el.addEventListener('blur', onCellNumBlur);
  });
  area.querySelectorAll('button[data-action="del"]').forEach(btn => {
    btn.addEventListener('click', () => deleteRow(parseInt(btn.dataset.row, 10)));
  });
  const addBtn = $('#btn-add-row');
  if (addBtn) addBtn.addEventListener('click', addRow);
}

function markRowStatus(idx, status) {
  if (!pendingPreview?.rows?.[idx]) return;
  // 不覆寫 'added' 標記（新增列維持綠底）
  const cur = pendingPreview.rows[idx]._status;
  if (cur === 'added') return;
  if (status === 'updated' && cur === 'unchanged') {
    pendingPreview.rows[idx]._status = 'updated';
    const tr = document.querySelector(`tr[data-row="${idx}"]`);
    if (tr) { tr.classList.remove('unchanged'); tr.classList.add('updated'); }
  }
}

function onCellEdit(e) {
  if (!pendingPreview) return;
  const el = e.target;
  const idx = parseInt(el.dataset.row, 10);
  const key = el.dataset.key;
  if (isNaN(idx) || !key) return;
  const row = pendingPreview.rows[idx];
  if (!row) return;
  if (key === 'level') {
    row.level = parseInt(el.value, 10) || 0;
    // 重新渲染以更新 lv-X class 與縮排
    renderPreview(pendingPreview, $('#sel-col').value);
    return;
  }
  if (key === 'name') {
    row.name = el.value;
  } else {
    // 數字欄：存「無千分號」字串（與其他資料一致）
    const n = parseNum(el.value);
    row[key] = n == null ? '' : String(n);
  }
  markRowStatus(idx, 'updated');
}

function onCellNumBlur(e) {
  const el = e.target;
  const n = parseNum(el.value);
  el.value = n == null ? '' : fmtNum(n);
}

function deleteRow(idx) {
  if (!pendingPreview?.rows?.[idx]) return;
  pendingPreview.rows.splice(idx, 1);
  // 重新計算 merge 模式的統計
  if (pendingPreview.mode === 'merge') {
    pendingPreview.updated = pendingPreview.rows.filter(r => r._status === 'updated').length;
    pendingPreview.added   = pendingPreview.rows.filter(r => r._status === 'added').length;
  }
  renderPreview(pendingPreview, $('#sel-col').value);
}

function addRow() {
  if (!pendingPreview) return;
  const newRow = isPlanTarget()
    ? { level: 1, name: '', _status: 'added' }
    : { name: '', _status: 'added' };
  pendingPreview.rows.push(newRow);
  if (pendingPreview.mode === 'merge') {
    pendingPreview.added = (pendingPreview.added || 0) + 1;
  }
  renderPreview(pendingPreview, $('#sel-col').value);
  // 將焦點移到新列的名稱欄
  const newTr = document.querySelector(`tr[data-row="${pendingPreview.rows.length - 1}"]`);
  newTr?.querySelector('input[data-key="name"]')?.focus();
}

function refreshPreview() {
  if (!loadedData) { renderPreview(null); return; }
  const tid = targetTableId();
  const rows = (loadedData.tables?.[tid] || []).map(r => ({ ...r, _status: 'unchanged' }));
  pendingPreview = { mode: isPlanTarget() ? 'replace' : 'merge', rows };
  renderPreview(pendingPreview, $('#sel-col').value);
}

function doPreview() {
  if (!loadedData) { setStatus('paste-status', '請先載入 JSON', 'err'); return; }
  const col = $('#sel-col').value;
  const text = $('#ta-paste').value;
  const parsed = parsePaste(text);
  if (!parsed.length) { setStatus('paste-status', '貼上區是空的', 'err'); return; }
  const tid = targetTableId();
  if (isPlanTarget()) {
    pendingPreview = buildPreviewForPlan(parsed, col);
  } else {
    pendingPreview = buildPreviewForControl(parsed, col, loadedData.tables?.[tid] || []);
  }
  renderPreview(pendingPreview, col);
  setStatus('paste-status', `✓ 預覽完成（${parsed.length} 行輸入）`, 'ok');
}

function doApply() {
  if (!loadedData) { setStatus('paste-status', '請先載入 JSON', 'err'); return; }
  if (!pendingPreview) doPreview();
  if (!pendingPreview) return;
  const tid = targetTableId();
  if (!loadedData.tables) loadedData.tables = {};
  // 丟掉 _status 欄
  loadedData.tables[tid] = pendingPreview.rows.map(r => {
    const out = { ...r };
    delete out._status;
    return out;
  });
  setStatus('paste-status', `✓ 已套用到 ${tid}`, 'ok');
}

// ===== 載入 / 下載 =====
async function loadFromProject() {
  try {
    const f = currentFund();
    const url = `sample_fund_${f.code}.json`;
    const r = await fetch(url, { cache: 'no-cache' });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    loadedData = await r.json();
    setStatus('load-status', `✓ 已載入 ${url}`, 'ok');
    refreshPreview();
  } catch (e) {
    setStatus('load-status', `載入失敗：${e.message}`, 'err');
  }
}

function loadFromFile(file) {
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      loadedData = JSON.parse(ev.target.result);
      // 嘗試從檔名偵測基金
      const m = file.name.match(/sample_fund_(\w+)\.json/);
      if (m) {
        const fund = FUNDS.find(x => x.code === m[1]);
        if (fund) $('#sel-fund').value = fund.code;
      }
      setStatus('load-status', `✓ 已載入 ${file.name}`, 'ok');
      refreshPreview();
    } catch (err) {
      setStatus('load-status', `解析失敗：${err.message}`, 'err');
    }
  };
  reader.readAsText(file);
}

// 共用：把 parsed 結果套用到 loadedData 的當前基金 src/use
function _applyParsedToCurrentFund(parsed, fileNameLabel) {
  const fund = currentFund();
  const fundId = fund.id; // plan_agri / plan_forest / ...
  const fundData = parsed.funds[fundId];
  if (!fundData) {
    setStatus('load-status', `${fileNameLabel} 中找不到 ${fund.name} 的資料`, 'err');
    return false;
  }
  if (!loadedData) loadedData = { tables: {} };
  if (!loadedData.tables) loadedData.tables = {};
  loadedData.tables[`${fundId}_src`] = fundData.src;
  loadedData.tables[`${fundId}_use`] = fundData.use;
  if (parsed.reviews?.plan?.org || parsed.reviews?.plan?.gov) {
    if (!loadedData.reviews) loadedData.reviews = {};
    loadedData.reviews.plan = { ...(loadedData.reviews.plan || {}), ...parsed.reviews.plan };
  }
  const total = parsed.counts[fundId] || 0;
  const warn = parsed.warnings?.length ? `（${parsed.warnings.length} 條警告，見 console）` : '';
  if (parsed.warnings?.length) console.warn(`${fileNameLabel} 警告：\n` + parsed.warnings.join('\n'));
  if (parsed.docxMessages?.length) console.info('mammoth messages:', parsed.docxMessages);
  setStatus('load-status', `✓ 已套用 ${fund.short} (src+use 共 ${total} 列) ${warn}`, 'ok');
  pendingPreview = null;
  refreshPreview();
  return true;
}

// 從 word_table_to_json_converter 的 JSON 輸出載入
async function loadFromConverterJSON(file) {
  if (typeof parseWordConverterJSON !== 'function') {
    setStatus('load-status', '找不到 parseWordConverterJSON（word-converter.js 未載入）', 'err');
    return;
  }
  if (!loadedData) {
    try { await loadFromProject(); } catch (e) { /* ignore */ }
  }
  const reader = new FileReader();
  reader.onload = ev => {
    let parsed;
    try {
      const raw = JSON.parse(ev.target.result);
      parsed = parseWordConverterJSON(raw);
    } catch (err) {
      setStatus('load-status', `轉換 JSON 載入失敗：${err.message}`, 'err');
      return;
    }
    _applyParsedToCurrentFund(parsed, 'Word 轉換 JSON');
  };
  reader.readAsText(file);
}

// 直接從 .docx 載入（mammoth.js）
async function loadFromDocxFile(file) {
  if (typeof parseDocxFile !== 'function') {
    setStatus('load-status', '找不到 parseDocxFile（word-converter.js 未載入）', 'err');
    return;
  }
  if (typeof mammoth === 'undefined') {
    setStatus('load-status', 'mammoth.js 未載入，無法解析 .docx', 'err');
    return;
  }
  if (!loadedData) {
    try { await loadFromProject(); } catch (e) { /* ignore */ }
  }
  setStatus('load-status', '⏳ 解析 .docx 中…');
  try {
    const parsed = await parseDocxFile(file);
    _applyParsedToCurrentFund(parsed, `.docx（${file.name}）`);
  } catch (err) {
    setStatus('load-status', `.docx 解析失敗：${err.message}`, 'err');
  }
}

function doDownload() {
  if (!loadedData) { setStatus('dl-status', '尚未載入資料', 'err'); return; }
  const f = currentFund();
  const fname = `sample_fund_${f.code}.json`;
  const blob = new Blob([JSON.stringify(loadedData, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fname;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  setStatus('dl-status', `✓ 已下載 ${fname}`, 'ok');
}

// ===== 初始化 =====
document.addEventListener('DOMContentLoaded', () => {
  // Fund selector
  const fundSel = $('#sel-fund');
  FUNDS.forEach(f => {
    const opt = document.createElement('option');
    opt.value = f.code;
    opt.textContent = `${f.short}（${f.name}）`;
    fundSel.appendChild(opt);
  });

  $('#sel-fund').addEventListener('change', refreshPreview);
  $('#sel-target').addEventListener('change', refreshPreview);
  $('#sel-col').addEventListener('change', () => {
    if (pendingPreview) renderPreview(pendingPreview, $('#sel-col').value);
  });

  $('#btn-load').addEventListener('click', loadFromProject);
  $('#btn-load-file').addEventListener('click', () => $('#file-input').click());
  $('#file-input').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    loadFromFile(f);
    e.target.value = '';
  });
  $('#btn-load-converter').addEventListener('click', () => $('#file-input-converter').click());
  $('#file-input-converter').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    loadFromConverterJSON(f);
    e.target.value = '';
  });
  $('#btn-load-docx').addEventListener('click', () => $('#file-input-docx').click());
  $('#file-input-docx').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    loadFromDocxFile(f);
    e.target.value = '';
  });

  $('#btn-preview').addEventListener('click', doPreview);
  $('#btn-apply').addEventListener('click', doApply);
  $('#btn-download').addEventListener('click', doDownload);

  // Auto-load on first visit
  loadFromProject();
});
