// 自訂表單查填：由 Excel 範本產生表頭與表單
// template = { title, columns: [{ key: 'B', label, type: 'num'|'text', formula?: 'B-C' }] }
// 第一欄固定為項目名稱；formula 為同一列欄位代號的四則運算（CSP 禁 eval，故自寫小解析器）

// ========== 1. 公式：Excel → 欄位運算式 → 求值 ==========
// "SUM(B2:D2)-$E$2" → "(B+C+D)-E"；跨列參照或不支援的函數回傳 null
function fb_translateFormula(f, row) {
    const XLSX_ = typeof XLSX !== 'undefined' ? XLSX : null;
    const colIdx = s => XLSX_ ? XLSX_.utils.decode_col(s) : [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const colStr = i => XLSX_ ? XLSX_.utils.encode_col(i) : (i >= 26 ? colStr(Math.floor(i / 26) - 1) : '') + String.fromCharCode(65 + i % 26);
    let ok = true;
    let expr = f.replace(/\s+/g, '').toUpperCase()
        .replace(/SUM\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)/g, (m, c1, r1, c2, r2) => {
            if (+r1 !== row || +r2 !== row) { ok = false; return m; }
            const cols = [];
            for (let i = colIdx(c1); i <= colIdx(c2); i++) cols.push(colStr(i));
            return '(' + cols.join('+') + ')';
        })
        .replace(/\$?([A-Z]+)\$?(\d+)/g, (m, c, r) => { if (+r !== row) ok = false; return c; });
    if (!ok || !/^[A-Z0-9.+\-*/()]+$/.test(expr)) return null;
    try { fb_evalFormula(expr, () => 1); } catch (e) { return null; }
    return expr;
}

// 遞迴下降：expr := term (('+'|'-') term)* ; term := factor (('*'|'/') factor)* ; factor := -factor | (expr) | number | 欄位
function fb_evalFormula(expr, get) {
    const tokens = expr.match(/[A-Z]+|\d+(?:\.\d+)?|[-+*/()]/g) || [];
    if (tokens.join('') !== expr) throw new Error('bad formula: ' + expr);
    let i = 0;
    const peek = () => tokens[i], next = () => tokens[i++];
    const parseExpr = () => {
        let v = parseTerm();
        while (peek() === '+' || peek() === '-') v = next() === '+' ? v + parseTerm() : v - parseTerm();
        return v;
    };
    const parseTerm = () => {
        let v = parseFactor();
        while (peek() === '*' || peek() === '/') {
            if (next() === '*') v *= parseFactor();
            else { const d = parseFactor(); v = d === 0 ? 0 : v / d; } // ponytail: 除以 0 顯示 0，不顯示 #DIV/0!
        }
        return v;
    };
    const parseFactor = () => {
        const t = next();
        if (t === '-') return -parseFactor();
        if (t === '(') { const v = parseExpr(); if (next() !== ')') throw new Error('missing )'); return v; }
        if (t === undefined) throw new Error('unexpected end');
        if (/^\d/.test(t)) return parseFloat(t);
        if (/^[A-Z]+$/.test(t)) return get(t);
        throw new Error('unexpected ' + t);
    };
    const v = parseExpr();
    if (i !== tokens.length) throw new Error('trailing tokens');
    return v;
}

if (typeof module !== 'undefined') module.exports = { fb_translateFormula, fb_evalFormula };
if (typeof document === 'undefined') { /* node 測試用 */ } else {

// ========== 2. 讀 Excel 範本 ==========
let fb_template = null;

function fb_readWorkbook(buf, fileName) {
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true, sheetStubs: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const cell = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
    const text = (r, c) => { const x = cell(r, c); return x ? String(x.w ?? x.v ?? '').trim() : ''; };

    // 表頭 = 第一個有內容的列
    let hr = range.s.r;
    while (hr <= range.e.r && ![...Array(range.e.c - range.s.c + 1)].some((_, k) => text(hr, range.s.c + k))) hr++;
    if (hr > range.e.r) throw new Error('找不到表頭列');

    const columns = [], warnings = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const label = text(hr, c);
        if (!label) continue;
        const key = XLSX.utils.encode_col(c);
        const col = { key, label, type: columns.length === 0 ? 'text' : 'num' };
        if (columns.length > 0) {
            // 取表頭下方第一個公式當該欄公式
            for (let r = hr + 1; r <= range.e.r; r++) {
                const x = cell(r, c);
                if (x && x.f) {
                    const expr = fb_translateFormula(x.f, r + 1);
                    if (expr) col.formula = expr; else warnings.push(`「${label}」公式 =${x.f} 不支援，改為手填`);
                    break;
                }
                if (x && x.t === 's' && x.v.trim()) col.type = 'text';
            }
        }
        columns.push(col);
    }
    // 公式只能參照範本內的欄
    const keys = new Set(columns.map(c => c.key));
    columns.forEach(c => {
        if (c.formula && (c.formula.match(/[A-Z]+/g) || []).some(k => !keys.has(k))) {
            warnings.push(`「${c.label}」公式參照了沒有表頭的欄，改為手填`);
            delete c.formula;
        }
    });

    // 資料列 → 預設項目（第一欄有值才算）
    const rows = [];
    for (let r = hr + 1; r <= range.e.r; r++) {
        if (!text(r, XLSX.utils.decode_col(columns[0].key))) continue;
        const row = {};
        columns.forEach(c => {
            if (c.formula) return;
            const x = cell(r, XLSX.utils.decode_col(c.key));
            row[c.key] = x ? (c.type === 'num' && typeof x.v === 'number' ? x.v : String(x.v ?? '')) : '';
        });
        rows.push(row);
    }

    return { template: { title: fileName.replace(/\.[^.]+$/, ''), sheet: wb.SheetNames[0], columns }, rows, warnings };
}

// ========== 3. 畫表 + 計算 ==========
const fb_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

function fb_render(template, rows) {
    fb_template = template;
    document.getElementById('fb-publish').classList.add('hidden');
    document.getElementById('fb-title').value = template.title;
    document.getElementById('fb-heading').textContent = template.title;
    document.getElementById('fb-thead').innerHTML =
        '<tr><th class="w-8"></th>' + template.columns.map(c => `<th>${fb_esc(c.label)}${c.formula ? ' 🔒' : ''}</th>`).join('') + '</tr>';
    document.getElementById('fb-tfoot').innerHTML =
        '<tr><td>∑</td>' + template.columns.map((c, i) =>
            i === 0 ? '<td class="text-center">合計</td>' : `<td>${c.type === 'num' ? `<input type="text" class="t-${c.key}" readonly>` : ''}</td>`).join('') + '</tr>';
    const tbody = document.getElementById('fb-tbody');
    tbody.innerHTML = '';
    (rows.length ? rows : [{}]).forEach(fb_addRow);
    document.getElementById('fb-card').classList.remove('hidden');
    fb_calc();
}

function fb_addRow(data = {}) {
    const tr = document.createElement('tr');
    tr.innerHTML = '<td class="text-center"><button class="fb-del text-red-500 font-bold" title="刪除此列">×</button></td>' +
        fb_template.columns.map((c, i) => {
            const type = c.type === 'num' ? 'number' : 'text';
            const cls = `v-${c.key}${i === 0 || c.type === 'text' ? ' text-left' : ''}`;
            return `<td><input type="${type}" class="${cls}" value="${fb_esc(data[c.key])}"${c.formula ? ' readonly tabindex="-1"' : ''}></td>`;
        }).join('');
    document.getElementById('fb-tbody').appendChild(tr);
}

function fb_calc() {
    if (!fb_template) return;
    const cols = fb_template.columns, totals = {};
    document.querySelectorAll('#fb-tbody tr').forEach(tr => {
        const get = k => parseFloat(tr.querySelector('.v-' + k)?.value) || 0;
        // 依欄序重算數次，讓公式欄可以參照其他公式欄
        for (let pass = 0; pass < cols.length; pass++) {
            cols.forEach(c => { if (c.formula) tr.querySelector('.v-' + c.key).value = +fb_evalFormula(c.formula, get).toFixed(4); });
        }
        cols.forEach(c => {
            if (c.type !== 'num') return;
            const v = get(c.key);
            tr.querySelector('.v-' + c.key).classList.toggle('negative-value', v < 0);
            totals[c.key] = (totals[c.key] || 0) + v; // ponytail: 合計一律直接加總，比率類公式欄的合計無意義
        });
    });
    cols.forEach(c => {
        const el = document.querySelector('#fb-tfoot .t-' + c.key);
        if (el) { el.value = (totals[c.key] || 0).toLocaleString(); el.classList.toggle('negative-value', (totals[c.key] || 0) < 0); }
    });
}

// ========== 4. JSON 匯出 / 匯入 ==========
function fb_collectRows() {
    return [...document.querySelectorAll('#fb-tbody tr')].map(tr => {
        const row = {};
        fb_template.columns.forEach(c => {
            const v = tr.querySelector('.v-' + c.key).value;
            row[c.key] = c.type === 'num' && v !== '' ? Number(v) : v;
        });
        return row;
    }).filter(r => String(r[fb_template.columns[0].key] ?? '').trim());
}

const fb_saveJSON = (data, fileName) =>
    saveAs(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), fileName);

function fb_exportJSON() {
    if (!fb_template) return alert('請先上傳 Excel 範本或匯入 JSON');
    const val = id => document.getElementById(id).value;
    const data = { template: fb_template, metadata: { org: val('fb-org'), year: val('fb-year'), user: val('fb-user') }, rows: fb_collectRows() };
    fb_saveJSON(data, `${fb_template.title}_${val('fb-org') || '機關'}_${val('fb-year')}年度.json`);
}

// 固定表單範本：放到 repo 的 forms/ 資料夾，以 form-builder.html?form=檔名 開啟
// 目前表格的列（不含公式欄）會成為預設項目；不含機關／填表人
function fb_exportTemplate() {
    if (!fb_template) return alert('請先上傳 Excel 範本或匯入 JSON');
    const rows = fb_collectRows().map(r => {
        fb_template.columns.forEach(c => { if (c.formula) delete r[c.key]; });
        return r;
    });
    fb_saveJSON({ template: fb_template, rows }, `${fb_template.title}.json`);
    fb_msg('');
    document.getElementById('fb-publish-file').textContent = document.getElementById('fb-publish-file2').textContent = `${fb_template.title}.json`;
    document.getElementById('fb-publish-github').href = FB_UPLOAD_URL;
    document.getElementById('fb-share-url').value = FB_SITE_URL + 'form-builder.html?form=' + encodeURIComponent(fb_template.title);
    document.getElementById('fb-publish').classList.remove('hidden');
}

// ponytail: repo 寫死，fork 到別的帳號時改這兩行
const FB_UPLOAD_URL = 'https://github.com/cormort/local_gov_budget/upload/main/forms';
const FB_SITE_URL = 'https://cormort.github.io/local_gov_budget/';

function fb_load(data) {
    if (!data.template || !Array.isArray(data.template.columns)) throw new Error('不是自訂表單的 JSON（缺 template）');
    const m = data.metadata || {};
    document.getElementById('fb-org').value = m.org || '';
    if (m.year) document.getElementById('fb-year').value = m.year;
    document.getElementById('fb-user').value = m.user || '';
    fb_render(data.template, data.rows || []);
}

function fb_importJSON(text) {
    const data = JSON.parse(text);
    fb_load(data);
    fb_msg(`已匯入「${data.template.title}」，${(data.rows || []).length} 列。`);
}

// ?form=名稱 → 讀 forms/名稱.json
async function fb_loadFixedForm(name) {
    if (/[\\/]|\.\./.test(name)) throw new Error('表單名稱不合法：' + name);
    const res = await fetch(`forms/${encodeURIComponent(name)}.json`, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`找不到固定表單「${name}」（forms/${name}.json，HTTP ${res.status}）`);
    fb_load(await res.json());
    document.getElementById('fb-btn-xlsx').classList.add('hidden');
    document.getElementById('fb-btn-template').classList.add('hidden');
    document.getElementById('fb-link-aggregate').classList.add('hidden');
    fb_msg(`固定表單「${name}」：請填寫機關、填表人與各列數值後按「💾 匯出 JSON」回傳。`);
}

function fb_msg(text, warnings = []) {
    document.getElementById('fb-msg').textContent = text + (warnings.length ? ' ⚠️ ' + warnings.join('；') : '');
}

// ========== 5. 綁定事件 ==========
document.addEventListener('DOMContentLoaded', () => {
    if (!document.getElementById('fb-tbody')) return; // 彙整頁只借用公式函數
    const $ = id => document.getElementById(id);
    $('fb-btn-xlsx').onclick = () => $('fb-xlsx-file').click();
    $('fb-btn-import').onclick = () => $('fb-json-file').click();
    $('fb-btn-export').onclick = fb_exportJSON;
    $('fb-btn-template').onclick = fb_exportTemplate;
    $('fb-btn-copy').onclick = async () => {
        const url = $('fb-share-url');
        try { await navigator.clipboard.writeText(url.value); } catch (e) { url.select(); document.execCommand('copy'); }
        $('fb-btn-copy').textContent = '已複製 ✓';
        setTimeout(() => { $('fb-btn-copy').textContent = '複製連結'; }, 1500);
    };
    $('fb-btn-add').onclick = () => { fb_addRow(); fb_calc(); };

    $('fb-xlsx-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try {
            const { template, rows, warnings } = fb_readWorkbook(await file.arrayBuffer(), file.name);
            fb_render(template, rows);
            const nf = template.columns.filter(c => c.formula).length;
            fb_msg(`已由「${file.name}」產生表單：${template.columns.length} 欄（${nf} 個公式欄）、${rows.length} 個預設項目。`, warnings);
        } catch (err) { alert('Excel 解析失敗：' + err.message); }
    };
    $('fb-json-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try { fb_importJSON(await file.text()); } catch (err) { alert('JSON 匯入失敗：' + err.message); }
    };

    $('fb-tbody').addEventListener('input', fb_calc);
    $('fb-tbody').addEventListener('click', e => {
        if (e.target.classList.contains('fb-del')) { e.target.closest('tr').remove(); fb_calc(); }
    });

    const form = new URLSearchParams(location.search).get('form');
    if (form) fb_loadFixedForm(form).catch(err => fb_msg('⚠️ ' + err.message));
});

}
