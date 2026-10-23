// 自訂表單彙整：讀入多份 form-builder 匯出的 JSON 或已填好的 Excel（同一表單），產生繳交清單／依項目彙總／明細，可匯出 Excel
// 公式欄在彙總時以加總後的數值重算（比率類欄位才正確），求值沿用 form-builder.js 的 fb_evalFormula

let ag_tpl = null;      // 第一份檔案的 template，之後的檔案必須相同
let ag_subs = [];       // [{ file, org, year, user, rows }]

const ag_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const ag_sig = t => JSON.stringify([t.title, t.columns.map(c => [c.key, c.label, c.type, c.formula || ''])]);
const ag_num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const ag_fmt = v => typeof v === 'number' ? (+v.toFixed(4)).toLocaleString() : v;

// ========== 1. 讀檔 ==========
async function ag_addFiles(files) {
    const notes = [];
    for (const file of files) {
        try {
            // Excel：沒有機關欄位，以檔名當機關；與已匯入的表單只比對欄名（和填表頁匯入已填 Excel 同一規則）
            if (/\.xlsx?$/i.test(file.name)) {
                const { template, rows } = fb_readWorkbook(await file.arrayBuffer(), file.name);
                if (!ag_tpl) ag_tpl = { ...template, title: '自訂表單', fromXlsx: true }; // 檔名是機關名，不能當表單名稱
                const bad = fb_headerMismatch(ag_tpl.columns.map(c => c.label), template.columns.map(c => c.label));
                if (bad) { notes.push(`${file.name}：${bad.split('\n')[0]}，略過`); continue; }
                const org = file.name.replace(/\.[^.]+$/, '');
                const dup = ag_subs.findIndex(s => s.org === org && s.year === '');
                const sub = { file: file.name, org, year: '', user: '', rows: fb_rollup(rows, ag_tpl.columns) };
                if (dup >= 0) ag_subs[dup] = sub; else ag_subs.push(sub);
                continue;
            }
            const data = JSON.parse(await file.text());
            if (!data.template || !Array.isArray(data.template.columns)) { notes.push(`${file.name}：不是自訂表單 JSON`); continue; }
            if (!ag_tpl) ag_tpl = data.template;
            else if (ag_tpl.fromXlsx ? fb_headerMismatch(ag_tpl.columns.map(c => c.label), data.template.columns.map(c => c.label)) : ag_sig(data.template) !== ag_sig(ag_tpl)) { notes.push(`${file.name}：表單「${data.template.title}」與「${ag_tpl.title}」不同，略過`); continue; }
            const m = data.metadata || {};
            const sub = { file: file.name, org: m.org || '(未填機關)', year: m.year || '', user: m.user || '', rows: fb_rollup(data.rows || [], ag_tpl.columns) };
            const dup = ag_subs.findIndex(s => s.org === sub.org && s.year === sub.year);
            if (dup >= 0) { ag_subs[dup] = sub; notes.push(`${sub.org} ${sub.year} 年度重複，以 ${file.name} 取代`); }
            else ag_subs.push(sub);
        } catch (err) { notes.push(`${file.name}：${err.message}`); }
    }
    ag_render();
    ag_msg(notes);
}

function ag_msg(notes = []) {
    const base = ag_tpl ? `表單「${ag_tpl.title}」：已匯入 ${ag_subs.length} 份。` : '尚未匯入檔案。';
    document.getElementById('ag-msg').textContent = base + (notes.length ? ' ⚠️ ' + notes.join('；') : '');
}

// ========== 2. 計算 ==========
// 彙總：依第一欄（項目名稱）分組，數值欄加總、公式欄以加總值重算、文字欄留空
function ag_summary() {
    const cols = ag_tpl.columns, nameKey = cols[0].key, groups = new Map();
    const totalRow = { [nameKey]: '合計', _n: new Set() };
    ag_subs.forEach(s => { const parent = fb_parents(s.rows.map(r => r[nameKey])); s.rows.forEach((r, i) => {
        const name = String(r[nameKey] ?? '').trim();
        if (!name) return;
        if (!groups.has(name)) groups.set(name, { [nameKey]: name, _n: new Set() });
        // 合計只加最上層的列：有編號的下層項目已含在上層的小計裡
        [groups.get(name), ...(parent[i] < 0 ? [totalRow] : [])].forEach(g => {
            g._n.add(s.org + '|' + s.year);
            cols.forEach(c => { if (c.type === 'num' && !c.formula) g[c.key] = (g[c.key] || 0) + ag_num(r[c.key]); });
        });
    }); });
    const rows = [...groups.values(), totalRow];
    rows.forEach(g => {
        cols.forEach(c => { if (c.type === 'num' && !c.formula) g[c.key] = g[c.key] || 0; });
        for (let pass = 0; pass < cols.length; pass++) {
            cols.forEach(c => { if (c.formula) g[c.key] = fb_evalFormula(c.formula, k => g[k] || 0); });
        }
        g._count = g._n.size;
    });
    return rows;
}

// ========== 3. 畫面 ==========
// headers 可混用字串與 template 欄位（有 levels 的欄位會排成多層表頭，與填報表單相同）
// 回傳 thead 的列（給 HTML colspan/rowspan）以及等價的二維陣列＋合併範圍（給 Excel）
function ag_headGrid(headers) {
    const rows = fb_theadRows(headers.map(h => typeof h === 'string' ? { label: h } : h));
    const D = rows.length, N = headers.length;
    const aoa = [...Array(D)].map(() => Array(N).fill('')), merges = [];
    const used = [...Array(D)].map(() => Array(N).fill(false));
    rows.forEach((row, k) => {
        let c = 0;
        row.forEach(h => {
            while (used[k][c]) c++;
            aoa[k][c] = h.text;
            for (let r = k; r < k + h.rowspan; r++) for (let cc = c; cc < c + h.colspan; cc++) used[r][cc] = true;
            if (h.rowspan > 1 || h.colspan > 1) merges.push({ s: { r: k, c }, e: { r: k + h.rowspan - 1, c: c + h.colspan - 1 } });
            c += h.colspan;
        });
    });
    return { rows, aoa, merges };
}

function ag_table(headId, bodyId, headers, rows) {
    document.getElementById(headId).innerHTML = ag_headGrid(headers).rows.map(row =>
        '<tr>' + row.map(h => `<th colspan="${h.colspan}" rowspan="${h.rowspan}">${ag_esc(h.text)}</th>`).join('') + '</tr>').join('');
    document.getElementById(bodyId).innerHTML = rows.map(r =>
        '<tr>' + r.map(v => `<td class="${typeof v === 'number' ? 'text-right' : 'text-left'} px-2${typeof v === 'number' && v < 0 ? ' negative-value' : ''}">${ag_fmt(v)}</td>`).join('') + '</tr>').join('');
}

function ag_views() {
    const cols = ag_tpl.columns;
    const cell = (c, v) => c.type === 'num' ? (v === '' || v === undefined ? '' : ag_num(v)) : String(v ?? '');
    const detail = [];
    ag_subs.forEach(s => s.rows.forEach(r => {
        if (!String(r[cols[0].key] ?? '').trim()) return;
        // 明細的公式欄以該列數值重算，不信任檔案內的值
        const calc = { ...r };
        cols.forEach(c => { if (c.formula) calc[c.key] = fb_evalFormula(c.formula, k => ag_num(calc[k])); });
        detail.push([s.org, s.year, ...cols.map(c => cell(c, calc[c.key]))]);
    }));
    return {
        subs: { headers: ['機關', '年度', '填表人', '項目數', '檔名'], rows: ag_subs.map(s => [s.org, s.year, s.user, s.rows.length, s.file]) },
        sum: { headers: [...cols, '填報機關數'], rows: ag_summary().map(g => [...cols.map(c => c.type === 'num' ? g[c.key] : (c === cols[0] ? g[c.key] : '')), g._count]) },
        detail: { headers: ['機關', '年度', ...cols], rows: detail },
    };
}

function ag_render() {
    document.getElementById('ag-result').classList.toggle('hidden', !ag_subs.length);
    if (!ag_subs.length) return;
    const v = ag_views();
    document.getElementById('ag-subs-title').textContent = `繳交清單（${ag_subs.length} 份）`;
    ag_table('ag-subs-head', 'ag-subs', [...v.subs.headers, ''], v.subs.rows.map((r, i) =>
        [...r.map(x => typeof x === 'number' ? x : ag_esc(x)), `<button class="ag-del text-red-500 font-bold" data-i="${i}" title="移除">×</button>`]));
    // 字串先跳脫，數字留給 ag_table 格式化並靠右
    const esc = rows => rows.map(r => r.map(x => typeof x === 'number' ? x : ag_esc(x)));
    ag_table('ag-sum-head', 'ag-sum', v.sum.headers, esc(v.sum.rows));
    ag_table('ag-detail-head', 'ag-detail', v.detail.headers, esc(v.detail.rows));
}

// ========== 4. 匯出 Excel ==========
function ag_exportXlsx() {
    if (!ag_subs.length) return alert('請先匯入檔案');
    const v = ag_views(), wb = XLSX.utils.book_new();
    [['依項目彙總', v.sum], ['明細', v.detail], ['繳交清單', v.subs]].forEach(([name, t]) => {
        const head = ag_headGrid(t.headers);
        const ws = XLSX.utils.aoa_to_sheet([...head.aoa, ...t.rows]);
        ws['!merges'] = head.merges;
        ws['!cols'] = t.headers.map((h, i) => ({ wch: i === 0 ? 20 : 14 }));
        XLSX.utils.book_append_sheet(wb, ws, name);
    });
    XLSX.writeFile(wb, `${ag_tpl.title}_彙整_${ag_subs.length}份.xlsx`);
}

// ========== 5. 綁定事件 ==========
document.addEventListener('DOMContentLoaded', () => {
    const $ = id => document.getElementById(id), drop = $('ag-drop');
    drop.onclick = () => $('ag-file').click();
    $('ag-file').onchange = e => { ag_addFiles([...e.target.files]); e.target.value = ''; };
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('dragover'); ag_addFiles([...e.dataTransfer.files]); });
    $('ag-btn-xlsx').onclick = ag_exportXlsx;
    $('ag-btn-clear').onclick = () => { ag_tpl = null; ag_subs = []; ag_render(); ag_msg(); };
    $('ag-subs').addEventListener('click', e => {
        if (!e.target.classList.contains('ag-del')) return;
        ag_subs.splice(+e.target.dataset.i, 1);
        if (!ag_subs.length) ag_tpl = null;
        ag_render(); ag_msg();
    });
});
