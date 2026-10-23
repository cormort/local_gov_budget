// 自訂表單彙整：讀入多份 form-builder 匯出的 JSON 或已填好的 Excel（同一表單），產生繳交清單／依項目彙總／明細，可匯出 Excel
// 公式欄在彙總時以加總後的數值重算（比率類欄位才正確），求值沿用 form-builder.js 的 fb_evalFormula

let ag_tpl = null;      // 第一份檔案的 template，之後的檔案必須相同
let ag_subs = [];       // [{ file, org, year, user, rows }]；rows 保留原值，列上 _off = true 表示重複填報時不採計
let ag_oneOwner = null; // 是否檢查重複填報：null = 依資料自動判斷（fb_itemOwners），true / false = 使用者指定

const ag_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const ag_sig = t => JSON.stringify([t.title, t.columns.map(c => [c.key, c.label, c.type, c.formula || ''])]);
const ag_num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const ag_fmt = v => typeof v === 'number' ? (+v.toFixed(4)).toLocaleString() : v;

// ========== 1. 讀檔 ==========
async function ag_addFiles(files) {
    const notes = [];
    for (const file of files) {
        try {
            // Excel：與已匯入的表單只比對欄名（和填表頁匯入已填 Excel 同一規則）
            if (/\.xlsx?$/i.test(file.name)) {
                const buf = await file.arrayBuffer();
                const { template, rows, metadata: m, sheet, sheets } = ag_tpl ? fb_readMatching(buf, file.name, ag_tpl.columns.map(c => c.label), ag_tpl.sheet) : fb_readWorkbook(buf, file.name);
                if (!ag_tpl) ag_tpl = { ...template, title: '自訂表單', fromXlsx: true }; // 檔名是機關名，不能當表單名稱
                const bad = fb_headerMismatch(ag_tpl.columns.map(c => c.label), template.columns.map(c => c.label));
                if (bad) { notes.push(`${file.name}：${bad.split('\n')[0]}，略過`); continue; }
                // 填表頁匯出的 Excel 帶「填報資訊」工作表；沒有的話以檔名當機關
                const sub = { file: sheets.length > 1 ? `${file.name}［${sheet}］` : file.name, org: m.org || file.name.replace(/\.[^.]+$/, ''), year: m.year || '', user: m.user || '', rows };
                const dup = ag_subs.findIndex(s => s.org === sub.org && s.year === sub.year);
                if (dup >= 0) { ag_subs[dup] = sub; notes.push(`${sub.org} ${sub.year} 年度重複，以 ${file.name} 取代`); }
                else ag_subs.push(sub);
                continue;
            }
            const data = JSON.parse(await file.text());
            if (!data.template || !Array.isArray(data.template.columns)) { notes.push(`${file.name}：不是自訂表單 JSON`); continue; }
            if (!ag_tpl) ag_tpl = data.template;
            else if (ag_tpl.fromXlsx ? fb_headerMismatch(ag_tpl.columns.map(c => c.label), data.template.columns.map(c => c.label)) : ag_sig(data.template) !== ag_sig(ag_tpl)) { notes.push(`${file.name}：表單「${data.template.title}」與「${ag_tpl.title}」不同，略過`); continue; }
            const m = data.metadata || {};
            const sub = { file: file.name, org: m.org || '(未填機關)', year: m.year || '', user: m.user || '', rows: data.rows || [] };
            const dup = ag_subs.findIndex(s => s.org === sub.org && s.year === sub.year);
            if (dup >= 0) { ag_subs[dup] = sub; notes.push(`${sub.org} ${sub.year} 年度重複，以 ${file.name} 取代`); }
            else ag_subs.push(sub);
        } catch (err) { notes.push(`${file.name}：${err.message}`); }
    }
    ag_defaultOff();
    ag_render();
    ag_msg(notes);
}

// 重複填報的預設：數字完全相同 → 只採計第一份；數字不同 → 全部先採計，等使用者勾選（已勾過的列不動）
function ag_defaultOff() {
    if (!ag_tpl) return;
    const cols = ag_tpl.columns;
    fb_itemOwners(ag_subs, cols).owners.forEach(list => {
        const rows = list.map(o => ag_subs[o.s].rows[o.i]);
        const same = rows.every(r => fb_sameNums(r, rows[0], cols));
        rows.forEach((r, k) => { if (r._off === undefined && list.length > 1) r._off = same && k > 0; });
    });
}

const ag_checkDup = () => ag_oneOwner ?? fb_itemOwners(ag_subs, ag_tpl.columns).oneOwner;

// 實際採計的列：不採計的列數值歸零，再重算上層小計（excel 的上層快取值、JSON 的上層值都不直接相信）
function ag_effRows(s) {
    const cols = ag_tpl.columns, on = ag_checkDup();
    return fb_rollup(s.rows.map(r => {
        if (!(on && r._off)) return { ...r };
        const z = { ...r };
        cols.forEach(c => { if (c.type === 'num' && !c.formula) z[c.key] = ''; });
        return z;
    }), cols);
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
    ag_subs.forEach(s => { const rows = ag_effRows(s), parent = fb_parents(rows.map(r => r[nameKey])); rows.forEach((r, i) => {
        const name = String(r[nameKey] ?? '').trim();
        if (!name) return;
        if (!groups.has(name)) groups.set(name, { [nameKey]: name, _n: new Set() });
        // 合計只加最上層的列：有編號的下層項目已含在上層的小計裡
        [groups.get(name), ...(parent[i] < 0 ? [totalRow] : [])].forEach(g => {
            if (fb_hasData(r, cols)) g._n.add(s.org + '|' + s.year); // 空白列不算有填報
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
    // 狀態：沒有任何一份填數字 → 未填報；清單型的表同一項目仍有多份採計 → 重複
    const on = ag_checkDup(), owners = fb_itemOwners(ag_subs, cols).owners;
    rows.forEach(g => {
        g._status = g === totalRow ? '' : !g._count ? '❌ 未填報' : on && owners.get(g[nameKey])?.length > 1 && g._count > 1 ? '⚠️ 重複' : '✅';
    });
    return rows;
}

// ========== 3. 畫面 ==========
function ag_table(headId, bodyId, headers, rows) {
    document.getElementById(headId).innerHTML = fb_headGrid(headers).rows.map(row =>
        '<tr>' + row.map(h => `<th colspan="${h.colspan}" rowspan="${h.rowspan}">${ag_esc(h.text)}</th>`).join('') + '</tr>').join('');
    document.getElementById(bodyId).innerHTML = rows.map(r =>
        '<tr>' + r.map(v => `<td class="${typeof v === 'number' ? 'text-right' : 'text-left'} px-2${typeof v === 'number' && v < 0 ? ' negative-value' : ''}">${ag_fmt(v)}</td>`).join('') + '</tr>').join('');
}

function ag_views() {
    const cols = ag_tpl.columns;
    const cell = (c, v) => c.type === 'num' ? (v === '' || v === undefined ? '' : ag_num(v)) : String(v ?? '');
    const detail = [];
    let blank = 0;
    const filled = ag_subs.map(s => {
        let n = 0;
        const rows = ag_effRows(s), parent = fb_parents(rows.map(r => r[cols[0].key]));
        rows.forEach((r, i) => {
            if (!String(r[cols[0].key] ?? '').trim()) return;
            if (!fb_hasData(r, cols)) { blank++; return; } // 明細只列有數字的列
            if (!parent.includes(i)) n++; // 「有數字的列」只算實際填的列，上層小計不算
            // 明細的公式欄以該列數值重算，不信任檔案內的值
            const calc = { ...r };
            cols.forEach(c => { if (c.formula) calc[c.key] = fb_evalFormula(c.formula, k => ag_num(calc[k])); });
            detail.push([s.org, s.year, ...cols.map(c => cell(c, calc[c.key]))]);
        });
        return n;
    });
    return {
        subs: { headers: ['機關', '年度', '填表人', '項目數', '有數字的列', '檔名'], rows: ag_subs.map((s, i) => [s.org, s.year, s.user, s.rows.length, filled[i], s.file]) },
        sum: { headers: [...cols, '填報機關數', '狀態'], rows: ag_summary().map(g => [...cols.map(c => c.type === 'num' ? g[c.key] : (c === cols[0] ? g[c.key] : '')), g._count, g._status]) },
        detail: { headers: ['機關', '年度', ...cols], rows: detail },
        blank,
    };
}

// 重複填報審核：清單型的表，同一項目有多份填數字 → 逐列勾選是否採計
function ag_renderDups() {
    const cols = ag_tpl.columns, on = ag_checkDup(), auto = fb_itemOwners(ag_subs, cols);
    const nums = cols.filter(c => c.type === 'num' && !c.formula);
    document.getElementById('ag-dup-mode').checked = on;
    document.getElementById('ag-dup-auto').textContent = ag_oneOwner === null ? `（自動判斷：${auto.oneOwner ? '各項目多半只有一份在填，視為清單型' : '各份填同一批項目，不檢查'}）` : '（手動設定）';
    const groups = on ? [...auto.owners].filter(([, l]) => l.length > 1) : [];
    document.getElementById('ag-dup-card').classList.toggle('hidden', !groups.length);
    if (!groups.length) return;
    document.getElementById('ag-dup-title').textContent = `⚠️ 重複填報（${groups.length} 項）`;
    document.getElementById('ag-dup-head').innerHTML = '<tr>' + ['採計', '項目', '機關', ...nums.map(c => c.label), '備註'].map(h => `<th>${ag_esc(h)}</th>`).join('') + '</tr>';
    document.getElementById('ag-dup').innerHTML = groups.map(([name, list]) => {
        const rows = list.map(o => ({ o, r: ag_subs[o.s].rows[o.i] }));
        const same = rows.every(x => fb_sameNums(x.r, rows[0].r, cols));
        return rows.map(({ o, r }, k) => `<tr class="${r._off ? 'opacity-50' : ''}${k === 0 ? ' border-t-2 border-slate-400' : ''}">
            <td class="text-center"><input type="checkbox" class="ag-dup-use" data-s="${o.s}" data-i="${o.i}"${r._off ? '' : ' checked'}></td>
            <td class="text-left px-2 font-bold">${k === 0 ? ag_esc(name) : ''}</td>
            <td class="text-left px-2" title="${ag_esc(ag_subs[o.s].file)}">${ag_esc(ag_subs[o.s].org)}</td>
            ${nums.map(c => `<td class="text-right px-2">${ag_fmt(ag_num(r[c.key]))}</td>`).join('')}
            <td class="text-left px-2 text-xs">${k === 0 ? (same ? '數字相同，預設只採計第一份' : '數字不同，請勾選要採計的') : ''}</td></tr>`).join('');
    }).join('');
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
    document.getElementById('ag-detail-title').textContent = `明細（${v.detail.rows.length} 列有數字` + (v.blank ? `，略過 ${v.blank} 列空白）` : '）');
    ag_renderDups();
}

// ========== 4. 匯出 Excel ==========
function ag_exportXlsx() {
    if (!ag_subs.length) return alert('請先匯入檔案');
    const v = ag_views(), wb = XLSX.utils.book_new();
    [['依項目彙總', v.sum], ['明細', v.detail], ['繳交清單', v.subs]].forEach(([name, t]) => {
        const head = fb_headGrid(t.headers);
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
    $('ag-btn-clear').onclick = () => { ag_tpl = null; ag_subs = []; ag_oneOwner = null; ag_render(); ag_msg(); };
    $('ag-dup-mode').onchange = e => { ag_oneOwner = e.target.checked; ag_defaultOff(); ag_render(); };
    $('ag-dup').addEventListener('change', e => {
        if (!e.target.classList.contains('ag-dup-use')) return;
        ag_subs[+e.target.dataset.s].rows[+e.target.dataset.i]._off = !e.target.checked;
        ag_render();
    });
    $('ag-subs').addEventListener('click', e => {
        if (!e.target.classList.contains('ag-del')) return;
        ag_subs.splice(+e.target.dataset.i, 1);
        if (!ag_subs.length) ag_tpl = null;
        ag_render(); ag_msg();
    });
});
