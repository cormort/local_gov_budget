// 貼表：把各機關寄回的 Excel（照空白範本填寫）依標準機關順序一列接一列貼到範本下方，保留原格式，匯出一份彙總表
// 版面：?form=名稱 時用 forms/名稱.xlsx，沒有就用第一個拖入的檔案；取第一個沒隱藏的工作表，表頭到「合計」列（沒有合計列就到表頭最後一列）照抄，「註」以下的說明列移到最後
// 小計：範本「合計」列有 COUNTIF 的欄重算勾選數、有「共…案」的欄重算案數；主管列算到下一個主管，基金列算到下一個標題列
// ponytail: 提案列＝第 2 欄（提案序號）有填、標題列＝第 1 欄有填且第 2 欄空白，是照協商彙總表的版面寫的；別的版面要另訂規則

let ps_base = null;   // { name, buf, fromFile, sheet, placeholders }：fromFile = 版面取自第一個拖入的回傳檔（沒有指定範本）
let ps_files = [];    // [{ file, org, rank, wb, ws, rows, pinned }]；順序就是貼上順序，可手動調整；pinned = 手動選的工作表
let ps_orgs = { sortOrder: [], aliasMap: {} };
let ps_headEnd = 0;   // 範本表頭最後一列（1 起算），沒有「合計」列時資料從下一列開始

const ps_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const ps_text = c => String(c?.text ?? '').trim();
const ps_slave = c => c.isMerged && c.master.address !== c.address; // 合併儲存格的非左上格：ExcelJS 會回傳左上格的值
const ps_own = c => ps_slave(c) ? '' : ps_text(c);
const ps_sheet = wb => wb.worksheets.find(w => w.state === 'visible') || wb.worksheets[0];
const ps_filler = a => /ＯＯ|ＸＸ|○○|^【範例/.test(a); // 空白表留下的佔位列（ＯＯ主管、ＸＸ基金、範例）
const ps_value = v => (v && typeof v === 'object' && 'formula' in v) || (v && typeof v === 'object' && 'sharedFormula' in v) ? v.result ?? null : v; // 公式只留值，參照位置貼過來就不對了

// 一個工作表的版面：合計列、說明列起點、欄數
function ps_layout(ws) {
    let total = 0, notes = 0;
    ws.eachRow((row, r) => {
        const a = ps_text(row.getCell(1));
        if (!total && a === '合計') total = r;
        else if (total && !notes && a.startsWith('註')) notes = r;
    });
    return { total: total || ps_headEnd, hasTotal: !!total, notes: notes || ws.rowCount + 1, cols: ws.columnCount };
}

// 回傳檔要貼的列：合計列以下、說明列以上；空白列與範本原有的佔位列（ＯＯ主管、範例…）略過
function ps_pickRows(ws, placeholders) {
    const { total, notes } = ps_layout(ws);
    const rows = [];
    for (let r = (total || 0) + 1; r < notes; r++) {
        const row = ws.getRow(r);
        const cells = [];
        row.eachCell(c => cells.push(ps_text(c)));
        if (!cells.some(Boolean)) continue;
        const a = ps_text(row.getCell(1));
        if (placeholders.has(a) || ps_filler(a)) continue;
        rows.push(r);
    }
    return rows;
}

function ps_copyRow(src, sr, dst, dr, cols) {
    const s = src.getRow(sr), d = dst.getRow(dr);
    if (s.height) d.height = s.height;
    for (let c = 1; c <= cols; c++) {
        const from = s.getCell(c), to = d.getCell(c);
        to.value = ps_slave(from) ? null : ps_value(from.value);
        to.style = JSON.parse(JSON.stringify(from.style || {}));
    }
    // 同一列的合併儲存格（標題列的 A:D、F:H）照搬
    (src.model.merges || []).forEach(m => {
        const [a, b] = m.split(':');
        const ra = +a.replace(/\D/g, ''), rb = +b.replace(/\D/g, '');
        if (ra === sr && rb === sr) dst.mergeCells(a.replace(/\d+/, dr) + ':' + b.replace(/\d+/, dr));
    });
}

async function ps_loadBase(name) {
    if (/[\\/]|\.\./.test(name)) throw new Error('範本名稱不合法：' + name);
    const res = await fetch('forms/' + encodeURIComponent(name) + '.xlsx', { cache: 'no-cache' });
    if (!res.ok) throw new Error(`找不到空白 Excel「forms/${name}.xlsx」`);
    await ps_setBase(name, await res.arrayBuffer());
}

// 換版面後，已讀入的檔案依新版面重新挑工作表與要貼的列
async function ps_setBase(name, buf, fromFile = false) {
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(buf);
    const bws = ps_sheet(book), bl = ps_layout(bws);
    // 空白範本：合計列以下原有的列（ＯＯ主管、範例、註）都是佔位；版面取自回傳檔時那些是真資料，只認佔位字樣
    const placeholders = new Set();
    if (!fromFile) for (let r = bl.total + 1; r <= bws.rowCount; r++) { const a = ps_text(bws.getRow(r).getCell(1)); if (a) placeholders.add(a); }
    ps_base = { name, buf, fromFile, sheet: bws.name, placeholders };
    ps_files.forEach(f => { f.ws = f.wb.getWorksheet(f.pinned || bws.name) || ps_sheet(f.wb); f.rows = ps_pickRows(f.ws, placeholders); });
    // 表頭範圍沿用網頁表單的判斷（多層合併表頭、略過表名與單位列）
    const wb = XLSX.read(buf, { type: 'array' }), ws = wb.Sheets[wb.SheetNames.find((_, i) => !wb.Workbook?.Sheets?.[i]?.Hidden) || wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const text = (r, c) => { const x = ws[XLSX.utils.encode_cell({ r, c })]; return x ? String(x.w ?? x.v ?? '').trim() : ''; };
    ps_headEnd = (fb_headerBlock(text, ws['!merges'] || [], range.s.r, range.e.r, range.s.c, range.e.c)?.he ?? -1) + 1;
}

async function ps_addFiles(files) {
    const notes = [];
    const first = files.find(f => /\.xlsx$/i.test(f.name));
    if (!ps_base && first) await ps_setBase(first.name.replace(/\.[^.]+$/, ''), await first.arrayBuffer(), true);
    if (!ps_base) { ps_render(['請拖入 .xlsx 檔']); return; }
    const { sheet, placeholders } = ps_base;
    for (const file of files) {
        if (!/\.xlsx$/i.test(file.name)) { notes.push(`${file.name}：不是 .xlsx，略過`); continue; }
        try {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(await file.arrayBuffer());
            let ws = wb.getWorksheet(sheet);
            if (!ws) { ws = ps_sheet(wb); notes.push(`${file.name}：沒有「${sheet}」工作表，改讀「${ws.name}」`); }
            const stem = file.name.replace(/\.[^.]+$/, '');
            const hits = fb_detectOrgs(stem, ps_orgs.sortOrder, ps_orgs.aliasMap);
            if (hits.length !== 1) notes.push(`${file.name}：${hits.length ? '檔名含多個機關（' + hits.join('、') + '）' : '檔名認不出機關'}，排在最後`);
            const org = hits.length === 1 ? hits[0] : stem;
            const entry = { file: file.name, org, rank: hits.length === 1 ? ps_orgs.sortOrder.indexOf(org) : Infinity, wb, ws, rows: ps_pickRows(ws, placeholders) };
            // 同機關重複：原位置換掉；新的依標準順序插入，不打亂已手動調整的順序
            const dup = hits.length === 1 ? ps_files.findIndex(f => f.org === org) : -1;
            if (dup >= 0) ps_files[dup] = entry;
            else { const at = ps_files.findIndex(f => f.rank > entry.rank); ps_files.splice(at < 0 ? ps_files.length : at, 0, entry); }
        } catch (e) { notes.push(`${file.name}：${e.message}`); }
    }
    ps_render(notes);
}

async function ps_build() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(ps_base.buf);
    const ws = ps_sheet(wb), { total, hasTotal, notes, cols } = ps_layout(ws);
    wb.worksheets.filter(w => w !== ws).forEach(w => wb.removeWorksheet(w.id)); // 往年供參的工作表不帶
    const ps_countCols = [];
    for (let c = 1; c <= cols; c++) for (let r = total; r < notes; r++) {
        const v = ws.getRow(r).getCell(c).value;
        if (v && typeof v === 'object' && /COUNTIF/i.test(v.formula || v.sharedFormula || '')) { ps_countCols.push(c); break; }
    }

    // 說明列先存起來（值、格式、列高、合併），清掉合計列以下再貼
    const tmp = new ExcelJS.Workbook(), keep = tmp.addWorksheet('n');
    const footer = [];
    for (let r = notes; r <= ws.rowCount; r++) { ps_copyRow(ws, r, keep, footer.length + 1, cols); footer.push(r); }
    (ws.model.merges || []).forEach(m => { if (+m.split(':')[0].replace(/\D/g, '') > total) ws.unMergeCells(m); });
    ws.spliceRows(total + 1, ws.rowCount - total);

    let r = total + 1;
    const placed = []; // { r, kind: 'head'|'item', level }
    ps_files.forEach(f => f.rows.forEach(sr => {
        ps_copyRow(f.ws, sr, ws, r, cols);
        const a = ps_own(ws.getRow(r).getCell(1)), b = ps_own(ws.getRow(r).getCell(2));
        placed.push({ r, kind: b ? 'item' : a ? 'head' : 'item', top: a.endsWith('主管') });
        r++;
    }));
    const last = r - 1;
    footer.forEach((_, i) => ps_copyRow(keep, i + 1, ws, r + i, cols));

    // 小計：範本合計、小計列有 COUNTIF 的欄、合計列有「共…案」的欄（清掉範本列之前已記下 ps_countCols）
    const countCols = ps_countCols, caseCols = [];
    for (let c = 1; c <= cols; c++) if (/^共.*案$/.test(ps_text(ws.getRow(total).getCell(c)))) caseCols.push(c);
    const fill = (row, from, to) => {
        const items = placed.filter(p => p.r >= from && p.r <= to && p.kind === 'item');
        countCols.forEach(c => {
            const col = ws.getColumn(c).letter;
            const n = items.filter(p => ps_text(ws.getRow(p.r).getCell(c)).toLowerCase() === 'v').length;
            ws.getRow(row).getCell(c).value = from > to ? 0 : { formula: `COUNTIF(${col}${from}:${col}${to},"v")`, result: n };
        });
        caseCols.forEach(c => { ws.getRow(row).getCell(c).value = `共${items.length}案`; });
    };
    if (hasTotal) fill(total, total + 1, last);
    placed.forEach((p, i) => {
        if (p.kind !== 'head') return;
        const next = placed.slice(i + 1).find(q => q.kind === 'head' && (p.top ? q.top : true));
        fill(p.r, p.r + 1, next ? next.r - 1 : last);
    });
    return wb.xlsx.writeBuffer();
}

function ps_render(notes = []) {
    const $ = id => document.getElementById(id);
    const n = ps_files.length;
    const sheets = f => f.wb.worksheets.filter(w => w.state === 'visible' || w === f.ws)
        .map(w => `<option${w === f.ws ? ' selected' : ''}>${ps_esc(w.name)}</option>`).join('');
    $('ps-list').innerHTML = ps_files.map((f, i) => `<tr><td>${i + 1}</td><td class="text-left">${ps_esc(f.org)}</td><td class="text-left">${ps_esc(f.file)}</td>
        <td><select class="ps-sheet text-sm" data-i="${i}" aria-label="${ps_esc(f.org)} 的工作表">${sheets(f)}</select></td><td>${f.rows.length}</td>
        <td class="whitespace-nowrap"><button class="btn btn-sm" data-move="${i}" data-d="-1" aria-label="上移 ${ps_esc(f.org)}"${i ? '' : ' disabled'}>↑</button>
            <button class="btn btn-sm" data-move="${i}" data-d="1" aria-label="下移 ${ps_esc(f.org)}"${i < n - 1 ? '' : ' disabled'}>↓</button>
            <button class="btn btn-sm btn-danger" data-remove="${i}" aria-label="移除 ${ps_esc(f.org)}">移除</button></td></tr>`).join('');
    $('ps-result').classList.toggle('hidden', !ps_files.length);
    $('ps-btn-xlsx').disabled = !ps_files.length;
    $('ps-form').textContent = ps_base ? (ps_base.fromFile ? `依第一個檔案「${ps_base.name}」` : `範本「${ps_base.name}」`) : '依第一個拖入的檔案（或按「指定範本檔」）';
    $('ps-msg').textContent = (ps_files.length ? `已讀入 ${ps_files.length} 份，共 ${ps_files.reduce((n, f) => n + f.rows.length, 0)} 列。` : '尚未匯入檔案。') + (notes.length ? ' ⚠️ ' + notes.join('；') : '');
}

document.addEventListener('DOMContentLoaded', async () => {
    const $ = id => document.getElementById(id);
    const name = new URLSearchParams(location.search).get('form');
    try {
        if (name) await ps_loadBase(name);
        ps_orgs = await fetch(FB_ORGS, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
        ps_orgs.sortOrder ||= []; ps_orgs.aliasMap ||= {};
        ps_render();
    } catch (e) { $('ps-msg').textContent = e.message; return; }
    const drop = $('ps-drop'), input = $('ps-file');
    drop.onclick = () => input.click();
    input.onchange = () => { ps_addFiles([...input.files]); input.value = ''; };
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('dragover'); ps_addFiles([...e.dataTransfer.files]); });
    const tpl = $('ps-tpl-file');
    $('ps-btn-tpl').onclick = () => tpl.click();
    tpl.onchange = async () => {
        const f = tpl.files[0]; tpl.value = '';
        if (!f) return;
        try { await ps_setBase(f.name.replace(/\.[^.]+$/, ''), await f.arrayBuffer()); ps_render(); } catch (e) { ps_render([`${f.name}：${e.message}`]); }
    };
    $('ps-list').addEventListener('change', e => {
        if (!e.target.classList.contains('ps-sheet')) return;
        const f = ps_files[+e.target.dataset.i];
        f.pinned = e.target.value;
        f.ws = f.wb.getWorksheet(f.pinned);
        f.rows = ps_pickRows(f.ws, ps_base.placeholders);
        ps_render();
    });
    $('ps-list').addEventListener('click', e => {
        const b = e.target.closest('button');
        if (!b) return;
        if (b.dataset.remove) ps_files.splice(+b.dataset.remove, 1);
        else { const i = +b.dataset.move, j = i + +b.dataset.d; [ps_files[i], ps_files[j]] = [ps_files[j], ps_files[i]]; }
        ps_render();
    });
    $('ps-btn-clear').onclick = () => { ps_files = []; if (ps_base?.fromFile) ps_base = null; ps_render(); };
    $('ps-btn-xlsx').onclick = async () => {
        const buf = await ps_build();
        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        a.download = `${ps_base.name}-貼表.xlsx`;
        a.click();
        URL.revokeObjectURL(a.href);
    };
});
