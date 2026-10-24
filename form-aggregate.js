// 自訂表單彙整：讀入多份 form-builder 匯出的 JSON 或已填好的 Excel（同一表單），產生繳交清單／依項目彙總／明細，可匯出 Excel
// 表單可有多個工作表（每個是一個分頁），彙總、明細、重複填報都依工作表分開算
// 公式欄在彙總時以加總後的數值重算（比率類欄位才正確），求值沿用 form-builder.js 的 fb_evalFormula

let ag_form = null;     // 表單 { title, sheets: [{ name, columns, caption }] }：來自範本 JSON 或第一份回傳檔
let ag_subs = [];       // [{ file, org, year, user, tables: { 工作表名: rows } }]；rows 保留原值，列上 _off = true 表示重複填報時不採計
let ag_oneOwner = {};   // 各工作表是否檢查重複填報：沒設定 = 依資料自動判斷（fb_itemOwners），true / false = 使用者指定
let ag_cur = 0;         // 目前顯示的工作表
let ag_fixed = false;   // 由 ?form= 載入的範本：清空回傳檔時保留表單
let ag_survey = null;   // 由 ?survey= 開啟的調查專案 { id, title, form, respondents, … }

// 回傳檔帶的調查代號不是這次的調查 → 不收（沒帶代號的自製 Excel、舊 JSON 照收）
const ag_otherSurvey = m => ag_survey && m.survey && m.survey !== ag_survey.id;

const ag_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const ag_num = fb_num; // 文字格式的 "1,234" 用 parseFloat 會變成 1
const ag_fmt = v => typeof v === 'number' ? fb_fmt(v) : v;
const ag_sheet = k => ag_form.sheets[k];
const ag_rowsOf = (s, k) => s.tables[ag_sheet(k).name] || [];
const ag_view = k => ag_subs.map(s => ({ rows: ag_rowsOf(s, k) })); // fb_itemOwners 要的形狀

// 同一份表單：工作表名稱與各表欄名都相同（Excel 回傳的只能比欄名，所以一律只比欄名）
const ag_formMismatch = f => {
    const a = ag_form.sheets.map(s => s.name).join('、'), b = f.sheets.map(s => s.name).join('、');
    if (a !== b) return `工作表「${b}」與表單「${a}」不同`;
    for (const [k, s] of f.sheets.entries()) {
        const bad = fb_headerMismatch(ag_sheet(k).columns.map(c => c.label), s.columns.map(c => c.label));
        if (bad) return `［${s.name}］${bad.split('\n')[0]}`;
    }
    return null;
};

// ========== 1. 讀檔 ==========
// 第一份 Excel 決定表單：填表頁匯出的查填檔（有「填報資訊」）→ 其餘工作表都是表單；其他 Excel 只取第一個工作表
// ponytail: 多工作表的表單若收到的都是自製 Excel，請先拖入「匯出範本」的 .json
function ag_formFromXlsx(buf, fileName) {
    const wb = fb_openWorkbook(buf);
    const names = wb.Sheets[FB_INFO_SHEET] ? fb_visibleSheets(wb).filter(n => n !== FB_INFO_SHEET) : [fb_visibleSheets(wb)[0] || wb.SheetNames[0]];
    const list = names.flatMap(n => { try { return [fb_readWorkbook(wb, fileName, n)]; } catch (e) { return []; } });
    if (!list.length) throw new Error('找不到表頭列');
    return { ...fb_formFromParsed('自訂表單', list), fromXlsx: true }; // 檔名是機關名，不能當表單名稱
}

function ag_addSub(sub, notes) {
    const dup = ag_subs.findIndex(s => s.org === sub.org && s.year === sub.year);
    if (dup >= 0) { ag_subs[dup] = sub; notes.push(`${sub.org} ${sub.year} 年度重複，以 ${sub.file} 取代`); }
    else ag_subs.push(sub);
}

async function ag_addFiles(files) {
    const notes = [];
    for (const file of files) {
        try {
            if (/\.xlsx?$/i.test(file.name)) {
                const buf = await file.arrayBuffer();
                if (!ag_form) ag_form = ag_formFromXlsx(buf, file.name);
                const { found, notes: miss, metadata: m } = fb_matchSheets(buf, file.name, ag_form.sheets);
                const hit = Object.keys(found);
                if (!hit.length) { notes.push(`${file.name}：${miss.join('；')}，略過`); continue; }
                if (ag_otherSurvey(m)) { notes.push(`${file.name}：是調查「${m.survey}」的回傳檔，不是這次調查，略過`); continue; }
                if (miss.length) notes.push(`${file.name}：${miss.join('；')}`);
                const bad = hit.flatMap(n => found[n].warnings.filter(w => w.includes('不是數字')));
                if (bad.length) notes.push(`${file.name}：${bad.join('；')}`);
                // 工作表名稱與範本不同時標出實際讀的是哪一張
                const renamed = hit.filter(n => found[n].sheet !== n).map(n => `${n}←${found[n].sheet}`);
                ag_addSub({
                    file: file.name + (renamed.length ? `［${renamed.join('、')}］` : ''),
                    org: m.org || file.name.replace(/\.[^.]+$/, ''), year: m.year || '', user: m.user || '', // 沒有「填報資訊」時以檔名當機關
                    tables: Object.fromEntries(hit.map(n => [n, found[n].rows])),
                }, notes);
                continue;
            }
            const data = JSON.parse(await file.text());
            const form = fb_normalize(data);
            if (!ag_form) ag_form = { title: form.title, sheets: form.sheets };
            else { const bad = ag_formMismatch(form); if (bad) { notes.push(`${file.name}：${bad}，略過`); continue; } }
            if (!data.metadata) { notes.push(`${file.name}：範本（不是回傳檔），已用來決定表單`); continue; } // 「匯出範本」的 .json
            const m = data.metadata;
            if (ag_otherSurvey(m)) { notes.push(`${file.name}：是調查「${m.survey}」的回傳檔，不是這次調查，略過`); continue; }
            ag_addSub({ file: file.name, org: m.org || '(未填機關)', year: m.year || '', user: m.user || '', tables: form.tables }, notes);
        } catch (err) { notes.push(`${file.name}：${err.message}`); }
    }
    ag_defaultOff();
    ag_render();
    ag_msg(notes);
}

// ?form=名稱：範本庫的「彙整」，先載入 forms/名稱.json（沒有才讀 .xlsx 第一個工作表）當表單，之後只要拖入回傳檔
async function ag_loadForm(name) {
    if (/[\\/]|\.\./.test(name)) throw new Error('表單名稱不合法：' + name);
    const base = 'forms/' + encodeURIComponent(name);
    const res = await fetch(base + '.json', { cache: 'no-cache' });
    if (res.ok) { const f = fb_normalize(await res.json()); ag_form = { title: f.title, sheets: f.sheets }; }
    else {
        const x = await fetch(base + '.xlsx', { cache: 'no-cache' });
        if (!x.ok) throw new Error(`找不到範本「${name}」`);
        ag_form = { ...ag_formFromXlsx(await x.arrayBuffer(), name + '.xlsx'), title: name };
    }
    ag_fixed = true;
    document.getElementById('ag-msg').textContent = `表單「${ag_form.title}」${ag_form.sheets.length > 1 ? `（${ag_form.sheets.length} 個工作表）` : ''}：已載入範本，請拖入各機關回傳的 JSON 或 Excel。`;
}

// ?survey=代號：讀專案 → 載入專案用的範本；繳交清單依應繳名單列出已繳／未繳
async function ag_loadSurvey(id) {
    if (fb_badName(id)) throw new Error('調查代號不合法：' + id);
    const idx = await fetch(FB_SURVEY_INDEX, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
    const s = idx.surveys?.[id];
    if (!s) throw new Error(`找不到調查「${id}」`);
    ag_survey = { id, ...s };
    await ag_loadForm(s.form);
    const n = (s.respondents || []).length;
    document.getElementById('ag-msg').textContent = `調查「${s.title || id}」（範本：${s.form}）：已載入${n ? `，應繳 ${n} 個機關` : ''}，請拖入各機關回傳的 JSON 或 Excel。`;
    document.title = `${s.title || id}｜彙整`;
}

function ag_msg(notes = []) {
    const base = ag_form ? `${ag_survey ? `調查「${ag_survey.title || ag_survey.id}」` : `表單「${ag_form.title}」`}${ag_form.sheets.length > 1 ? `（${ag_form.sheets.length} 個工作表）` : ''}：已匯入 ${ag_subs.length} 份。` : '尚未匯入檔案。';
    document.getElementById('ag-msg').textContent = base + (notes.length ? ' ⚠️ ' + notes.join('；') : '');
}

// 重複填報的預設：數字完全相同 → 只採計第一份；數字不同 → 全部先採計，等使用者勾選（已勾過的列不動）
function ag_defaultOff() {
    if (!ag_form) return;
    ag_form.sheets.forEach((sh, k) => fb_itemOwners(ag_view(k), sh.columns).owners.forEach(list => {
        const rows = list.map(o => ag_rowsOf(ag_subs[o.s], k)[o.i]);
        const same = rows.every(r => fb_sameNums(r, rows[0], sh.columns));
        rows.forEach((r, i) => { if (r._off === undefined && list.length > 1) r._off = same && i > 0; });
    }));
}

const ag_checkDup = k => ag_oneOwner[ag_sheet(k).name] ?? fb_itemOwners(ag_view(k), ag_sheet(k).columns).oneOwner;

// 實際採計的列：不採計的列數值歸零，再重算上層小計（excel 的上層快取值、JSON 的上層值都不直接相信）
function ag_effRows(s, k) {
    const cols = ag_sheet(k).columns, on = ag_checkDup(k);
    return fb_rollup(ag_rowsOf(s, k).map(r => {
        if (!(on && r._off)) return { ...r };
        const z = { ...r };
        cols.forEach(c => { if (c.type === 'num' && !c.formula) z[c.key] = ''; });
        return z;
    }), cols);
}

// ========== 2. 計算 ==========
// 彙總：依第一欄（項目名稱）分組，數值欄加總、公式欄以加總值重算、文字欄留空
function ag_summary(k) {
    const cols = ag_sheet(k).columns, nameKey = cols[0].key, groups = new Map();
    const totalRow = { [nameKey]: '合計', _n: new Set() };
    ag_subs.forEach(s => {
        const rows = ag_effRows(s, k), parent = fb_parents(rows.map(r => r[nameKey]));
        rows.forEach((r, i) => {
            const name = String(r[nameKey] ?? '').trim();
            if (!name) return;
            if (!groups.has(name)) groups.set(name, { [nameKey]: name, _n: new Set() });
            // 合計只加最上層的列：下層項目已含在上層的小計裡
            [groups.get(name), ...(parent[i] < 0 ? [totalRow] : [])].forEach(g => {
                if (fb_hasData(r, cols)) g._n.add(s.org + '|' + s.year); // 空白列不算有填報
                cols.forEach(c => { if (c.type === 'num' && !c.formula) g[c.key] = (g[c.key] || 0) + ag_num(r[c.key]); });
            });
        });
    });
    const rows = [...groups.values(), totalRow];
    rows.forEach(g => {
        cols.forEach(c => { if (c.type === 'num' && !c.formula) g[c.key] = g[c.key] || 0; });
        for (let pass = 0; pass < cols.length; pass++) {
            cols.forEach(c => { if (c.formula) g[c.key] = fb_evalFormula(c.formula, key => g[key] || 0); });
        }
        g._count = g._n.size;
    });
    // 狀態：沒有任何一份填數字 → 未填報；清單型的表同一項目仍有多份採計 → 重複
    const on = ag_checkDup(k), owners = fb_itemOwners(ag_view(k), cols).owners;
    rows.forEach(g => {
        g._status = g === totalRow ? '' : !g._count ? '❌ 未填報' : on && owners.get(g[nameKey])?.length > 1 && g._count > 1 ? '⚠️ 重複' : '✅';
    });
    return rows;
}

// 「有數字的列」只算實際填的列，上層小計不算
function ag_filled(s, k) {
    const cols = ag_sheet(k).columns, rows = ag_effRows(s, k), parent = fb_parents(rows.map(r => r[cols[0].key]));
    return rows.filter((r, i) => !parent.includes(i) && fb_hasData(r, cols)).length;
}

// ========== 3. 畫面 ==========
function ag_table(headId, bodyId, headers, rows) {
    document.getElementById(headId).innerHTML = fb_headGrid(headers).rows.map(row =>
        '<tr>' + row.map(h => `<th scope="${h.colspan > 1 ? 'colgroup' : 'col'}" colspan="${h.colspan}" rowspan="${h.rowspan}">${ag_esc(h.text)}</th>`).join('') + '</tr>').join('');
    document.getElementById(bodyId).innerHTML = rows.map(r =>
        '<tr>' + r.map(v => `<td class="${typeof v === 'number' ? 'text-right' : 'text-left'} px-2${typeof v === 'number' && v < 0 ? ' negative-value' : ''}">${ag_fmt(v)}</td>`).join('') + '</tr>').join('');
}

// 繳交清單；調查專案有應繳名單時，名單上的機關逐一列出（沒交的也列），名單外的標出來。index：對應 ag_subs（未繳交為 null）
function ag_subsView() {
    const multi = ag_form.sheets.length > 1;
    const row = s => [s.org, s.year, s.user, multi ? ag_form.sheets.map((_, k) => ag_filled(s, k)).join('／') : ag_filled(s, 0), s.file];
    const headers = ['機關', '年度', '填表人', multi ? `有數字的列（${ag_form.sheets.map(s => s.name).join('／')}）` : '有數字的列', '檔名'];
    const expected = ag_survey?.respondents || [];
    if (!expected.length) return { headers, rows: ag_subs.map(row), index: ag_subs.map((_, i) => i) };
    const res = fb_submissions(expected, ag_subs.map(s => s.org));
    const label = { done: '✅ 已繳交', missing: '❌ 未繳交', extra: '⚠️ 不在名單' };
    return {
        headers: [...headers, '狀態'],
        rows: res.rows.map(r => [...(r.index === undefined ? [r.org, '', '', '', ''] : row(ag_subs[r.index])), label[r.status]]),
        index: res.rows.map(r => r.index ?? null),
        done: res.done, total: res.total,
    };
}

function ag_sheetViews(k) {
    const cols = ag_sheet(k).columns;
    const cell = (c, v) => c.type === 'num' ? (v === '' || v === undefined ? '' : ag_num(v)) : String(v ?? '');
    const detail = [];
    let blank = 0;
    ag_subs.forEach(s => ag_effRows(s, k).forEach(r => {
        if (!String(r[cols[0].key] ?? '').trim()) return;
        if (!fb_hasData(r, cols)) { blank++; return; } // 明細只列有數字的列
        // 明細的公式欄以該列數值重算，不信任檔案內的值
        const calc = { ...r };
        cols.forEach(c => { if (c.formula) calc[c.key] = fb_evalFormula(c.formula, key => ag_num(calc[key])); });
        detail.push([s.org, s.year, ...cols.map(c => cell(c, calc[c.key]))]);
    }));
    return {
        sum: { headers: [...cols, '填報機關數', '狀態'], rows: ag_summary(k).map(g => [...cols.map(c => c.type === 'num' ? g[c.key] : (c === cols[0] ? g[c.key] : '')), g._count, g._status]) },
        detail: { headers: ['機關', '年度', ...cols], rows: detail },
        blank,
    };
}

// 重複填報審核：清單型的表，同一項目有多份填數字 → 逐列勾選是否採計
function ag_renderDups(k) {
    const cols = ag_sheet(k).columns, on = ag_checkDup(k), auto = fb_itemOwners(ag_view(k), cols);
    const nums = cols.filter(c => c.type === 'num' && !c.formula);
    document.getElementById('ag-dup-mode').checked = on;
    document.getElementById('ag-dup-auto').textContent = ag_oneOwner[ag_sheet(k).name] === undefined ? `（自動判斷：${auto.oneOwner ? '各項目多半只有一份在填，視為清單型' : '各份填同一批項目，不檢查'}）` : '（手動設定）';
    const groups = on ? [...auto.owners].filter(([, l]) => l.length > 1) : [];
    document.getElementById('ag-dup-card').classList.toggle('hidden', !groups.length);
    if (!groups.length) return;
    document.getElementById('ag-dup-title').textContent = `⚠️ 重複填報（${groups.length} 項）`;
    document.getElementById('ag-dup-head').innerHTML = '<tr>' + ['採計', '項目', '機關', ...nums.map(c => c.label), '備註'].map(h => `<th>${ag_esc(h)}</th>`).join('') + '</tr>';
    document.getElementById('ag-dup').innerHTML = groups.map(([name, list]) => {
        const rows = list.map(o => ({ o, r: ag_rowsOf(ag_subs[o.s], k)[o.i] }));
        const same = rows.every(x => fb_sameNums(x.r, rows[0].r, cols));
        return rows.map(({ o, r }, i) => `<tr class="${r._off ? 'opacity-50' : ''}${i === 0 ? ' border-t-2 border-slate-400' : ''}">
            <td class="text-center"><input type="checkbox" class="ag-dup-use" data-s="${o.s}" data-i="${o.i}"${r._off ? '' : ' checked'}></td>
            <td class="text-left px-2 font-bold">${i === 0 ? ag_esc(name) : ''}</td>
            <td class="text-left px-2" title="${ag_esc(ag_subs[o.s].file)}">${ag_esc(ag_subs[o.s].org)}</td>
            ${nums.map(c => `<td class="text-right px-2">${ag_fmt(ag_num(r[c.key]))}</td>`).join('')}
            <td class="text-left px-2 text-xs">${i === 0 ? (same ? '數字相同，預設只採計第一份' : '數字不同，請勾選要採計的') : ''}</td></tr>`).join('');
    }).join('');
}

function ag_render() {
    document.getElementById('ag-result').classList.toggle('hidden', !ag_subs.length);
    if (!ag_subs.length) return;
    ag_cur = Math.min(ag_cur, ag_form.sheets.length - 1);
    const esc = rows => rows.map(r => r.map(x => typeof x === 'number' ? x : ag_esc(x))); // 字串先跳脫，數字留給 ag_table 格式化並靠右
    const subs = ag_subsView();
    document.getElementById('ag-subs-title').textContent = subs.total ? `繳交清單（已繳 ${subs.done}／應繳 ${subs.total}，共收 ${ag_subs.length} 份）` : `繳交清單（${ag_subs.length} 份）`;
    ag_table('ag-subs-head', 'ag-subs', [...subs.headers, ''], esc(subs.rows).map((r, i) =>
        [...r, subs.index[i] === null ? '' : `<button class="ag-del text-red-500 font-bold" data-i="${subs.index[i]}" title="移除">×</button>`]));
    document.getElementById('ag-tabs').innerHTML = ag_form.sheets.length > 1
        ? ag_form.sheets.map((s, k) => `<button class="fb-tab${k === ag_cur ? ' fb-tab-on' : ''}" data-k="${k}">${ag_esc(s.name)}</button>`).join('') : '';
    const v = ag_sheetViews(ag_cur);
    ag_table('ag-sum-head', 'ag-sum', v.sum.headers, esc(v.sum.rows));
    ag_table('ag-detail-head', 'ag-detail', v.detail.headers, esc(v.detail.rows));
    document.getElementById('ag-detail-title').textContent = `明細（${v.detail.rows.length} 列有數字` + (v.blank ? `，略過 ${v.blank} 列空白）` : '）');
    ag_renderDups(ag_cur);
}

// ========== 4. 匯出 Excel ==========
// 每個工作表各一張「彙總」、一張「明細」，最後一張繳交清單（Excel 工作表名稱上限 31 字）
function ag_exportXlsx() {
    if (!ag_subs.length) return alert('請先匯入檔案');
    const wb = XLSX.utils.book_new(), multi = ag_form.sheets.length > 1;
    const add = (name, t) => {
        const head = fb_headGrid(t.headers);
        const ws = XLSX.utils.aoa_to_sheet([...head.aoa, ...t.rows]);
        ws['!merges'] = head.merges;
        ws['!cols'] = t.headers.map((h, i) => ({ wch: i === 0 ? 20 : 14 }));
        XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    };
    ag_form.sheets.forEach((s, k) => {
        const v = ag_sheetViews(k), tag = `${k + 1}_${s.name}`;
        add(multi ? `彙總${tag}` : '依項目彙總', v.sum);
        add(multi ? `明細${tag}` : '明細', v.detail);
    });
    add('繳交清單', ag_subsView());
    XLSX.writeFile(wb, `${ag_form.title}_彙整_${ag_subs.length}份.xlsx`);
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
    $('ag-btn-clear').onclick = () => { if (!ag_fixed) ag_form = null; ag_subs = []; ag_oneOwner = {}; ag_cur = 0; ag_render(); ag_msg(); };
    $('ag-tabs').addEventListener('click', e => { const b = e.target.closest('.fb-tab'); if (b) { ag_cur = +b.dataset.k; ag_render(); } });
    $('ag-dup-mode').onchange = e => { ag_oneOwner[ag_sheet(ag_cur).name] = e.target.checked; ag_defaultOff(); ag_render(); };
    $('ag-dup').addEventListener('change', e => {
        if (!e.target.classList.contains('ag-dup-use')) return;
        ag_rowsOf(ag_subs[+e.target.dataset.s], ag_cur)[+e.target.dataset.i]._off = !e.target.checked;
        ag_render();
    });
    $('ag-subs').addEventListener('click', e => {
        if (!e.target.classList.contains('ag-del')) return;
        ag_subs.splice(+e.target.dataset.i, 1);
        if (!ag_subs.length) { if (!ag_fixed) ag_form = null; ag_oneOwner = {}; ag_cur = 0; }
        ag_render(); ag_msg();
    });
    const params = new URLSearchParams(location.search), form = params.get('form'), survey = params.get('survey');
    if (survey) ag_loadSurvey(survey).catch(err => ag_msg([err.message]));
    else if (form) ag_loadForm(form).catch(err => ag_msg([err.message]));
});
