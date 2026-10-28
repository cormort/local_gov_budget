// 貼表：把各機關寄回的 Excel（照空白範本填寫）依標準機關順序一列接一列貼到範本下方，保留原格式，匯出一份彙總表
// 版面：在順序表勾選的範本檔（不貼入），沒勾選就取最多檔案共用的版面；取第一個沒隱藏的工作表（範本檔可改選），表頭到「合計」列（沒有合計列就到表頭最後一列）照抄，「註」以下的說明列移到最後
// 小計：範本「合計」列有 COUNTIF 的欄重算勾選數、有「共…案」的欄重算案數；主管列算到下一個主管，基金列算到下一個標題列
// ponytail: 提案列＝第 2 欄（提案序號）有填、標題列＝第 1 欄有填且第 2 欄空白，是照協商彙總表的版面寫的；別的版面要另訂規則
// 兩種貼法：接續貼上（協商表：各檔的列接在一起）、依項目填入（決算表：每份都有全部基金，各填自己那一列 → 依第一欄名稱對位填進範本）
// 依項目填入：欄依表頭文字對位（有人插欄也對得上）；同一欄同一個非零數字出現在 3 列以上視為往下拖曳的錯誤不採用；
// 一列有多份檔案填時，採檔名認得出是該基金的那份；只有別的機關填的照填但標淺橘色；對不上的、沒人填的都列在「查對」

// 首頁就是貼表：舊的查填、預覽連結（?form= / ?survey=）在進階功能頁，原樣轉過去
if (typeof location !== 'undefined' && /[?&](form|survey)=/.test(location.search)) location.replace('advanced.html' + location.search + location.hash);

let ps_base = null;   // { name, buf, fromFile, sheet, placeholders }：fromFile = 版面取自第一個回傳檔（沒有範本）
let ps_files = [];    // [{ file, org, rank, buf, wb, ws, rows, pinned, tpl }]；順序就是貼上順序，可手動調整；pinned = 手動選的工作表；tpl = 勾選為範本（不貼入）
let ps_orgs = { sortOrder: [], aliasMap: {} };
let ps_headEnd = 0;   // 範本表頭最後一列（1 起算），沒有「合計」列時資料從下一列開始
let ps_matchCol = 1;  // 比對欄位：範本第幾欄放基金名稱（像 excel_merge 可選第 1、2、3 欄）
let ps_order = null;  // 自訂順序清單（每行一個基金或機關名稱，可用簡稱）；null = 標準基金清單
let ps_mode = 'stack'; // 'stack' 接續貼上（預設，大多數調查表）| 'fill' 依項目填入（每份都有全部基金的表，手動切換）

const ps_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const ps_text = c => String(c?.text ?? '').trim();
const ps_slave = c => c.isMerged && c.master.address !== c.address; // 合併儲存格的非左上格：ExcelJS 會回傳左上格的值
const ps_own = c => ps_slave(c) ? '' : ps_text(c);
const ps_sheet = wb => wb.worksheets.find(w => w.state === 'visible') || wb.worksheets[0];
const ps_filler = a => /ＯＯ|ＸＸ|○○|^【範例/.test(a); // 空白表留下的佔位列（ＯＯ主管、ＸＸ基金、範例）
const ps_stem = n => n.replace(/\.[^.]+$/, '');
const ps_cellText = c => { const v = c?.value; return String(v?.richText ? v.richText.map(t => t.text).join('') : c?.text ?? '').trim(); };
const ps_isFormula = v => !!v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v);
const ps_key = s => String(s ?? '').replace(/\s/g, '');
const ps_std = name => { const h = fb_detectOrgs(name, ps_orgs.sortOrder, ps_orgs.aliasMap); return h.length === 1 ? h[0] : null; };
const ps_value = v => (v && typeof v === 'object' && 'formula' in v) || (v && typeof v === 'object' && 'sharedFormula' in v) ? v.result ?? null : v; // 公式只留值，參照位置貼過來就不對了

// 一個工作表的版面：合計列、說明列起點、欄數；key = 名稱欄（有人在最左邊插欄時不是第 1 欄）
function ps_layout(ws, key = 1) {
    let total = 0, notes = 0;
    ws.eachRow((row, r) => {
        const a = ps_text(row.getCell(key));
        if (!total && a === '合計') total = r;
        else if (total && !notes && a.startsWith('註')) notes = r;
    });
    return { total: total || ps_headEnd, hasTotal: !!total, notes: notes || ws.rowCount + 1, cols: ws.columnCount };
}

// 回傳檔要貼的列：合計列以下、說明列以上；空白列與範本原有的佔位列（ＯＯ主管、範例…）略過
// start：從哪一列之後開始（預設檔案自己的合計列之後）
function ps_pickRows(ws, placeholders, key = 1, start = null) {
    const { total, notes } = ps_layout(ws, key);
    const rows = [];
    for (let r = (start ?? total ?? 0) + 1; r < notes; r++) {
        const row = ws.getRow(r);
        const cells = [];
        row.eachCell(c => cells.push(ps_text(c)));
        if (!cells.some(Boolean)) continue;
        const a = ps_text(row.getCell(key));
        if (placeholders.has(a) || ps_filler(a)) continue;
        rows.push(r);
    }
    return rows;
}

// map：範本欄 → 來源欄（來源檔有插欄時），沒給就同一欄
function ps_copyRow(src, sr, dst, dr, cols, map = null) {
    const s = src.getRow(sr), d = dst.getRow(dr), at = c => (map ? map[c] : c);
    if (s.height) d.height = s.height;
    for (let c = 1; c <= cols; c++) {
        const to = d.getCell(c);
        if (!at(c)) { to.value = null; continue; }
        const from = s.getCell(at(c));
        to.value = ps_slave(from) ? null : ps_value(from.value);
        to.style = JSON.parse(JSON.stringify(from.style || {}));
    }
    // 同一列的合併儲存格（標題列的 A:D、F:H）照搬，欄位換算成範本的位置
    const inv = {}; for (let c = 1; c <= cols; c++) if (at(c)) inv[at(c)] = c;
    (src.model.merges || []).forEach(m => {
        const [a, b] = m.split(':').map(x => src.getCell(x));
        if (a.row !== sr || b.row !== sr || !inv[a.col] || !inv[b.col] || inv[a.col] >= inv[b.col]) return;
        dst.mergeCells(dr, inv[a.col], dr, inv[b.col]);
    });
}

// 依目前的勾選決定版面
async function ps_rebase() {
    // 沒有範本時取最多檔案共用的版面（列數、欄數相同），避開只交自己那幾列的短版、自己插欄的檔案
    const data = ps_files.filter(f => !f.tpl), shape = f => f.ws.rowCount + 'x' + f.ws.columnCount;
    const n = {}; data.forEach(f => { n[shape(f)] = (n[shape(f)] || 0) + 1; });
    const t = ps_files.find(f => f.tpl), first = data.reduce((a, f) => (!a || n[shape(f)] > n[shape(a)] ? f : a), null);
    if (t) await ps_setBase(ps_stem(t.file), t.buf, false, t.pinned);
    else if (first) await ps_setBase(ps_stem(first.file), first.buf, true);
    else ps_base = null;
}

// 換版面後，已讀入的檔案依新版面重新挑工作表與要貼的列
async function ps_setBase(name, buf, fromFile = false, sheet = null) {
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(buf);
    const bws = (sheet && book.getWorksheet(sheet)) || ps_sheet(book), bl = ps_layout(bws);
    // 空白範本：合計列以下原有的列（ＯＯ主管、範例、註）都是佔位；版面取自回傳檔時那些是真資料，只認佔位字樣
    const placeholders = new Set();
    if (!fromFile) for (let r = bl.total + 1; r <= bws.rowCount; r++) { const a = ps_text(bws.getRow(r).getCell(1)); if (a) placeholders.add(a); }
    // 表頭範圍沿用網頁表單的判斷（多層合併表頭、略過表名與單位列）
    const wb = XLSX.read(buf, { type: 'array' }), ws = wb.Sheets[bws.name];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const text = (r, c) => { const x = ws[XLSX.utils.encode_cell({ r, c })]; return x ? String(x.w ?? x.v ?? '').trim() : ''; };
    const block = fb_headerBlock(text, ws['!merges'] || [], range.s.r, range.e.r, range.s.c, range.e.c);
    ps_headEnd = (block?.he ?? -1) + 1;
    const headRows = block ? Array.from({ length: block.he - block.hr + 1 }, (_, i) => block.hr + 1 + i) : [1];
    ps_base = { name, buf, fromFile, sheet: bws.name, placeholders, ws: bws, headRows };
    ps_files.forEach(f => {
        f.ws = f.wb.getWorksheet(f.pinned || bws.name) || ps_sheet(f.wb);
        f.cmap = ps_colMap(f.ws);
        // 自己的基金列可能在檔案自己的合計列上面（各機關自訂的短版），所以找的範圍從表頭下面開始
        const c = f.cmap.key ? ps_candidates(f, ps_pickRows(f.ws, placeholders, f.cmap.key, headRows.at(-1)), ps_pickRows(f.ws, placeholders, f.cmap.key)) : { cand: [], del: new Set() };
        f.cand = c.cand; f.del = c.del; f.del0 = new Set(c.del); // del0：預設勾的，「回到預設」用
        ps_applyDel(f);
    });
}

// 接續貼上的候選列與預設刪除（像 excel_merge：每一列都列出來，可疑的預先勾刪除，彙整的人可以改）
// 預設保留自己那一段：從檔名認出的基金那列（比對欄位）到下一個別的基金、「主管」、「一、」大標題或合計為止
// （底下的明細、提案都算，上面緊接的「主管」列一併）；其他列預設刪除。
// all：表頭以下所有的列（自己的基金列可能在檔案自己的合計列上面）；picked：合計列以下的列
// 完全沒有基金列的一般調查表、檔名認不出機關：picked 整份保留；檔案裡只有別的基金：全部預設刪除（f.missing）
function ps_candidates(f, all, picked) {
    f.missing = false; f.others = [];
    const col = f.cmap.map[ps_matchCol] || f.cmap.key;
    const name = r => ps_cellText(f.ws.getRow(r).getCell(col)).trim();
    if (!f.owners.length || !all.some(r => ps_std(name(r)))) return { cand: picked, del: new Set() };
    // 「一、作業基金」這種大標題才停；「(一)」編號是基金自己的明細
    const stop = a => !a ? false : a.startsWith('合計') || a.endsWith('主管') || fb_level(a) === 1 || !!ps_std(a);
    const keep = new Set();
    all.forEach((r, i) => {
        if (!f.owners.includes(ps_std(name(r)))) return;
        if (i > 0 && name(all[i - 1]).endsWith('主管')) keep.add(all[i - 1]);
        keep.add(r);
        for (let j = i + 1; j < all.length && !stop(name(all[j])); j++) keep.add(all[j]);
    });
    // 別的基金的列有填資料（例：檔名只寫「文作、文發」，卻也填了國立文化機構）→ 查對提醒
    const has = r => Object.entries(f.cmap.map).some(([tc, fc]) => +tc !== 1 && fc !== col && (v => v != null && !ps_isFormula(v) && String(v.richText ? 'x' : v).trim() !== '')(f.ws.getRow(r).getCell(fc).value));
    f.others = all.filter(r => !keep.has(r) && ps_std(name(r)) && !f.owners.includes(ps_std(name(r))) && has(r)).map(r => ps_std(name(r)));
    f.missing = !keep.size;
    return { cand: all, del: new Set(all.filter(r => !keep.has(r))) };
}
const ps_applyDel = f => { f.rows = f.cand.filter(r => !f.del.has(r)); };

// 來源檔的欄怎麼對到範本：欄數相同照位置（.ods 轉檔後表頭文字會錯位，不能靠欄名）；
// 欄數不同（有人插欄，連最左邊也算）每一欄都依表頭文字對位，名稱欄也是 → { map: {範本欄: 來源欄}, extra: [範本沒有的欄名], key: 名稱欄或 null }
function ps_colMap(ws) {
    const bws = ps_base.ws, map = {}, extra = [];
    if (ws.columnCount === bws.columnCount) { for (let c = 1; c <= bws.columnCount; c++) map[c] = c; return { map, extra, key: 1 }; }
    const byHead = new Map([...ps_heads(bws)].map(([c, t]) => [t, c]));
    ps_heads(ws).forEach((t, c) => { const tc = byHead.get(t); if (tc && !map[tc]) map[tc] = c; else extra.push(ps_label(t)); });
    return { map, extra, key: map[1] || null };
}
// 回傳檔自己藏起來（範本沒藏）、但有資料的欄與列：照樣讀入，列在查對提醒確認
function ps_hiddenChecks(f, rows, only = null) {
    const out = [], { map, key } = f.cmap, heads = ps_heads(f.ws);
    const has = (r, c) => { const v = f.ws.getRow(r).getCell(c).value; return v != null && !ps_isFormula(v) && (v.richText ? v.richText.some(t => t.text.trim()) : String(v).trim() !== ''); };
    const data = Object.entries(map).filter(([tc]) => +tc !== 1 && (!only || only.has(+tc)));
    data.forEach(([tc, fc]) => {
        if (f.ws.getColumn(fc).hidden && !ps_base.ws.getColumn(+tc).hidden && rows.some(r => has(r, fc)))
            out.push(['隱藏欄有資料', ps_label(heads.get(fc) || '') || f.ws.getColumn(fc).letter + ' 欄', f.file, '已讀入，請確認要不要彙整']);
    });
    rows.filter(r => f.ws.getRow(r).hidden && data.some(([, fc]) => has(r, fc)))
        .forEach(r => out.push(['隱藏列有資料', ps_cellText(f.ws.getRow(r).getCell(key)) || `第 ${r} 列`, f.file, '已讀入，請確認要不要彙整']));
    return out;
}
const ps_label = t => t.replace(/^\/+|\/+$/g, '');
const ps_keyLabel = () => ps_label(ps_heads(ps_base.ws).get(1) || '第一欄');

// ========== 依項目填入 ==========
// 表頭文字（多層以 / 串起來）當欄名
function ps_heads(ws) {
    const m = new Map();
    for (let c = 1; c <= ws.columnCount; c++) {
        const t = ps_base.headRows.map(r => ps_key(ps_cellText(ws.getRow(r).getCell(c)))).join('/');
        if (t.replace(/\//g, '')) m.set(c, t);
    }
    return m;
}

// 項目列：表頭以下、說明列以上，第一欄有名稱、不是佔位列
// sec = 上面最近的基金列：基金底下的明細列（「(一)業務計畫」）各基金可能同名，要連同所屬基金一起對位，也算那個基金的
function ps_items(ws, key = 1) {
    const out = [], { notes } = ps_layout(ws, key);
    let sec = null;
    for (let r = ps_base.headRows.at(-1) + 1; r < notes; r++) {
        const a = ps_key(ps_cellText(ws.getRow(r).getCell(key)));
        if (!a || ps_filler(a)) continue;
        const std = ps_std(a);
        if (std) sec = std;
        out.push({ r, a, std, sec: std ? null : sec });
    }
    return out;
}



// 各檔要填進範本的值與查對清單：{ items, fills: [{ r, data: {欄: 值}, flag }], checks: [[類別, 項目, 檔案, 說明]], filled: Map(檔案 → 列數) }
function ps_fillPlan() {
    const bws = ps_base.ws, heads = ps_heads(bws);
    const items = ps_items(bws).map(i => ({ ...i, inputs: [...heads.keys()].filter(c => c > 1 && !ps_isFormula(bws.getRow(i.r).getCell(c).value)), cands: [] }));
    // 基金列：名稱相同或對到同一個標準基金；明細列：所屬基金與名稱都相同
    const find = i => i.std ? items.find(x => x.a === i.a) || items.find(x => x.std === i.std) : items.find(x => !x.std && x.a === i.a && x.sec === i.sec);
    const checks = [], filled = new Map();
    const plain = v => { if (v == null || ps_isFormula(v)) return null; if (v.richText) v = v.richText.map(t => t.text).join(''); return typeof v === 'string' && !v.trim() ? null : v; };
    for (const f of ps_files.filter(f => !f.tpl)) {
        const { map: colMap, extra, key } = f.cmap;
        extra.forEach(t => checks.push(['範本沒有的欄', t, f.file, '未採用']));
        if (!key) { checks.push(['找不到名稱欄', ps_keyLabel(), f.file, '這份沒有讀入，請確認表頭']); filled.set(f, 0); continue; }
        const rows = ps_items(f.ws, key), cols = Object.keys(colMap).map(Number).filter(c => c > 1);
        // 往下拖曳是從最上面那格複製下去：最上面那格保留，下面重複的不採用
        const count = {}, top = {};
        rows.forEach(i => cols.forEach(tc => { const v = f.ws.getRow(i.r).getCell(colMap[tc]).value; if (typeof v === 'number' && v) { const k = tc + ':' + v; count[k] = (count[k] || 0) + 1; top[k] ??= i.r; } }));
        const noise = (tc, v, r) => typeof v === 'number' && count[tc + ':' + v] >= 3 && top[tc + ':' + v] !== r;
        Object.entries(count).filter(([, n]) => n >= 3).forEach(([k, n]) => checks.push(['重複數字（疑似往下拖曳）', ps_label(heads.get(+k.split(':')[0]) || ''), f.file, `${Number(k.split(':')[1]).toLocaleString()} 出現在 ${n} 列，只採用最上面一列`]));
        let n = 0;
        rows.forEach(i => {
            const it = find(i), data = {};
            (it ? it.inputs.filter(c => colMap[c]) : cols).map(c => [c, colMap[c]]).forEach(([tc, fc]) => {
                const v = plain(f.ws.getRow(i.r).getCell(fc).value);
                if (v != null && !noise(tc, v, i.r)) data[tc] = v;
            });
            if (!Object.keys(data).length) return;
            n++;
            if (!it) { checks.push(['範本沒有的列', i.a, f.file, '未填入']); return; }
            const fund = it.std || it.sec;
            it.cands.push({ f, r: i.r, owner: !!fund && f.owners.includes(fund), data, key: JSON.stringify(data) });
        });
        filled.set(f, n);
        if (!n) checks.push(['沒有填任何資料', '', f.file, '請確認是否漏填或填在別的工作表']);
    }
    const fills = [];
    items.forEach(it => {
        const c = it.cands;
        if (!c.length) { if (it.std) checks.push(['沒有對應到的基金', it.a, '', '沒有檔案填這一列']); return; }
        const own = c.filter(x => x.owner);
        const pick = own[0] || (new Set(c.map(x => x.key)).size === 1 ? c[0] : null);
        if (!pick) { checks.push(['多份檔案填的不同', it.a, c.map(x => x.f.file).join('、'), '未填入，請擇一']); return; }
        if (!own.length) checks.push(['由別的機關的檔案填入', it.a, c.map(x => x.f.file).join('、'), '已填入並標成淺橘色，請確認']);
        c.filter(x => x !== pick && x.key !== pick.key).forEach(x => checks.push(['其他檔案也填了這列', it.a, x.f.file, `數字不同，採用「${pick.f.file}」`]));
        fills.push({ r: it.r, data: pick.data, flag: !own.length && c.map(x => x.f.file).join('、') });
        pick.used = Object.keys(pick.data).map(Number);
    });
    // 隱藏欄、列：很多機關會把別人的列藏起來只看自己那列，那些不會被採用；只提醒真的填進彙總表的
    ps_files.filter(f => !f.tpl && f.cmap.key).forEach(f => {
        const used = items.flatMap(it => it.cands.filter(x => x.f === f && x.used));
        if (used.length) checks.push(...ps_hiddenChecks(f, used.map(x => x.r), new Set(used.flatMap(x => x.used))));
    });
    // 依重要性排：沒人填的、別的機關填的、多份不同的先看
    checks.push(...ps_orderChecks());
    const order = ['找不到名稱欄', '清單上沒有對應的檔案', '沒有對應到的基金', '不在自訂清單', '由別的機關的檔案填入', '多份檔案填的不同', '其他檔案也填了這列', '隱藏欄有資料', '隱藏列有資料', '沒有填任何資料', '重複數字（疑似往下拖曳）', '範本沒有的列', '範本沒有的欄'];
    checks.sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0]));
    return { items, fills, checks, filled };
}

async function ps_buildFill() {
    const plan = ps_fillPlan();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(ps_base.buf);
    const ws = wb.getWorksheet(ps_base.sheet);
    wb.worksheets.filter(w => w !== ws).forEach(w => wb.removeWorksheet(w.id));
    // 範本若是回傳檔，先清掉它自己填的，全部由各檔重新填
    plan.items.forEach(it => it.inputs.forEach(c => { const cell = ws.getRow(it.r).getCell(c); if (!ps_slave(cell)) cell.value = null; }));
    plan.fills.forEach(({ r, data, flag }) => Object.entries(data).forEach(([c, v]) => {
        const cell = ws.getRow(r).getCell(+c);
        cell.value = v;
        // 淺橘色＋註解：範本本身常用淺黃、淺綠標輸入格，避開
        // 同格式的格子共用同一個樣式物件，要先複製一份再改，不然整片都變色
        if (flag) { cell.style = { ...JSON.parse(JSON.stringify(cell.style || {})), fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8CBAD' } } }; cell.note = `由「${flag}」填入，請確認`; }
    }));
    ps_checkSheet(wb, plan.checks);
    wb.calcProperties.fullCalcOnLoad = true; // 合計、執行率等公式打開時重算
    return wb.xlsx.writeBuffer();
}

// 接續貼上的查對：檔名認不出機關、沒有列可貼
const ps_stackChecks = () => [...ps_orderChecks(), ...ps_files.filter(f => !f.tpl).flatMap(f => [
    ...(f.cmap.key ? [] : [['找不到名稱欄', ps_keyLabel(), f.file, '這份沒有貼入，請確認表頭']]),
    ...f.cmap.extra.map(t => ['範本沒有的欄', t, f.file, '未貼入']),
    ...(f.cmap.key ? ps_hiddenChecks(f, f.rows) : []),
    ...(f.owners.length === 1 || ps_order ? [] : [f.owners.length ? ['檔名含多個機關', f.owners.join('、'), f.file, `排在「${f.owners[0]}」的位置，請確認順序`] : ['檔名認不出機關', '', f.file, '排在最後，請確認順序']]),
    ...(f.missing ? [['找不到自己基金的列', f.owners.join('、'), f.file, '檔案裡只有別的基金，預設全部不貼，請在逐列檢查確認']] : []),
    ...(f.others?.length ? [['也填了別的基金', f.others.join('、'), f.file, `${f.others.length} 個，預設不貼，請在逐列檢查確認`]] : []),
    ...(f.rows.length || !f.cmap.key || f.missing ? [] : [['沒有可貼的列', '', f.file, '請確認是否漏填或填在別的工作表']]),
])];

function ps_checkSheet(wb, checks) {
    if (!checks.length) return;
    const ck = wb.addWorksheet('查對');
    ck.addRow(['類別', '項目', '檔案', '說明']).font = { bold: true };
    checks.forEach(x => ck.addRow(x));
    [24, 30, 50, 36].forEach((w, i) => { ck.getColumn(i + 1).width = w; });
}

// 讀檔進度：一份一份讀，每份之前讓畫面更新一次
function ps_progress(done, total, text) {
    const bar = document.getElementById('ps-progress'), msg = document.getElementById('ps-msg');
    bar.classList.toggle('hidden', done === null);
    if (done === null) return;
    bar.max = total; bar.value = done;
    msg.textContent = text;
}

async function ps_addFiles(files) {
    const notes = [];
    document.getElementById('ps-btn-xlsx').disabled = true;
    for (const [i, file] of files.entries()) {
        ps_progress(i, files.length, `讀取中 ${i + 1}／${files.length}：${file.name}`);
        await new Promise(r => setTimeout(r)); // 讓瀏覽器先畫出進度
        if (!/\.(xlsx|xls|ods)$/i.test(file.name)) { notes.push(`${file.name}：不是 Excel／ODS 檔，略過`); continue; }
        try {
            let buf = await file.arrayBuffer();
            // .xls、.ods 先轉成 .xlsx 再讀（格式會簡化，數字與文字不變）
            if (!/\.xlsx$/i.test(file.name)) {
                const x = XLSX.read(buf, { type: 'array' });
                // ODS 的工作表名稱可以超過 Excel 上限 31 字，截短才寫得出去
                // ponytail: 截短後撞名不處理，真的遇到再加編號
                x.SheetNames = x.SheetNames.map(n => { const t = n.slice(0, 31); if (t !== n) { x.Sheets[t] = x.Sheets[n]; delete x.Sheets[n]; } return t; });
                (x.Workbook?.Sheets || []).forEach((w, i) => { w.name = x.SheetNames[i]; });
                // ODS 允許重疊的合併儲存格，Excel 不行：重疊的只留第一個
                Object.values(x.Sheets).forEach(sh => {
                    const kept = [], hit = (a, b) => a.s.r <= b.e.r && b.s.r <= a.e.r && a.s.c <= b.e.c && b.s.c <= a.e.c;
                    (sh['!merges'] || []).forEach(m => { if (!kept.some(k => hit(k, m))) kept.push(m); });
                    if (sh['!merges']) sh['!merges'] = kept;
                });
                buf = XLSX.write(x, { type: 'array', bookType: 'xlsx' });
            }
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buf);
            const stem = ps_stem(file.name);
            // 檔名有「空白」「範本」的預設勾成範本（還沒有範本時），排在最上面
            const tpl = !ps_files.some(f => f.tpl) && /空白|範本/.test(stem);
            const hits = tpl ? [] : fb_detectOrgs(stem, ps_orgs.sortOrder, ps_orgs.aliasMap);
            // 一份含多個機關（「園區、健保」）：依標準順序裡最前面的那個排；認不出的排最後
            const org = hits.length ? hits.join('、') : stem;
            const entry = { file: file.name, org, owners: hits, buf, wb, ws: ps_sheet(wb), rows: [], tpl };
            entry.rank = ps_rankOf(entry);
            // 同機關重複：原位置換掉；新的依標準順序插入，不打亂已手動調整的順序
            const dup = hits.length === 1 ? ps_files.findIndex(f => f.org === org) : -1;
            if (dup >= 0) ps_files[dup] = entry;
            else { const at = ps_files.findIndex(f => f.rank > entry.rank); ps_files.splice(at < 0 ? ps_files.length : at, 0, entry); }
        } catch (e) { notes.push(`${file.name}：${e.message}`); }
    }
    ps_progress(files.length, files.length, `已讀取 ${files.length} 份，整理版面中…`);
    await new Promise(r => setTimeout(r));
    await ps_rebase();
    ps_progress(null);
    ps_render(notes);
}

// 排序名次：範本最前；有自訂清單照清單（清單外的排最後），沒有就照標準基金清單（一份含多個機關的依最前面那個）
// 清單的一行對到檔案：那行認得出標準機關（「大學」→國立大學校院校務基金）就只比機關，免得「大學」也對到「陽明交通大學附設醫院」；
// 認不出（清單外的名稱）才退一步看那行文字有沒有出現在檔名裡
const ps_lineHits = (line, f) => {
    const hits = fb_detectOrgs(line, ps_orgs.sortOrder, ps_orgs.aliasMap);
    return hits.length ? hits.some(o => f.owners.includes(o)) : ps_key(ps_stem(f.file)).includes(ps_key(line));
};
function ps_rankOf(f) {
    if (f.tpl) return -1;
    if (!ps_order) return f.owners.length ? ps_orgs.sortOrder.indexOf(f.owners[0]) : Infinity;
    const i = ps_order.findIndex(line => ps_lineHits(line, f));
    return i < 0 ? Infinity : i;
}
function ps_setOrder(text) {
    const lines = String(text ?? '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    ps_order = lines.length ? lines : null;
    ps_files.forEach(f => { f.rank = ps_rankOf(f); });
    ps_resort();
}
// 自訂清單的查對：清單上沒有檔案的（未繳）、檔案不在清單上的
const ps_orderChecks = () => !ps_order ? [] : [
    ...ps_order.filter(line => !ps_files.some(f => !f.tpl && ps_lineHits(line, f))).map(line => ['清單上沒有對應的檔案', line, '', '可能未繳或檔名認不出']),
    ...ps_files.filter(f => !f.tpl && f.rank === Infinity).map(f => ['不在自訂清單', f.org, f.file, '排在最後']),
];

// 依清單重排：手動調過的順序回到清單順序（範本留在最上面，同名次維持原順序）
const ps_resort = () => { ps_files.sort((a, b) => a.rank - b.rank); };

const ps_build = () => ps_mode === 'fill' ? ps_buildFill() : ps_buildStack();

async function ps_buildStack() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(ps_base.buf);
    const ws = wb.getWorksheet(ps_base.sheet), { total, hasTotal, notes, cols } = ps_layout(ws);
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
    let lastTop = null; // 同一個主管的基金分在不同檔案時，各檔都帶了主管列：連續的同名主管列只留第一個
    // 自訂清單上沒有檔案的，依清單位置插「❌ 未回傳」列（像 excel_merge 的 checklist）
    const missing = (ps_order || []).map((line, i) => ({ line, i })).filter(({ line }) => !ps_files.some(f => !f.tpl && ps_lineHits(line, f)));
    const nameCol = ps_matchCol <= cols ? ps_matchCol : 1;
    const putMissing = upto => { while (missing.length && missing[0].i < upto) {
        const { line } = missing.shift();
        ws.getRow(r).getCell(nameCol).value = line;
        ws.getRow(r).getCell(nameCol === 1 ? 2 : 1).value = '❌ 未回傳／無資料';
        placed.push({ r, kind: 'missing' }); r++;
    } };
    ps_files.filter(f => !f.tpl).forEach(f => { putMissing(f.rank); f.rows.forEach(sr => {
        const top = ps_cellText(f.ws.getRow(sr).getCell(f.cmap.key));
        if (top.endsWith('主管') && top === lastTop) return;
        ps_copyRow(f.ws, sr, ws, r, cols, f.cmap.map);
        const a = ps_own(ws.getRow(r).getCell(1)), b = ps_own(ws.getRow(r).getCell(2));
        placed.push({ r, kind: b ? 'item' : a ? 'head' : 'item', top: a.endsWith('主管') });
        if (a.endsWith('主管')) lastTop = a;
        r++;
    }); });
    putMissing(Infinity);
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
    ps_checkSheet(wb, ps_stackChecks());
    return wb.xlsx.writeBuffer();
}

// 逐列檢查：各檔每一列（候選列），勾「刪除」的不貼；顯示範本前三欄的內容
function ps_renderReview(data) {
    const $ = id => document.getElementById(id), heads = ps_base ? ps_heads(ps_base.ws) : new Map();
    const total = data.reduce((n, f) => n + f.cand.length, 0), del = data.reduce((n, f) => n + f.del.size, 0);
    $('ps-review').classList.toggle('hidden', !total);
    $('ps-review-title').textContent = `逐列檢查（共 ${total} 列，勾選刪除 ${del} 列）`;
    [1, 2, 3].forEach(c => { $('ps-review-c' + c).textContent = ps_label(heads.get(c) || '') || `第${c}欄`; });
    $('ps-review-list').innerHTML = data.map((f, fi) => f.cand.map(r => {
        const on = f.del.has(r), cell = c => f.cmap.map[c] ? ps_esc(ps_cellText(f.ws.getRow(r).getCell(f.cmap.map[c])).slice(0, 40)) : '';
        return `<tr${on ? ' class="text-slate-400 line-through"' : ''}><td><input type="checkbox" class="ps-del" data-f="${fi}" data-r="${r}"${on ? ' checked' : ''} aria-label="刪除 ${ps_esc(f.file)} 第 ${r} 列"></td>
            <td class="text-left whitespace-nowrap">${ps_esc(f.org)}</td><td class="text-left whitespace-nowrap max-w-[14rem] truncate" title="${ps_esc(f.file)}">${ps_esc(f.file)}</td><td class="text-left whitespace-nowrap">${cell(1)}</td><td class="text-left whitespace-nowrap">${cell(2)}</td><td class="text-left whitespace-nowrap">${cell(3)}</td></tr>`;
    }).join('')).join('');
}

function ps_render(notes = []) {
    const $ = id => document.getElementById(id);
    const n = ps_files.length, mode = ps_mode;
    const plan = ps_base && mode === 'fill' ? ps_fillPlan() : null;
    const count = f => plan ? plan.filled.get(f) : f.rows.length;
    const sheets = f => f.wb.worksheets.filter(w => w.state === 'visible' || w === f.ws)
        .map(w => `<option${w === f.ws ? ' selected' : ''}>${ps_esc(w.name)}</option>`).join('');
    $('ps-list').innerHTML = ps_files.map((f, i) => `<tr${f.tpl ? ' class="text-slate-500"' : ''}><td>${i + 1}</td>
        <td><input type="checkbox" class="ps-tpl" data-i="${i}"${f.tpl ? ' checked' : ''} aria-label="以 ${ps_esc(f.file)} 當範本"></td>
        <td class="text-left">${f.tpl ? '（範本）' : ps_esc(f.org)}</td><td class="text-left">${ps_esc(f.file)}</td>
        <td><select class="ps-sheet text-sm" data-i="${i}" aria-label="${ps_esc(f.file)} 的工作表">${sheets(f)}</select></td><td>${f.tpl ? '不貼入' : count(f)}</td>
        <td class="whitespace-nowrap"><button class="btn btn-sm" data-move="${i}" data-d="-1" aria-label="上移 ${ps_esc(f.org)}"${i ? '' : ' disabled'}>↑</button>
            <button class="btn btn-sm" data-move="${i}" data-d="1" aria-label="下移 ${ps_esc(f.org)}"${i < n - 1 ? '' : ' disabled'}>↓</button>
            <button class="btn btn-sm btn-danger" data-remove="${i}" aria-label="移除 ${ps_esc(f.org)}">移除</button></td></tr>`).join('');
    const data = ps_files.filter(f => !f.tpl);
    ps_renderReview(mode === 'stack' ? data : []);
    // 比對欄位：範本前三欄
    const heads = ps_base ? ps_heads(ps_base.ws) : new Map();
    $('ps-match').innerHTML = [1, 2, 3].map(c => `<option value="${c}"${c === ps_matchCol ? ' selected' : ''}>第${c}欄${heads.get(c) ? '（' + ps_esc(ps_label(heads.get(c)).slice(0, 12)) + '）' : ''}</option>`).join('');
    // 沒有同名工作表、改讀第一張的檔案提醒一下
    data.forEach(f => { if (!f.pinned && f.ws.name !== ps_base?.sheet) notes.push(`${f.file}：沒有「${ps_base.sheet}」工作表，改讀「${f.ws.name}」`); });
    $('ps-result').classList.toggle('hidden', !ps_files.length);
    $('ps-btn-xlsx').disabled = !data.length;
    $('ps-form').textContent = ps_base ? (ps_base.fromFile ? `依多數檔案共用的版面（「${ps_base.name}」），可在順序表勾選範本` : `範本「${ps_base.name}」・工作表「${ps_base.sheet}」`) : '拖入檔案後在順序表勾選一份當範本；沒勾選就用多數檔案共用的版面';
    const checks = plan ? plan.checks : ps_stackChecks();
    $('ps-count-head').textContent = mode === 'fill' ? '填了幾列' : '貼上列數';
    $('ps-check').classList.toggle('hidden', !checks.length);
    $('ps-check-title').textContent = `查對（${checks.length}）`;
    $('ps-check-list').innerHTML = checks.map(x => `<tr>${x.map((t, i) => `<td class="${i ? 'text-left' : 'text-left whitespace-nowrap'}">${ps_esc(t)}</td>`).join('')}</tr>`).join('');
    $('ps-msg').textContent = (data.length ? `已讀入 ${data.length} 份，${mode === 'fill' ? `依項目填入 ${plan.fills.length} 個項目` : `共 ${data.reduce((n, f) => n + f.rows.length, 0)} 列`}${checks.length ? `，查對 ${checks.length} 項` : ''}。` : '尚未匯入檔案。') + (notes.length ? ' ⚠️ ' + notes.join('；') : '');
}

// 檔案拖到頁面任何地方都收下：不讓瀏覽器把檔案打開或下載（放在最外層，初始化出錯也照樣擋住）
if (typeof window !== 'undefined') {
    window.addEventListener('dragover', e => e.preventDefault());
    window.addEventListener('drop', e => { e.preventDefault(); if (e.dataTransfer?.files.length) ps_addFiles([...e.dataTransfer.files]); });
}

document.addEventListener('DOMContentLoaded', async () => {
    const $ = id => document.getElementById(id);
    // 先綁拖放與選檔，後面讀標準名單出錯也能用
    const drop = $('ps-drop'), input = $('ps-file');
    drop.onclick = () => input.click();
    input.onchange = () => { ps_addFiles([...input.files]); input.value = ''; };
    drop.addEventListener('dragover', () => drop.classList.add('dragover'));
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    drop.addEventListener('drop', () => drop.classList.remove('dragover'));
    try {
        ps_orgs = await fetch(FB_ORGS, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
        ps_orgs.sortOrder ||= []; ps_orgs.aliasMap ||= {};
        ps_render();
    } catch (e) { $('ps-msg').textContent = e.message; }
    // 勾範本（只能一個）、改工作表、移動、移除都可能換版面，一律重算
    $('ps-list').addEventListener('change', async e => {
        const f = ps_files[+e.target.dataset.i];
        if (e.target.classList.contains('ps-tpl')) { const on = e.target.checked; ps_files.forEach(x => { x.tpl = false; }); f.tpl = on; }
        else if (e.target.classList.contains('ps-sheet')) f.pinned = e.target.value;
        else return;
        await ps_rebase();
        ps_render();
    });
    $('ps-list').addEventListener('click', async e => {
        const b = e.target.closest('button');
        if (!b) return;
        if (b.dataset.remove) ps_files.splice(+b.dataset.remove, 1);
        else { const i = +b.dataset.move, j = i + +b.dataset.d; [ps_files[i], ps_files[j]] = [ps_files[j], ps_files[i]]; }
        await ps_rebase();
        ps_render();
    });
    document.querySelectorAll('input[name="ps-mode"]').forEach(r => r.addEventListener('change', () => { ps_mode = r.value; ps_render(); }));
    // 逐列檢查：勾、取消刪除；回到預設；全部保留；換比對欄位重算預設
    const data = () => ps_files.filter(f => !f.tpl);
    $('ps-review-list').addEventListener('change', e => {
        if (!e.target.classList.contains('ps-del')) return;
        const f = data()[+e.target.dataset.f], r = +e.target.dataset.r;
        e.target.checked ? f.del.add(r) : f.del.delete(r);
        ps_applyDel(f);
        ps_render();
    });
    $('ps-btn-del-default').onclick = () => { data().forEach(f => { f.del = new Set(f.del0); ps_applyDel(f); }); ps_render(); };
    $('ps-btn-del-none').onclick = () => { data().forEach(f => { f.del = new Set(); ps_applyDel(f); }); ps_render(); };
    $('ps-match').onchange = async e => { ps_matchCol = +e.target.value; await ps_rebase(); ps_render(); };
    $('ps-btn-resort').onclick = () => { ps_resort(); ps_render(); };
    // 自訂順序清單：存在這個瀏覽器（localStorage，讀不到就算了）
    const box = $('ps-order'), state = () => { $('ps-order-state').textContent = ps_order ? `目前：自訂清單 ${ps_order.length} 項` : '目前：標準基金清單'; };
    try { box.value = localStorage.getItem('ps-order') || ''; } catch (e) { /* 無痕視窗等讀不到 */ }
    if (box.value.trim()) { ps_setOrder(box.value); $('ps-order-box').open = true; }
    state();
    $('ps-btn-order').onclick = () => { ps_setOrder(box.value); try { localStorage.setItem('ps-order', box.value); } catch (e) { /* 存不了就只用這次 */ } state(); ps_render(); };
    $('ps-btn-order-clear').onclick = () => { box.value = ''; ps_setOrder(''); try { localStorage.removeItem('ps-order'); } catch (e) { /* 同上 */ } state(); ps_render(); };
    $('ps-btn-clear').onclick = async () => { ps_files = []; await ps_rebase(); ps_render(); };
    $('ps-btn-xlsx').onclick = async () => {
        const buf = await ps_build();
        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        a.download = `${ps_base.name}-貼表.xlsx`;
        a.click();
        URL.revokeObjectURL(a.href);
    };
});
