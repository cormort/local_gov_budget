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

// ========== 1b. Excel 貼上：剪貼簿是 TSV（含換行或 tab 的儲存格以雙引號包住） ==========
function fb_parseTSV(text) {
    const rows = [[]];
    let cell = '', i = 0, quoted = false;
    for (; i < text.length; i++) {
        const ch = text[i];
        if (quoted) {
            if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
            else if (ch === '"') quoted = false;
            else cell += ch;
        } else if (ch === '"' && cell === '') quoted = true;
        else if (ch === '\t') { rows[rows.length - 1].push(cell); cell = ''; }
        else if (ch === '\n' || ch === '\r') {
            if (ch === '\r' && text[i + 1] === '\n') i++;
            rows[rows.length - 1].push(cell); cell = ''; rows.push([]);
        } else cell += ch;
    }
    rows[rows.length - 1].push(cell);
    if (rows.length > 1 && rows[rows.length - 1].join('') === '') rows.pop(); // Excel 複製結尾的換行
    return rows;
}

// "1,234" → 1234、"(500)" → -500、"－" 或空白 → ''；不是數字就原樣回傳讓使用者看到（不猜）
// 規則借鏡 multi-xlsx_merge 的 toNumber：
// - 全形數字、括號、負號、逗號、小數點、星號先轉半形（從 Word／網頁貼上的「（５００）」）；「、」當逗號；− – — ﹣ 當負號
// - 去掉單位 元、千元、百萬元、NT$、US$、$，以及結尾的附註星號「1,234*」
// - 千分位只移除「後面剛好三位數字」的逗號：「1,5」不會變成 15，而是當成不是數字讓使用者看到
// - 百分比、日期、科學記號一律不是數字
function fb_cleanNum(s) {
    let t = String(s ?? '').replace(/[（-．０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
        .replace(/[−–—﹣]/g, '-').replace(/、/g, ',').replace(/[\s　 ]/g, '')
        .replace(/NT\$|US\$|\$|百萬元|千元|元/gi, '');
    let prev;
    do { prev = t; t = t.replace(/(\d),(?=\d{3}(?:\D|$))/g, '$1'); } while (t !== prev);
    t = t.replace(/\*+$/, '');
    if (t === '' || t === '-') return '';
    let neg = false;
    if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
    if (t.startsWith('-')) { neg = !neg; t = t.slice(1); } else if (t.startsWith('+')) t = t.slice(1);
    if (!/^\d*\.?\d+$/.test(t)) return String(s ?? '').trim();
    const n = Number(t);
    return String(neg && n ? -n : n);
}

// 數值欄的值 → 數字（文字格式的 "1,234" 不能交給 parseFloat，會變成 1）；清不掉的當 0
const fb_num = v => { if (typeof v === 'number') return v; const n = Number(fb_cleanNum(v)); return isNaN(n) ? 0 : n; };
// 千分位、小數最多 4 位（與列值的 toFixed(4) 一致；toLocaleString 預設只到 3 位）
const fb_fmt = v => (+v.toFixed(4)).toLocaleString('zh-TW', { maximumFractionDigits: 4 });

// ========== 1b-2. 項目層級：依編號判斷，父項的數值欄自動 = 子項加總 ==========
// 一、=1 （一）=2 1.=3 （1）=4；沒有編號 = 0（獨立項目，不當父項也不當子項）
function fb_level(name) {
    const n = String(name ?? '').trim();
    if (/^[一二三四五六七八九十百]+[、.．]/.test(n)) return 1;
    if (/^[（(][一二三四五六七八九十百]+[）)]/.test(n)) return 2;
    if (/^\d+[.、．](?!\d)/.test(n)) return 3;
    if (/^[（(]\d+[）)]/.test(n)) return 4;
    return 0;
}

// 每列的深度：有編號照編號；沒編號的列是上方最近一個編號項目的下層（「一、作業基金」下的各基金），
// 前面有空白再深一層（「交通作業基金」下縮排的「  民航事業作業基金」）；合計、小計、註、說明列為 0（不參與加總層級）
// ponytail: 縮排只分「有／沒有」一層，多層縮排要再細分時改成依空白數
function fb_depths(names) {
    let last = 0;
    return names.map(n => {
        const s = String(n ?? ''), l = fb_level(s);
        if (l) return (last = l);
        if (!s.trim() || /^[\s\u3000]*(合計|總計|小計|註|備註|附註|說明)/.test(s)) return 0;
        return last + 1 + (/^[\s\u3000]/.test(s) ? 1 : 0);
    });
}

// 每列的父列索引（-1 = 最上層）：往上找最近一個深度較淺的列
function fb_parents(names) {
    const d = fb_depths(names);
    return d.map((l, i) => {
        if (!l) return -1;
        for (let j = i - 1; j >= 0; j--) if (d[j] && d[j] < l) return j;
        return -1;
    });
}

// 資料列（{ 欄代號: 值 }）的父項數值欄改為子項加總；給沒有畫面的彙整頁用（填表頁在 fb_calc 一起算）
function fb_rollup(rows, columns) {
    const nums = columns.filter(c => c.type === 'num' && !c.formula), name = columns[0].key;
    const parent = fb_parents(rows.map(r => r[name])), sums = rows.map(() => null);
    for (let i = rows.length - 1; i >= 0; i--) {
        if (sums[i]) nums.forEach(c => { rows[i][c.key] = +sums[i][c.key].toFixed(4); });
        const p = parent[i];
        if (p < 0) continue;
        sums[p] ??= Object.fromEntries(nums.map(c => [c.key, 0]));
        nums.forEach(c => { sums[p][c.key] += fb_num(rows[i][c.key]); });
    }
    return rows;
}

// 多工作表表單：template = { title, sheets: [{ name, columns, caption }] }，資料 tables = { 工作表名: rows }
// 舊格式（template.columns + rows，單一工作表）讀進來時轉成新格式，已發布的範本與回傳檔照樣能用
function fb_normalize(data) {
    const t = data.template || {};
    if (Array.isArray(t.sheets)) return { title: t.title, sheets: t.sheets, tables: data.tables || {} };
    if (!Array.isArray(t.columns)) throw new Error('不是自訂表單的 JSON（缺 template）');
    const name = t.sheet || t.title || '表單';
    const sheet = { name, columns: t.columns };
    if (t.caption) sheet.caption = t.caption;
    return { title: t.title, sheets: [sheet], tables: { [name]: data.rows || [] } };
}

// Excel 讀進來的列：數值欄的文字（"1,234"、"(300)"、全形）轉成數字，轉不了的（「無」）當空白；回傳警告（沒有就 null）
function fb_cleanRows(rows, columns) {
    const bad = [];
    rows.forEach(row => {
        columns.forEach(c => {
            if (c.type !== 'num' || c.formula || typeof row[c.key] !== 'string') return;
            const t = fb_cleanNum(row[c.key]);
            if (t !== '' && isNaN(Number(t))) { bad.push(`${row._r ? c.key + row._r : `「${row[columns[0].key] ?? ''}」${c.label}`}「${t}」`); row[c.key] = ''; } // JSON 沒有列號，改列項目與欄名
            else row[c.key] = t === '' ? '' : Number(t);
        });
        delete row._r;
    });
    return bad.length ? `${bad.length} 格不是數字，已當空白：${bad.slice(0, 5).join('、')}${bad.length > 5 ? ' 等' : ''}` : null;
}

// 數值欄（不含公式欄）有小數的格數：金額多半是整數，某份檔案出現小數常是千元與元混用（借鏡 multi-xlsx_merge 的 findDecimalFiles）
// sheets：[{ name, columns }]、tables：{ 工作表名: rows } → { count, sample }
function fb_decimalCells(sheets, tables) {
    let count = 0, sample = null;
    sheets.forEach(s => (tables[s.name] || []).forEach(r => s.columns.forEach(c => {
        if (c.type !== 'num' || c.formula || r[c.key] === '' || r[c.key] === undefined) return;
        const v = fb_num(r[c.key]);
        if (Math.abs(v - Math.round(v)) > 1e-9) { count++; sample ??= v; }
    })));
    return { count, sample };
}

// 年度預設：範本庫設定的年度 → 表單名稱開頭的年度（「114決算…」）→ 今年（民國）
function fb_guessYear(title, metaYear, now = new Date()) {
    if (String(metaYear ?? '').trim()) return String(metaYear).trim();
    const m = /^(\d{2,3})(?!\d)/.exec(String(title ?? '').trim());
    return m ? m[1] : String(now.getFullYear() - 1911);
}

// 列有沒有填數字（公式欄不算）；兩列的數值欄是否完全相同
const fb_hasData = (r, cols) => cols.some(c => c.type === 'num' && !c.formula && fb_num(r[c.key]));
const fb_sameNums = (a, b, cols) => cols.every(c => c.type !== 'num' || c.formula || fb_num(a[c.key]) === fb_num(b[c.key]));

// 彙整檢查（借鏡 excel_merge）：每個項目有哪幾份檔案填了數字，只看沒有下層的列（上層是加總出來的）
// subs: [{ rows }] → { owners: Map(項目 → [{ s, i }]), oneOwner }
// oneOwner：過半項目只有一份在填 → 「各填各的」清單型表（如各基金填自己那列），同項目多份即重複；
//           各機關填同一批項目的表（如經費表）每項都有多份，不算重複
function fb_itemOwners(subs, cols) {
    const key = cols[0].key, owners = new Map();
    subs.forEach((s, si) => {
        const parent = fb_parents(s.rows.map(r => r[key]));
        s.rows.forEach((r, i) => {
            const name = String(r[key] ?? '').trim();
            if (!name || parent.includes(i) || !fb_hasData(r, cols)) return;
            if (!owners.has(name)) owners.set(name, []);
            owners.get(name).push({ s: si, i });
        });
    });
    const filled = [...owners.values()];
    const oneOwner = subs.length > 1 && filled.length > 0 && filled.filter(o => o.length === 1).length / filled.length >= 0.5;
    return { owners, oneOwner };
}

// 公式是否參照了同一工作表的別列（例如合計列 =SUM(B3:B9)）；跨工作表或外部連結（=[1]大基金!$M$41）只是取值，不算
const fb_refsOtherRow = (f, row) => !f.includes('!') && [...f.toUpperCase().matchAll(/\$?[A-Z]+\$?(\d+)/g)].some(m => +m[1] !== row);

// 填好的 Excel 匯入前比對欄名，不一致就拒絕（回傳錯誤訊息，一致回傳 null）
function fb_headerMismatch(expected, got) {
    if (expected.length !== got.length) return `欄位數量不符：表單 ${expected.length} 欄，檔案 ${got.length} 欄\n表單：${expected.join('、')}\n檔案：${got.join('、')}`;
    const norm = s => String(s).replace(/\s+/g, ''); // 儲存格內換行（\r\n 或 \n）、空白不算差異
    const i = expected.findIndex((l, k) => norm(l) !== norm(got[k]));
    return i < 0 ? null : `第 ${i + 1} 欄應為「${expected[i]}」，檔案中為「${got[i]}」`;
}

// ========== 1c. 多列表頭（合併儲存格） ==========
// textAt(r, c) 取儲存格文字；merges 為 SheetJS 的 ws['!merges']（{ s:{r,c}, e:{r,c} }，0 起算）
const fb_mergeAt = (merges, r, c) => merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);

// 合併儲存格的文字只存在左上角，其他格視同左上角的文字
function fb_mergedText(textAt, merges, r, c) {
    const m = fb_mergeAt(merges, r, c);
    return m ? textAt(m.s.r, m.s.c) : textAt(r, c);
}

// 找表頭範圍 { hr, he }：
// - 跳過空白列，以及只有一格有字的列（表名「○○表」、「單位：千元」）
// - 表頭內有直向合併（如「計畫名稱」跨兩列）→ 表頭延伸到合併底端
// - 表頭內有橫向合併（如「預算數」跨三欄）且下一列在該範圍內有文字 → 下一列是子表頭
function fb_headerBlock(textAt, merges, r0, r1, c0, c1) {
    const filled = r => { let n = 0; for (let c = c0; c <= c1; c++) if (textAt(r, c)) n++; return n; };
    // 表名、單位列：空白、只有一格字、每格都是「標籤：值」（單位：千元、填表機關：），或第一欄橫向合併超過半寬（表名）
    const texts = r => { const t = []; for (let c = c0; c <= c1; c++) if (textAt(r, c)) t.push(textAt(r, c)); return t; };
    const isTitle = r => merges.some(m => m.s.r === r && m.s.c === c0 && (m.e.c - m.s.c + 1) * 2 > c1 - c0 + 1);
    // 只有一格字、但那格是不從第一欄開始的橫向合併，且下一列在它下面有字（項目欄空著、B1:C1 合併「金額」）→ 上層表頭，不是表名
    const isGroup = r => merges.some(m => m.s.r === r && m.s.c > c0 && m.e.c > m.s.c && textAt(r, m.s.c) && filled(r + 1) >= 2 &&
        [...Array(m.e.c - m.s.c + 1)].some((_, k) => textAt(r + 1, m.s.c + k)));
    const isCaption = r => { const t = texts(r); return !t.length || (c1 > c0 && t.length === 1 && !isGroup(r)) || t.every(x => /[：:]/.test(x)) || isTitle(r); };
    let hr = r0;
    while (hr <= r1 && isCaption(hr)) hr++; // ponytail: 單欄範本的表頭若只有一格會被當表名跳過
    if (hr > r1) return null;
    let he = hr, changed = true;
    while (changed) {
        changed = false;
        merges.filter(m => m.s.r >= hr && m.s.r <= he).forEach(m => {
            let end = m.e.r;
            if (m.e.c > m.s.c && m.e.r + 1 <= r1) {
                const below = m.e.r + 1;
                for (let c = m.s.c; c <= m.e.c; c++) {
                    const t = textAt(below, c);
                    if (t && isNaN(Number(t.replace(/,/g, '')))) { end = below; break; }
                }
            }
            if (end > he) { he = end; changed = true; }
        });
    }
    return { hr, he };
}

// 某欄由上而下的表頭文字，去掉合併造成的連續重複：["預算數","原編"]、["計畫名稱"]
function fb_columnLevels(textAt, merges, hr, he, c) {
    const levels = [];
    for (let r = hr; r <= he; r++) {
        const t = fb_mergedText(textAt, merges, r, c);
        if (t && t !== levels[levels.length - 1]) levels.push(t);
    }
    return levels;
}

// 依各欄 levels 排出多列 <thead>：同一上層連續的欄合成 colspan，層數較少的欄用 rowspan 補滿
function fb_theadRows(columns) {
    const L = columns.map(c => c.levels || [c.label]);
    const D = Math.max(1, ...L.map(l => l.length));
    const same = (a, b, k) => a.slice(0, k + 1).join('\u0000') === b.slice(0, k + 1).join('\u0000');
    const rows = [];
    for (let k = 0; k < D; k++) {
        const row = [];
        for (let j = 0; j < L.length;) {
            const lv = L[j];
            if (lv.length <= k) { j++; continue; }
            if (k === lv.length - 1) { row.push({ text: lv[k], colspan: 1, rowspan: D - k, col: columns[j] }); j++; continue; }
            let span = 1;
            while (j + span < L.length && L[j + span].length > k + 1 && same(L[j + span], lv, k)) span++;
            row.push({ text: lv[k], colspan: span, rowspan: 1 });
            j += span;
        }
        rows.push(row);
    }
    return rows;
}

// headers 可混用字串與 template 欄位（有 levels 的欄位會排成多層表頭，與填報表單相同）
// 回傳 thead 的列（給 HTML colspan/rowspan）以及等價的二維陣列＋合併範圍（給 Excel）
function fb_headGrid(headers) {
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

// 「前往 GitHub 發布」與分享連結跟著網站所在的帳號／repo 走，別人複製這個專案後不必改程式
// alice.github.io/local_gov_budget/… → github.com/alice/local_gov_budget；本機開發時用原作者的 repo
function fb_repoUrls(hostname, pathname) {
    const m = /^([^.]+)\.github\.io$/i.exec(hostname);
    if (!m) return { upload: 'https://github.com/cormort/local_gov_budget/upload/main/forms', site: 'https://cormort.github.io/local_gov_budget/', api: 'https://api.github.com/repos/cormort/local_gov_budget/contents/forms' };
    const owner = m[1].toLowerCase();
    const first = pathname.split('/')[1] || '';
    const isProjectSite = first && !first.includes('.');            // /repo/… 是專案網站；/xxx.html 是 owner.github.io 本身
    const repo = isProjectSite ? first : hostname.toLowerCase();
    return {
        upload: `https://github.com/${owner}/${repo}/upload/main/forms`,
        site: `https://${hostname.toLowerCase()}/${isProjectSite ? repo + '/' : ''}`,
        api: `https://api.github.com/repos/${owner}/${repo}/contents/forms`,
    };
}

// 範本庫的說明資料：forms/index.json = { templates: { 表單名稱: { category, unit, purpose, description, year, start, deadline, hidden } } }
// （分類、發起單位、用途、說明、年度、發起日期、截止日、隱藏）
// 由管理頁（manage.html）維護；沒有寫到的表單照樣列出，歸在「未分類」
const FB_INDEX = 'index.json';
const FB_NO_CATEGORY = '未分類';

// 依分類分組（分類照名稱排序、未分類最後），預設不列隱藏的範本
function fb_libraryGroups(entries, index = {}, showHidden = false) {
    const meta = index.templates || {}, groups = new Map();
    entries.forEach(e => {
        const m = meta[e.name] || {};
        if (m.hidden && !showHidden) return;
        const cat = (m.category || '').trim() || FB_NO_CATEGORY;
        if (!groups.has(cat)) groups.set(cat, []);
        groups.get(cat).push({ ...e, meta: m });
    });
    return [...groups].map(([category, items]) => ({ category, items }))
        .sort((a, b) => (a.category === FB_NO_CATEGORY) - (b.category === FB_NO_CATEGORY) || a.category.localeCompare(b.category, 'zh-Hant'));
}

// 調查專案的狀態（依發起日期 start、截止日 deadline，YYYY-MM-DD）：未開始／進行中／已結束；兩個都沒填就是空字串
const FB_STATUS = { upcoming: '未開始', open: '進行中', closed: '已結束' };
const FB_BADGE = { upcoming: 'text-slate-600', open: 'text-green-700', closed: 'text-red-600' };
// 今天的日期 YYYY-MM-DD，用電腦的本地時間（toISOString 是 UTC，台灣凌晨 0～8 點會算成前一天）
const fb_today = (d = new Date()) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
function fb_libraryStatus(m = {}, today = fb_today()) {
    if (m.start && m.start > today) return 'upcoming';
    if (m.deadline && m.deadline < today) return 'closed';
    return m.start || m.deadline ? 'open' : '';
}

// 範本庫／調查專案查詢：關鍵字比對名稱、專案名稱、範本、說明、用途、發起單位、分類、年度
// （空白分隔的多個字都要符合，不分大小寫）；cat 指定分類，status 指定調查狀態
function fb_libraryMatch(item, q = '', cat = '', status = '', today) {
    const m = item.meta || {};
    if (cat && ((m.category || '').trim() || FB_NO_CATEGORY) !== cat) return false;
    if (status && fb_libraryStatus(m, today) !== status) return false;
    const hay = [item.name, m.title, m.form, m.description, m.purpose, m.unit, m.category, m.year].filter(Boolean).join(' ').toLowerCase();
    return q.trim().toLowerCase().split(/\s+/).filter(Boolean).every(w => hay.includes(w));
}

// ========== 調查專案：surveys/index.json = { surveys: { 代號: { title, form, unit, purpose, year, start, deadline, respondents, hidden } } } ==========
// 代號用在網址 ?survey=代號，也當回傳檔的專案代號
const FB_SURVEY_INDEX = 'surveys/index.json';

// 範本名稱、專案代號能不能用（會出現在網址與檔名）：回傳錯誤訊息，可以用回傳 null
function fb_badName(name) {
    const n = String(name ?? '');
    if (!n.trim()) return '名稱不能空白';
    if (/[\/\\?#%&]|\.\.|^\.|\s/.test(n)) return '名稱不能有空白或 / \\ ? # % & 等符號';
    if (n === 'index') return '「index」保留給系統使用';
    return null;
}

// 專案清單 → [{ name: 代號, meta }]（與範本庫同形，可共用查詢）；依發起日期新的在前
function fb_surveyEntries(index = {}, showHidden = false) {
    return Object.entries(index.surveys || {}).filter(([, m]) => showHidden || !m.hidden)
        .map(([name, meta]) => ({ name, meta }))
        .sort((a, b) => (b.meta.start || '').localeCompare(a.meta.start || '') || a.name.localeCompare(b.name, 'zh-Hant'));
}

// 機關名稱比對：去空白、台→臺、去括號底線（名單與回傳檔常寫法不一）
const fb_orgKey = s => String(s ?? '').replace(/[\s　_()（）]/g, '').replace(/台/g, '臺');

// 標準機關名單：surveys/orgs.json = { sortOrder: [標準名稱…], aliasMap: { 簡稱: 標準名稱 } }（沿用 excel_merge 的 fund-config.json）
const FB_ORGS = 'surveys/orgs.json';
const FB_ORG_SUFFIX = /(作業基金|校務基金|基金)$/;

// 回傳檔的機關名稱 → 應繳名單的第幾個（借鏡 excel_merge 的 cellMatchesFund）。依序：
// 1 完全相同（或經別名表換成同一個標準名稱）、2 去掉「作業基金／校務基金／基金」後相同、
// 3 前 4 字相同且字數相同、4 名單名稱（至少 4 字）包含在回傳名稱裡、取最長的。2～4 只在恰好一個候選時才算
// 回傳 { i, level }，找不到 { i: -1 }
// 兩個名稱是否只差一個字（替換、多一字或少一字）；用來容許打錯字
function fb_oneEditApart(a, b) {
    const x = [...a], y = [...b];
    if (Math.abs(x.length - y.length) > 1) return false;
    let i = 0, j = 0, diff = 0;
    while (i < x.length && j < y.length) {
        if (x[i] === y[j]) { i++; j++; continue; }
        if (++diff > 1) return false;
        if (x.length > y.length) i++; else if (y.length > x.length) j++; else { i++; j++; }
    }
    return diff + (x.length - i) + (y.length - j) <= 1;
}

// known：標準機關名單（surveys/orgs.json 的 sortOrder）。交件機關本身就是另一個標準機關時，不再模糊配到名單上的別人
// （名單只列部分基金時，「國立臺灣師範大學校務基金」不能算成「國立臺灣海洋大學校務基金」）
function fb_orgMatcher(expected = [], aliasMap = {}, known = []) {
    const alias = new Map(Object.entries(aliasMap).map(([k, v]) => [fb_orgKey(k), fb_orgKey(v)]));
    const canon = s => { const k = fb_orgKey(s); return alias.get(k) || k; };
    const exp = expected.map(canon), strip = k => k.replace(FB_ORG_SUFFIX, '');
    const knownSet = new Set(known.map(canon));
    const one = list => list.length === 1 ? list[0] : -1;
    return org => {
        const k = canon(org);
        if (!k) return { i: -1 };
        let i = exp.indexOf(k);
        if (i >= 0) return { i, level: 1 };
        if (knownSet.has(k)) return { i: -1 }; // 是另一個標準機關，不在名單上
        const sk = strip(k);
        if (sk.length >= 2 && (i = one(exp.flatMap((e, j) => strip(e) === sk ? [j] : []))) >= 0) return { i, level: 2 };
        // 打錯一個字（6 字以上才容許，「臺北市／臺南市」這類短名稱不模糊比對）
        if ([...k].length >= 6 && (i = one(exp.flatMap((e, j) => fb_oneEditApart(e, k) ? [j] : []))) >= 0) return { i, level: 3 };
        const inside = exp.map((e, j) => ({ e, j })).filter(x => [...x.e].length >= 4 && k.includes(x.e));
        const longest = Math.max(0, ...inside.map(x => x.e.length));
        if ((i = one(inside.filter(x => x.e.length === longest).map(x => x.j))) >= 0) return { i, level: 4 };
        return { i: -1 };
    };
}

// 從檔名找機關（借鏡 excel_merge 的 detectFundsInText）：標準名稱、別名、去掉「作業基金／校務基金／基金」的短名，
// 由長到短比對，命中的字段遮蔽，避免短名稱重複命中；回傳找到的標準名稱（依 names 順序）
function fb_detectOrgs(text, names = [], aliasMap = {}) {
    const keys = new Map();
    const add = (key, name) => { if (key && [...key].length >= 2 && !keys.has(key)) keys.set(key, name); };
    names.forEach(n => add(n, n));
    Object.entries(aliasMap).forEach(([a, n]) => { if (names.includes(n)) add(a, n); });
    names.forEach(n => { const s = n.replace(FB_ORG_SUFFIX, ''); if ([...s].length >= 3 && s !== n) add(s, n); });
    let scan = String(text ?? '');
    const found = new Set();
    [...keys].sort((a, b) => b[0].length - a[0].length).forEach(([key, name]) => {
        if (!scan.includes(key)) return;
        found.add(name);
        scan = scan.split(key).join('\u0000'.repeat(key.length));
    });
    return [...new Set(names)].filter(n => found.has(n)); // 名單裡重複的名稱（orgs.json 有幾個基金列了兩次）只算一個
}

// 繳交對照：應繳名單 × 已收到的機關 → 每個機關 done（已繳）／missing（未繳）／extra（不在名單）
// 比對較寬鬆的（level > 1）在 rows 帶 fuzzy: true；先配完全相同的，模糊比對不會搶走別人的名額；同一機關兩份，後面那份算不在名單
function fb_submissions(expected = [], orgs = [], aliasMap = {}, known = []) {
    const match = fb_orgMatcher(expected, aliasMap, known);
    const found = orgs.map((o, idx) => ({ idx, ...match(o) })).filter(m => m.i >= 0).sort((a, b) => a.level - b.level || a.idx - b.idx);
    const taken = new Map(), used = new Set();
    found.forEach(m => { if (!taken.has(m.i)) { taken.set(m.i, m); used.add(m.idx); } });
    const rows = expected.map((o, i) => {
        const m = taken.get(i);
        return m ? { org: o, status: 'done', index: m.idx, ...(m.level > 1 ? { fuzzy: true } : {}) } : { org: o, status: 'missing' };
    });
    orgs.forEach((o, i) => { if (!used.has(i)) rows.push({ org: o, status: 'extra', index: i }); });
    return { rows, done: rows.filter(r => r.status === 'done').length, total: expected.length };
}

// 範本庫：GitHub API 列出的 forms/ 檔案 → [{ name, file }]；同名的 .json 與 .xlsx 只列一次（.json 為準，與 ?form= 相同）
function fb_libraryEntries(items) {
    const byName = new Map();
    items.filter(x => x.type === 'file' && /\.(json|xlsx)$/i.test(x.name) && x.name !== FB_INDEX).forEach(x => {
        const name = x.name.replace(/\.[^.]+$/, '');
        if (!byName.has(name) || /\.json$/i.test(x.name)) byName.set(name, x.name);
    });
    // .json 與 .xlsx 同名：表單以 .json 為準，.xlsx 是給填報人的空白 Excel，另外提供下載
    const xlsx = new Set(items.filter(x => /\.xlsx$/i.test(x.name)).map(x => x.name));
    return [...byName].map(([name, file]) => ({ name, file, ...(file !== name + '.xlsx' && xlsx.has(name + '.xlsx') ? { xlsx: name + '.xlsx' } : {}) })).sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));
}

// GitHub API 讀不到時的備援：forms/index.json 有登記的表單（不知道副檔名，file 為 null，不提供範本檔下載）
const fb_libraryFromIndex = index => Object.keys(index.templates || {}).map(name => ({ name, file: null })).sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));

if (typeof module !== 'undefined') module.exports = { fb_repoUrls, fb_translateFormula, fb_evalFormula, fb_parseTSV, fb_cleanNum, fb_num, fb_headerBlock, fb_columnLevels, fb_theadRows, fb_level, fb_depths, fb_parents, fb_refsOtherRow, fb_headerMismatch, fb_rollup, fb_cleanRows, fb_decimalCells, fb_guessYear, fb_hasData, fb_sameNums, fb_itemOwners, fb_normalize, fb_libraryEntries, fb_libraryFromIndex, fb_libraryGroups, fb_libraryMatch, fb_today, fb_libraryStatus, fb_badName, fb_surveyEntries, fb_orgKey, fb_oneEditApart, fb_orgMatcher, fb_detectOrgs, fb_submissions, fb_headGrid };
if (typeof document === 'undefined') { /* node 測試用 */ } else {

// ========== 2. 讀 Excel 範本 ==========
let fb_template = null; // 目前分頁的工作表 { title, sheet, columns, caption }
let fb_form = null;     // 整份表單 { title, sheets, tables, cur }；其他分頁的列存在 tables，換分頁時才畫到畫面上

// 工作表：預設讀第一個沒隱藏的；sheet 可指定名稱。buf 可以是檔案內容或已讀好的 workbook（換工作表時不必重讀）
// 用 function 宣告（不用 const）：這段在 if 區塊裡，彙整頁要能當全域函數呼叫
// cellStyles：SheetJS 要開這個才讀得到隱藏的欄、列（!cols / !rows 的 hidden）
function fb_openWorkbook(buf) { return buf.SheetNames ? buf : XLSX.read(buf, { type: 'array', cellFormula: true, sheetStubs: true, cellStyles: true }); }
function fb_visibleSheets(wb) { return wb.SheetNames.filter((n, i) => !wb.Workbook?.Sheets?.[i]?.Hidden); }

// formCols：匯入填好的檔案時傳入表單的欄，數值欄依表單判斷（檔案整欄存成文字時會被判成文字欄）
function fb_readWorkbook(buf, fileName, sheet, formCols, keepHiddenCols = false) {
    const wb = fb_openWorkbook(buf);
    if (sheet && !wb.Sheets[sheet]) throw new Error(`找不到工作表「${sheet}」`);
    sheet = sheet || fb_visibleSheets(wb)[0] || wb.SheetNames[0];
    const ws = wb.Sheets[sheet];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const cell = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
    const text = (r, c) => { const x = cell(r, c); return x ? String(x.w ?? x.v ?? '').trim() : ''; };

    // 表頭範圍：可能跨多列（合併儲存格），欄名以「－」串接各層，例如「預算數－原編」
    const merges = ws['!merges'] || [];
    const block = fb_headerBlock(text, merges, range.s.r, range.e.r, range.s.c, range.e.c);
    if (!block) throw new Error('找不到表頭列');
    const { hr, he } = block;

    // 隱藏的欄、列（輔助計算、往年資料）不當成表單的欄與項目（借鏡 excel_merge）
    // 欄：範本與回傳檔一致略過（回傳檔沿用原 Excel 時欄位才對得上）；
    // 列：只有範本略過，讀填好的檔案（formCols）時照讀，填表人隱藏了有數字的列也不會默默少算
    // keepHiddenCols：回傳檔的隱藏欄可能是填表人自己藏的（範本沒藏），比對表單欄名時要能改成照讀
    const hiddenCol = c => !keepHiddenCols && ws['!cols']?.[c]?.hidden, hiddenRow = r => !formCols && ws['!rows']?.[r]?.hidden;
    const columns = [], warnings = [], hiddenCols = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const levels = fb_columnLevels(text, merges, hr, he, c);
        if (!levels.length) continue;
        if (hiddenCol(c)) { hiddenCols.push(levels.join('－')); continue; }
        const label = levels.join('－');
        const key = XLSX.utils.encode_col(c);
        const col = { key, label, type: columns.length === 0 ? 'text' : 'num' };
        if (levels.length > 1) col.levels = levels;
        if (columns.length > 0) {
            // 取表頭下方第一個同列公式當該欄公式；跨列公式（合計、小計列）不算，改由層級小計與表尾合計計算
            let unsupported = null, nText = 0, nNum = 0;
            for (let r = he + 1; r <= range.e.r; r++) {
                const x = cell(r, c);
                if (x && x.f) {
                    const expr = fb_translateFormula(x.f, r + 1);
                    if (expr) { col.formula = expr; break; }
                    if (!fb_refsOtherRow(x.f, r + 1)) unsupported ??= x.f;
                }
                if (x && x.t === 'n') nNum++;
                else if (x && x.t === 's' && x.v.trim()) isNaN(Number(fb_cleanNum(x.v))) ? nText++ : nNum++;
            }
            // 文字比數字多才是文字欄：文字格式的「1,234」算數字，夾一兩格「無」的數值欄仍是數值欄
            if (nText > nNum) col.type = 'text';
            if (!col.formula && unsupported) warnings.push(`「${label}」公式 =${unsupported} 不支援，改為手填`);
        }
        columns.push(col);
    }
    if (columns.length < 2) throw new Error(`這個工作表看起來不像表格：只找到 ${columns.length} 欄表頭（需要項目欄與至少一個數值欄）`);
    // 公式只能參照範本內的欄
    const keys = new Set(columns.map(c => c.key));
    columns.forEach(c => {
        if (c.formula && (c.formula.match(/[A-Z]+/g) || []).some(k => !keys.has(k))) {
            warnings.push(`「${c.label}」公式參照了沒有表頭的欄，改為手填`);
            delete c.formula;
        }
    });

    // 資料列 → 預設項目（第一欄有值才算）
    // 沒有編號又用跨列公式加總的列（合計、小計）略過：表單由表尾合計與層級小計自動算，留著會重複加總
    const rows = [], skipped = [], hiddenRows = [];
    for (let r = he + 1; r <= range.e.r; r++) {
        const name = text(r, XLSX.utils.decode_col(columns[0].key));
        if (!name) continue;
        if (hiddenRow(r)) { hiddenRows.push(name); continue; }
        if (!fb_level(name) && columns.some(c => { const x = cell(r, XLSX.utils.decode_col(c.key)); return x?.f && fb_refsOtherRow(x.f, r + 1); })) { skipped.push(name); continue; }
        const row = {};
        columns.forEach(c => {
            if (c.formula) return;
            const x = cell(r, XLSX.utils.decode_col(c.key));
            // 錯誤值（#DIV/0! 等，SheetJS 存成 t:'e'、v 為錯誤代碼）當空白，不然空白列會被當成有填數字
            row[c.key] = x && x.t !== 'e' ? (typeof x.v === 'number' ? x.v : String(x.v ?? '')) : '';
        });
        row._r = r + 1; // Excel 列號，給 fb_cleanRows 的警告用
        rows.push(row);
    }
    if (skipped.length) warnings.push(`略過以公式加總的「${skipped.join('、')}」列（表單會自動計算合計）`);
    if (hiddenCols.length) warnings.push(`略過隱藏的欄「${hiddenCols.join('、')}」`);
    if (hiddenRows.length) warnings.push(`略過隱藏的列「${hiddenRows.join('、')}」`);
    const bad = fb_cleanRows(rows, formCols ? columns.map((c, i) => ({ ...c, type: formCols[i]?.type ?? c.type })) : columns);
    if (bad) warnings.push(bad);

    // 表頭上方的表名、「單位：千元」等文字列 → 顯示在表格上方
    const caption = [];
    for (let r = range.s.r; r < hr; r++) {
        const line = [];
        for (let c = range.s.c; c <= range.e.c; c++) if (text(r, c)) line.push(text(r, c));
        if (line.length) caption.push(line.join(' '));
    }

    const template = { title: fileName.replace(/\.[^.]+$/, ''), sheet, columns };
    if (caption.length) template.caption = caption;
    // 由填表頁「匯出 Excel」產生的檔案，第二張工作表帶機關、年度、填表人
    const info = wb.Sheets[FB_INFO_SHEET], metadata = {};
    if (info) XLSX.utils.sheet_to_json(info, { header: 1 }).forEach(([k, v]) => { if (FB_INFO_KEYS[k]) metadata[FB_INFO_KEYS[k]] = String(v ?? '').trim(); });
    return { template, rows, warnings, metadata, sheet, sheets: fb_visibleSheets(wb) };
}

const fb_samePrefix = (a, b, n) => [...a].slice(0, n).join('') === [...b].slice(0, n).join('') && [...a].length >= n;

// 已填的 Excel（填表頁匯入、彙整頁）：表單的每個工作表，先找同名工作表；沒有同名的，再從其他工作表找欄名一致的
// （排除表單其他工作表的名稱與已被認領的，免得兩張表頭相同的表互搶）。回傳 { found: { 表單工作表: 解析結果 }, notes, metadata }
function fb_matchSheets(buf, fileName, sheets) {
    const wb = fb_openWorkbook(buf), formNames = new Set(sheets.map(s => s.name)), used = new Set();
    const found = {}, notes = [];
    let metadata = {};
    sheets.forEach(s => {
        const labels = s.columns.map(c => c.label);
        const same = wb.Sheets[s.name] ? [s.name] : [];
        // 改過名的工作表要名稱相近（開頭 4 字相同，如「114非營業(環保基金)」）或只剩這一張；
        // 否則欄名相同的「113非營業(供參)」這類往年資料會被誤認
        const pool = fb_visibleSheets(wb).filter(n => !formNames.has(n) && !used.has(n) && n !== FB_INFO_SHEET);
        const others = pool.length === 1 ? pool : pool.filter(n => fb_samePrefix(n, s.name, 4));
        for (const n of same.length ? same : others) {
            // 先照範本規則略過隱藏欄；欄名對不上再連隱藏欄一起讀（填表人自己藏了某欄時，整份才不會被擋掉）
            let p, bad = '找不到表頭列';
            for (const keep of [false, true]) {
                try { p = fb_readWorkbook(wb, fileName, n, s.columns, keep); } catch (e) { p = null; continue; }
                bad = fb_headerMismatch(labels, p.template.columns.map(c => c.label));
                if (!bad) break;
            }
            if (!p) continue;
            if (!bad) { found[s.name] = p; used.add(n); metadata = p.metadata; return; }
            if (same.length) { notes.push(`工作表「${n}」${bad.split('\n')[0]}`); return; }
        }
        notes.push(`找不到工作表「${s.name}」（也沒有欄名一致的工作表）`);
    });
    return { found, notes, metadata };
}

// ========== 2b. 多工作表：分頁 ==========
// 目前分頁的列存回 fb_form（換分頁、匯出前）
function fb_saveSheet() {
    if (fb_form && fb_form.cur !== null) fb_form.tables[fb_form.sheets[fb_form.cur].name] = fb_collectRows();
}

function fb_showSheet(i) {
    fb_saveSheet();
    fb_form.cur = i;
    const s = fb_form.sheets[i];
    fb_render({ title: fb_form.title, sheet: s.name, columns: s.columns, caption: s.caption }, fb_form.tables[s.name] || []);
    document.querySelectorAll('#fb-tabs .fb-tab').forEach((b, k) => b.classList.toggle('fb-tab-on', k === i));
}

function fb_renderForm(form, metaYear) {
    fb_form = { title: form.title, sheets: form.sheets, tables: { ...form.tables }, cur: null };
    const year = document.getElementById('fb-year');
    if (!year.value) year.value = fb_guessYear(form.title, metaYear);
    document.getElementById('fb-library').classList.add('hidden');
    document.getElementById('fb-hub').classList.add('hidden');
    document.getElementById('fb-work').classList.remove('hidden');
    document.querySelectorAll('.fb-needs-form').forEach(b => b.classList.remove('hidden'));
    document.getElementById('fb-publish').classList.add('hidden');
    document.getElementById('fb-tabs').innerHTML = form.sheets.length > 1
        ? form.sheets.map((s, i) => `<button class="fb-tab" data-i="${i}">${fb_esc(s.name)}</button>`).join('') : '';
    fb_showSheet(0);
}

// fb_readWorkbook 的結果（可多個工作表）→ 表單
function fb_formFromParsed(title, list) {
    return {
        title,
        sheets: list.map(p => ({ name: p.sheet, columns: p.template.columns, ...(p.template.caption ? { caption: p.template.caption } : {}) })),
        tables: Object.fromEntries(list.map(p => [p.sheet, p.rows])),
    };
}

// ========== 3. 畫表 + 計算 ==========
const fb_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

function fb_render(template, rows) {
    fb_template = template;
    document.getElementById('fb-title').value = template.title;
    document.getElementById('fb-heading').textContent = template.title;
    document.getElementById('fb-caption').textContent = (template.caption || []).join('\n');
    document.getElementById('fb-btn-fill-xlsx').classList.remove('hidden');
    const head = fb_theadRows(template.columns);
    document.getElementById('fb-thead').innerHTML = head.map((row, k) =>
        '<tr>' + (k === 0 ? `<th class="w-10" rowspan="${head.length}"></th>` : '') +
        row.map(h => `<th scope="${h.colspan > 1 ? 'colgroup' : 'col'}" colspan="${h.colspan}" rowspan="${h.rowspan}">${fb_esc(h.text)}${h.col?.formula ? '<span class="font-normal text-slate-500">（計算）</span>' : ''}</th>`).join('') +
        '</tr>').join('');
    document.getElementById('fb-tfoot').innerHTML =
        '<tr><td></td>' + template.columns.map((c, i) =>
            i === 0 ? '<td class="text-center">合計</td>' : `<td>${c.type === 'num' ? `<input type="text" class="t-${c.key}" aria-label="${fb_esc(c.label)} 合計" readonly>` : ''}</td>`).join('') + '</tr>';
    const tbody = document.getElementById('fb-tbody');
    tbody.innerHTML = '';
    (rows.length ? rows : [{}]).forEach(fb_addRow);
    document.getElementById('fb-card').classList.remove('hidden');
    fb_calc();
}

function fb_addRow(data = {}) {
    const tr = document.createElement('tr');
    tr.innerHTML = '<td class="text-center"><button class="fb-del" title="刪除此列">×</button></td>' +
        fb_template.columns.map((c, i) => {
            const type = c.type === 'num' ? 'number' : 'text';
            const cls = `v-${c.key}${i === 0 || c.type === 'text' ? ' text-left' : ''}`;
            return `<td><input type="${type}" class="${cls}" aria-label="${fb_esc(c.label)}" value="${fb_esc(data[c.key])}"${c.formula ? ' readonly tabindex="-1"' : ''}></td>`;
        }).join('');
    document.getElementById('fb-tbody').appendChild(tr);
}

// 從游標所在儲存格開始，依 Excel 的欄位順序往右、往下填；公式欄跳過（照樣佔一欄，與 Excel 版面一致），列不夠就新增
function fb_onPaste(e) {
    const input = e.target.closest('input');
    const text = e.clipboardData?.getData('text/plain') || '';
    if (!input) return;
    if (!/[\t\n\r]/.test(text)) {
        // 單一儲存格：文字欄交給瀏覽器；數字欄先清理，不然「1,234」「（５００）」會被 type=number 丟掉
        if (input.type !== 'number') return;
        e.preventDefault();
        const t = fb_cleanNum(text);
        input.value = t;
        if (t !== '' && isNaN(Number(t))) fb_msg(`⚠️ 貼上的「${t}」不是數字，已當空白`);
        return fb_calc();
    }
    e.preventDefault();
    const cols = fb_template.columns;
    const startCol = cols.findIndex(c => input.classList.contains('v-' + c.key));
    let tr = input.closest('tr');
    const bad = [];
    fb_parseTSV(text).forEach((cells, r) => {
        if (r > 0) {
            if (!tr.nextElementSibling) fb_addRow();
            tr = tr.nextElementSibling;
        }
        cells.forEach((v, k) => {
            const c = cols[startCol + k];
            if (!c || c.formula) return;
            const t = c.type === 'num' ? fb_cleanNum(v) : v.trim();
            if (c.type === 'num' && t !== '' && isNaN(Number(t))) bad.push(`「${t}」`);
            tr.querySelector('.v-' + c.key).value = t; // 不是數字的值 type=number 會丟掉，所以上面先記下來
        });
    });
    if (bad.length) fb_msg(`⚠️ 貼上的 ${bad.length} 格不是數字，已當空白：${bad.slice(0, 5).join('、')}${bad.length > 5 ? ' 等' : ''}`);
    fb_calc();
}

function fb_calc() {
    if (!fb_template) return;
    const cols = fb_template.columns, totals = {};
    const trs = [...document.querySelectorAll('#fb-tbody tr')];
    const names = trs.map(tr => tr.querySelector('.v-' + cols[0].key).value);
    const parent = fb_parents(names), depth = fb_depths(names);
    const sums = trs.map(() => ({}));
    // 由下往上：子列一定在父列下方，算到父列時子列已加進 sums；合計只加最上層的列，避免重複加總
    for (let i = trs.length - 1; i >= 0; i--) {
        const tr = trs[i], el = k => tr.querySelector('.v-' + k);
        const get = k => parseFloat(el(k)?.value) || 0;
        const isParent = parent.includes(i);
        // 項目名稱依層級縮排，上層項目粗體
        const nameEl = el(cols[0].key);
        nameEl.style.paddingLeft = `${Math.max(depth[i] - 1, 0) * 1.5 + 0.25}rem`;
        nameEl.style.fontWeight = isParent ? '700' : '';
        cols.forEach(c => {
            if (c.type !== 'num' || c.formula) return;
            el(c.key).readOnly = isParent;
            if (isParent) el(c.key).value = +(sums[i][c.key] || 0).toFixed(4);
        });
        // 依欄序重算數次，讓公式欄可以參照其他公式欄；父列以加總後的值重算（比率類才正確）
        for (let pass = 0; pass < cols.length; pass++) {
            cols.forEach(c => { if (c.formula) el(c.key).value = +fb_evalFormula(c.formula, get).toFixed(4); });
        }
        const into = parent[i] >= 0 ? sums[parent[i]] : totals;
        cols.forEach(c => {
            if (c.type !== 'num') return;
            const v = get(c.key);
            el(c.key).classList.toggle('negative-value', v < 0);
            into[c.key] = (into[c.key] || 0) + v;
        });
    }
    // 公式欄的合計以數值欄的合計重算（比率類才正確，與父列、彙整頁相同）
    for (let pass = 0; pass < cols.length; pass++) {
        cols.forEach(c => { if (c.formula) totals[c.key] = fb_evalFormula(c.formula, k => totals[k] || 0); });
    }
    cols.forEach(c => {
        const el = document.querySelector('#fb-tfoot .t-' + c.key), v = totals[c.key] || 0;
        if (el) { el.value = fb_fmt(v); el.classList.toggle('negative-value', v < 0); }
    });
    fb_saveDraft();
}

// 草稿：固定表單（?form=）填寫中每次修改都存到這台電腦的瀏覽器，重新整理或關掉分頁後再開可還原
const FB_DRAFT = 'fb-draft:';
let fb_survey = null; // 從 ?survey= 開啟的調查專案 { id, title, form, unit, purpose, year, start, deadline, respondents }
// 草稿依專案分開（同一份範本的 114、115 年調查不互相蓋掉）；沒有專案時依表單名稱
const fb_draftKey = () => FB_DRAFT + (fb_survey ? 'survey:' + fb_survey.id : fb_form.title);
let fb_draftOn = false, fb_draftTimer = null;
function fb_saveDraft() {
    if (!fb_draftOn || !fb_form) return;
    clearTimeout(fb_draftTimer);
    fb_draftTimer = setTimeout(() => {
        document.getElementById('fb-draft').classList.add('hidden'); // 沒還原就開始填 → 舊草稿被新的取代
        fb_saveSheet();
        const val = id => document.getElementById(id).value;
        try { localStorage.setItem(fb_draftKey(), JSON.stringify({ time: Date.now(), metadata: { org: val('fb-org'), year: val('fb-year'), user: val('fb-user') }, tables: fb_form.tables })); } catch (e) { /* 無痕模式等不能存就算了 */ }
    }, 500);
}

// 有草稿就在表單上方顯示「還原／捨棄」（不用 confirm：被瀏覽器擋掉時會直接走「捨棄」）
function fb_offerDraft() {
    let d = null;
    try { d = JSON.parse(localStorage.getItem(fb_draftKey())); } catch (e) { /* 讀不到當沒有草稿 */ }
    const box = document.getElementById('fb-draft');
    box.classList.toggle('hidden', !d?.tables);
    if (d?.tables) {
        document.getElementById('fb-draft-time').textContent = new Date(d.time).toLocaleString('zh-TW');
        document.getElementById('fb-draft-restore').onclick = () => {
            fb_draftOn = false;
            fb_load({ template: { title: fb_form.title, sheets: fb_form.sheets }, metadata: d.metadata, tables: d.tables });
            box.classList.add('hidden');
            fb_draftOn = true;
        };
        document.getElementById('fb-draft-discard').onclick = () => {
            try { localStorage.removeItem(fb_draftKey()); } catch (e) { /* 同上 */ }
            box.classList.add('hidden');
        };
    }
    fb_draftOn = true;
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

var FB_INFO_SHEET = '填報資訊', FB_INFO_KEYS = { 機關: 'org', 年度: 'year', 填表人: 'user', 調查代號: 'survey' }; // var：彙整頁也要讀得到

// 已填的 Excel：欄位放回範本原本的欄（公式欄寫回 Excel 公式），可再匯入填表頁或直接拖進彙整頁
// blank = 發起人產生給查填者的空白查填檔（填報資訊留白，由查填者在 Excel 裡填）
function fb_sheetToWs(sheet, rows) {
    const cols = sheet.columns, caption = sheet.caption || [], ws = {};
    const put = (r, key, cell) => { ws[key + (r + 1)] = cell; };
    caption.forEach((t, r) => put(r, 'A', { t: 's', v: t }));
    // 表頭照範本排成多層並合併儲存格（再匯入時欄名一樣串成「預算數－原編」，可通過比對）
    const head = fb_headGrid(cols), colOf = idx => XLSX.utils.decode_col(cols[idx].key);
    head.aoa.forEach((line, k) => line.forEach((t, idx) => { if (t) put(caption.length + k, cols[idx].key, { t: 's', v: t }); }));
    ws['!merges'] = head.merges.map(m => ({ s: { r: caption.length + m.s.r, c: colOf(m.s.c) }, e: { r: caption.length + m.e.r, c: colOf(m.e.c) } }));
    const hr = caption.length + head.aoa.length - 1; // 最後一列表頭
    // 上層項目與最後的合計列寫成 Excel 公式（加總下層／最上層），在 Excel 裡填也會自動算；再匯入時合計列會被略過
    const parent = fb_parents(rows.map(r => r[cols[0].key]));
    const sumOf = (key, idx) => idx.length ? idx.map(j => key + (hr + 2 + j)).join('+') : '0';
    const numCell = (c, r, v, f) => put(r, c.key, { t: 'n', v: Number(v) || 0, f });
    rows.forEach((row, i) => {
        const r = hr + 1 + i, kids = rows.map((_, j) => j).filter(j => parent[j] === i);
        cols.forEach(c => {
            const v = row[c.key];
            if (c.formula) numCell(c, r, v, c.formula.replace(/[A-Z]+/g, k => k + (r + 1)));
            else if (c.type === 'num' && kids.length) numCell(c, r, v, sumOf(c.key, kids));
            else if (v !== '' && v !== undefined) put(r, c.key, typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v) });
        });
    });
    const tr = hr + 1 + rows.length, roots = rows.map((_, j) => j).filter(j => parent[j] < 0);
    put(tr, cols[0].key, { t: 's', v: '合計' });
    cols.slice(1).forEach(c => {
        if (c.formula) numCell(c, tr, 0, c.formula.replace(/[A-Z]+/g, k => k + (tr + 1)));
        else if (c.type === 'num') numCell(c, tr, 0, 'SUM(' + sumOf(c.key, roots).replace(/\+/g, ',') + ')');
    });
    const lastCol = cols.map(c => XLSX.utils.decode_col(c.key)).reduce((a, b) => Math.max(a, b), 0);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: tr, c: lastCol } });
    ws['!cols'] = cols.map((c, i) => ({ wch: i === 0 ? 24 : 14 }));
    return ws;
}

// 已填的 Excel：每個工作表一張（名稱照範本），最後一張「填報資訊」；可再匯入填表頁或直接拖進彙整頁
// blank = 發起人產生給查填者的空白查填檔（填報資訊留白，由查填者在 Excel 裡填）
function fb_exportXlsx(blank = false) {
    if (!fb_form) return alert('請先上傳 Excel 範本或匯入 JSON');
    fb_saveSheet();
    const val = id => blank ? '' : document.getElementById(id).value;
    const wb = XLSX.utils.book_new();
    fb_form.sheets.forEach(s => XLSX.utils.book_append_sheet(wb, fb_sheetToWs(s, fb_form.tables[s.name] || []), s.name.slice(0, 31)));
    // 調查代號沒有輸入欄，取目前的專案（空白查填檔也帶，回傳時彙整頁才認得是哪次調查）
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(Object.keys(FB_INFO_KEYS).map(k => [k, FB_INFO_KEYS[k] === 'survey' ? fb_survey?.id || '' : val('fb-' + FB_INFO_KEYS[k])])), FB_INFO_SHEET);
    const base = fb_survey?.id || fb_form.title;
    XLSX.writeFile(wb, blank ? `${base}_查填.xlsx` : `${base}_${val('fb-org') || '機關'}_${val('fb-year')}年度.xlsx`);
}

function fb_exportJSON() {
    if (!fb_form) return alert('請先上傳 Excel 範本或匯入 JSON');
    fb_saveSheet();
    const val = id => document.getElementById(id).value;
    const metadata = { org: val('fb-org'), year: val('fb-year'), user: val('fb-user'), ...(fb_survey ? { survey: fb_survey.id } : {}) };
    fb_saveJSON({ template: { title: fb_form.title, sheets: fb_form.sheets }, metadata, tables: fb_form.tables }, `${fb_survey?.id || fb_form.title}_${val('fb-org') || '機關'}_${val('fb-year')}年度.json`);
}

// 固定表單範本：放到 repo 的 forms/ 資料夾，以 ?form=檔名 開啟
// 目前表格的列（不含公式欄）會成為預設項目；不含機關／填表人
function fb_exportTemplate() {
    if (!fb_form) return alert('請先上傳 Excel 範本或匯入 JSON');
    fb_saveSheet();
    const tables = Object.fromEntries(fb_form.sheets.map(s => [s.name, (fb_form.tables[s.name] || []).map(r => {
        const x = { ...r };
        s.columns.forEach(c => { if (c.formula) delete x[c.key]; });
        return x;
    })]));
    fb_saveJSON({ template: { title: fb_form.title, sheets: fb_form.sheets }, tables }, `${fb_form.title}.json`);
    fb_msg('');
    document.getElementById('fb-publish-file').textContent = document.getElementById('fb-publish-file2').textContent = `${fb_form.title}.json`;
    document.getElementById('fb-publish-github').href = FB_UPLOAD_URL;
    document.getElementById('fb-share-url').value = FB_SITE_URL + '?form=' + encodeURIComponent(fb_form.title);
    document.getElementById('fb-publish').classList.remove('hidden');
}

const { upload: FB_UPLOAD_URL, site: FB_SITE_URL, api: FB_API_URL } = fb_repoUrls(location.hostname, location.pathname);

// 首頁（沒有 ?form=）列出 forms/ 已發布的表單；新上傳到 forms/ 的檔案會自動出現，不必另外維護清單
// GitHub API 未登入每小時 60 次（以 IP 計，機關共用對外 IP 很快就用完）：成功的清單在這台電腦快取 5 分鐘；
// 被限流或讀不到時，改用舊的快取清單，再不行就用 forms/index.json 有登記的表單，並把原因顯示在清單上方
// 範本庫讀一次，查詢只在畫面上篩選
let fb_libAll = [], fb_libIndex = {}, fb_libErr = '';
const FB_LIB_CACHE = 'fb-library', FB_LIB_TTL = 5 * 60 * 1000;

async function fb_loadLibrary() {
    const idxP = fetch('forms/' + FB_INDEX, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
    let cache = null;
    try { cache = JSON.parse(localStorage.getItem(FB_LIB_CACHE)); } catch (e) { /* 沒有快取 */ }
    let entries = cache && Date.now() - cache.time < FB_LIB_TTL ? cache.entries : null;
    fb_libErr = '';
    if (!entries) {
        try {
            const res = await fetch(FB_API_URL, { cache: 'no-cache' });
            if (!res.ok) throw new Error(res.status === 403 ? 'GitHub 查詢次數已達上限' : `GitHub 讀取失敗（${res.status}）`);
            entries = fb_libraryEntries(await res.json());
            try { localStorage.setItem(FB_LIB_CACHE, JSON.stringify({ time: Date.now(), entries })); } catch (e) { /* 存不了就每次查 */ }
        } catch (err) {
            fb_libErr = err.message;
            entries = cache?.entries || null;
            if (entries) fb_libErr += `，以下是 ${new Date(cache.time).toLocaleString('zh-TW')} 的清單`;
        }
    }
    fb_libIndex = await idxP;
    if (!entries) {
        entries = fb_libraryFromIndex(fb_libIndex);
        if (fb_libErr) fb_libErr += entries.length ? '，以下只列出範本庫有登記的表單（可能不完整）' : '，請稍後再試';
    }
    fb_libAll = entries;
    const cats = fb_libraryGroups(fb_libAll, fb_libIndex).map(g => g.category);
    document.getElementById('fb-lib-cat').innerHTML = '<option value="">全部分類</option>' + cats.map(c => `<option>${fb_esc(c)}</option>`).join('');
    fb_renderLibrary();
}

function fb_renderLibrary() {
    const list = document.getElementById('fb-library-list');
    const q = document.getElementById('fb-lib-q').value, cat = document.getElementById('fb-lib-cat').value;
    const groups = fb_libraryGroups(fb_libAll, fb_libIndex).map(g => ({ ...g, items: g.items.filter(it => fb_libraryMatch(it, q, cat)) })).filter(g => g.items.length);
    const err = fb_libErr ? `<li class="py-2 text-amber-700">⚠️ ${fb_esc(fb_libErr)}</li>` : ''; // 錯誤一直顯示，不被查詢結果蓋掉
    const used = name => fb_surveyEntries(fb_svIndex, true).filter(e => e.meta.form === name).length;
    list.innerHTML = err + (groups.length ? groups.map(g => `<li class="pt-3 pb-1 text-sm font-bold text-slate-500">${fb_esc(g.category)}</li>` + g.items.map(e => `<li class="flex flex-wrap items-center gap-2 py-2 border-b last:border-0">
            <div class="mr-auto"><span class="font-bold text-slate-800">${fb_esc(e.name)}</span>
                ${used(e.name) ? `<span class="text-xs text-slate-500 ml-2">已用於 ${used(e.name)} 個調查</span>` : ''}
                ${e.meta.description ? `<div class="text-xs text-slate-500">${fb_esc(e.meta.description)}</div>` : ''}</div>
            <a href="manage.html?form=${encodeURIComponent(e.name)}#surveys" class="btn btn-sm" title="建立用這份範本的調查專案（需要 GitHub 權杖）">以此範本發起調查</a>
            <a href="?form=${encodeURIComponent(e.name)}" class="btn btn-sm" title="開啟這份範本的表單（沒有專案資訊）">預覽</a>
            ${e.file ? `<a href="forms/${encodeURIComponent(e.file)}" download class="btn btn-sm" title="下載後可修改再上傳成新範本">範本檔（${fb_esc(e.file.split('.').pop())}）</a>` : ''}
            ${e.xlsx ? `<a href="forms/${encodeURIComponent(e.xlsx)}" download class="btn btn-sm" title="寄給填報人的空白 Excel">空白 Excel</a>` : ''}
        </li>`).join('')).join('') : (fb_libAll.length ? '<li class="py-2 text-slate-500">沒有符合的範本。</li>' : (fb_libErr ? '' : '<li class="py-2 text-slate-500">forms/ 裡還沒有範本。</li>')));
}

// ========== 調查專案（首頁「調查專案」「查填」「彙整」三個分頁） ==========
let fb_svIndex = {};

async function fb_loadSurveys() {
    fb_svIndex = await fetch(FB_SURVEY_INDEX, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
    fb_renderSurveys();
    if (fb_libAll.length) fb_renderLibrary(); // 範本的「已用於幾個調查」
}

// 一筆專案的說明文字：代號・範本・發起單位・年度・期間，下一行用途
function fb_surveyInfo(e) {
    const m = e.meta, st = fb_libraryStatus(m);
    const period = [m.start ? '發起 ' + m.start : '', m.deadline ? '截止 ' + m.deadline : ''].filter(Boolean).join('～');
    return `<span class="font-bold text-slate-800">${fb_esc(m.title || e.name)}</span>
        ${st ? `<span class="badge ${FB_BADGE[st]}">${FB_STATUS[st]}</span>` : ''}
        <div class="text-xs text-slate-500">${[`代號 ${e.name}`, m.form ? `範本 ${m.form}` : '', m.unit, m.year ? m.year + ' 年度' : '', period].filter(Boolean).map(fb_esc).join('・')}</div>
        ${m.purpose ? `<div class="text-xs text-slate-600">用途：${fb_esc(m.purpose)}</div>` : ''}`;
}

function fb_renderSurveys() {
    const all = fb_surveyEntries(fb_svIndex), today = fb_today();
    const q = document.getElementById('fb-sv-q').value, stf = document.getElementById('fb-sv-status').value;
    const empty = '<li class="py-2 text-slate-500">還沒有調查專案。到範本庫按「以此範本發起調查」，或到「管理範本與調查」新增。</li>';
    // 調查專案：依狀態分組（進行中、未開始、已結束、未設期間）
    const order = ['open', 'upcoming', 'closed', ''];
    const shown = all.filter(e => fb_libraryMatch(e, q, '', stf, today));
    document.getElementById('fb-survey-list').innerHTML = !all.length ? empty : !shown.length ? '<li class="py-2 text-slate-500">沒有符合的調查。</li>' :
        order.map(st => {
            const items = shown.filter(e => fb_libraryStatus(e.meta, today) === st);
            if (!items.length) return '';
            return `<li class="pt-3 pb-1 text-sm font-bold text-slate-500">${FB_STATUS[st] || '未設定期間'}</li>` + items.map(e => {
                const link = FB_SITE_URL + '?survey=' + encodeURIComponent(e.name), n = (e.meta.respondents || []).length;
                return `<li class="flex flex-wrap items-center gap-2 py-2 border-b last:border-0">
                    <div class="mr-auto">${fb_surveyInfo(e)}${n ? `<div class="text-xs text-slate-500">應繳 ${n} 個機關</div>` : ''}</div>
                    <button class="fb-lib-copy btn btn-sm" data-link="${fb_esc(link)}">複製查填連結</button>
                    <a href="?survey=${encodeURIComponent(e.name)}" class="btn btn-sm">開啟</a>
                    <a href="?survey=${encodeURIComponent(e.name)}&blank=1" class="btn btn-sm" title="下載空白查填用 Excel（含調查代號），寄給習慣用 Excel 的填報人">查填用 Excel</a>
                    <a href="form-aggregate.html?survey=${encodeURIComponent(e.name)}" class="btn btn-sm">彙整</a>
                </li>`;
            }).join('');
        }).join('');
    // 查填：只列進行中的
    const open = all.filter(e => fb_libraryStatus(e.meta, today) === 'open');
    document.getElementById('fb-fill-list').innerHTML = open.length ? open.map(e => `<li class="flex flex-wrap items-center gap-2 py-2 border-b last:border-0">
            <div class="mr-auto">${fb_surveyInfo(e)}</div>
            <a href="?survey=${encodeURIComponent(e.name)}" class="btn btn-sm">開始填寫</a>
        </li>`).join('') : '<li class="py-2 text-slate-500">目前沒有進行中的調查。</li>';
    // 彙整：全部（含已結束）
    document.getElementById('fb-agg-list').innerHTML = all.length ? all.map(e => `<li class="flex flex-wrap items-center gap-2 py-2 border-b last:border-0">
            <div class="mr-auto">${fb_surveyInfo(e)}</div>
            <a href="form-aggregate.html?survey=${encodeURIComponent(e.name)}" class="btn btn-sm">彙整</a>
        </li>`).join('') : empty;
}

// 首頁分頁：網址 #surveys／#templates／#fill／#agg 記住目前分頁
function fb_showTab(key) {
    if (!document.querySelector(`.fb-hub-panel[data-tab="${key}"]`)) key = 'surveys';
    document.querySelectorAll('.fb-hub-panel').forEach(p => p.classList.toggle('hidden', p.dataset.tab !== key));
    document.querySelectorAll('.fb-hub-tab').forEach(b => { b.classList.toggle('fb-tab-on', b.dataset.tab === key); b.setAttribute('aria-selected', b.dataset.tab === key); });
}

function fb_load(data, metaYear) {
    const form = fb_normalize(data);
    // 手改或其他程式產生的 JSON 可能把數字存成 "1,234"、"(300)"：與 Excel 匯入同樣清理
    form.warnings = form.sheets.flatMap(s => { const w = fb_cleanRows(form.tables[s.name] || [], s.columns); return w ? [form.sheets.length > 1 ? `［${s.name}］${w}` : w] : []; });
    const m = data.metadata || {};
    document.getElementById('fb-org').value = m.org || '';
    document.getElementById('fb-year').value = m.year || '';
    document.getElementById('fb-user').value = m.user || '';
    fb_renderForm(form, metaYear);
    return form;
}

function fb_importJSON(text) {
    const form = fb_load(JSON.parse(text));
    const n = Object.values(form.tables).reduce((a, r) => a + r.length, 0);
    fb_msg(`已匯入「${form.title}」，${form.sheets.length > 1 ? form.sheets.length + ' 個工作表、' : ''}${n} 列。`, form.warnings);
}

// ?form=名稱 → 讀 forms/名稱.json；沒有的話直接讀 forms/名稱.xlsx（可只上傳 Excel，不必先匯出範本）
// ?survey=代號：讀專案 → 載入專案用的範本，上方顯示專案資訊；機關欄有應繳名單可選
async function fb_loadSurvey(id) {
    if (fb_badName(id)) throw new Error('調查代號不合法：' + id);
    const idx = await fetch(FB_SURVEY_INDEX, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
    const s = idx.surveys?.[id];
    if (!s) throw new Error(`找不到調查「${id}」（可能已刪除或代號打錯，請洽承辦）`);
    fb_survey = { id, ...s };
    await fb_loadFixedForm(s.form, s.year);
    const st = fb_libraryStatus(s);
    const period = [s.start ? '發起 ' + s.start : '', s.deadline ? '截止 ' + s.deadline : ''].filter(Boolean).join('～');
    document.getElementById('fb-survey-bar').innerHTML = `<div class="font-bold text-slate-800 text-base">${fb_esc(s.title || id)}
        ${st ? `<span class="badge ${FB_BADGE[st]}">${FB_STATUS[st]}</span>` : ''}</div>
        <div class="text-slate-600">${[s.unit, s.year ? s.year + ' 年度' : '', period].filter(Boolean).map(fb_esc).join('・')}</div>
        ${s.purpose ? `<div class="text-slate-600">用途：${fb_esc(s.purpose)}</div>` : ''}
        ${st === 'closed' ? '<div class="text-red-600 font-bold mt-1">⚠️ 這次調查已截止，仍可填寫，但請先與承辦聯繫。</div>' : ''}`;
    document.getElementById('fb-survey-bar').classList.remove('hidden');
    document.getElementById('fb-orgs').innerHTML = (s.respondents || []).map(o => `<option value="${fb_esc(o)}">`).join('');
    document.title = `${s.title || id}｜自訂表單查填`;
}

async function fb_loadFixedForm(name, yearOverride) {
    if (/[\\/]|\.\./.test(name)) throw new Error('表單名稱不合法：' + name);
    const base = `forms/${encodeURIComponent(name)}`;
    let warnings = [];
    const [res, idx] = await Promise.all([
        fetch(base + '.json', { cache: 'no-cache' }),
        fetch('forms/' + FB_INDEX, { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({})),
    ]);
    const metaYear = yearOverride || idx.templates?.[name]?.year; // 專案年度 → 範本庫設定的年度
    if (res.ok) fb_load(await res.json(), metaYear);
    else {
        const xres = await fetch(base + '.xlsx', { cache: 'no-cache' });
        if (!xres.ok) throw new Error(`找不到固定表單「${name}」（forms/ 裡沒有 ${name}.json 或 ${name}.xlsx）`);
        // ponytail: 直接放 xlsx 只讀第一個工作表；多工作表的表單請發布「匯出範本」的 .json
        const parsed = fb_readWorkbook(await xres.arrayBuffer(), name + '.xlsx');
        fb_renderForm(fb_formFromParsed(name, [parsed]), metaYear);
        warnings = parsed.warnings;
    }
    document.getElementById('fb-btn-xlsx').classList.add('hidden');
    document.getElementById('fb-btn-template').classList.add('hidden');
    document.getElementById('fb-link-aggregate').classList.add('hidden');
    fb_msg(`固定表單「${name}」：請填寫機關、填表人與各列數值後按「匯出 JSON」回傳。`, warnings);
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
    $('fb-btn-export-xlsx').onclick = () => fb_exportXlsx();
    $('fb-btn-blank-xlsx').onclick = () => fb_exportXlsx(true);
    $('fb-btn-template').onclick = fb_exportTemplate;
    $('fb-btn-copy').onclick = async () => {
        const url = $('fb-share-url');
        try { await navigator.clipboard.writeText(url.value); } catch (e) { url.select(); document.execCommand('copy'); }
        $('fb-btn-copy').textContent = '已複製 ✓';
        setTimeout(() => { $('fb-btn-copy').textContent = '複製連結'; }, 1500);
    };
    $('fb-btn-add').onclick = () => { fb_addRow(); fb_calc(); };

    // 範本有多個工作表時列出勾選框，勾幾個就有幾個分頁要填（預設第一個）；改勾選即重新產生表單
    let tplBook = null, tplName = '';
    const loadTemplate = names => {
        const all = fb_visibleSheets(tplBook);
        names = names || [all[0] || tplBook.SheetNames[0]];
        const list = [], warnings = [];
        names.forEach(n => {
            try {
                const p = fb_readWorkbook(tplBook, tplName, n);
                list.push(p);
                warnings.push(...p.warnings.map(w => names.length > 1 ? `［${n}］${w}` : w));
            } catch (err) { warnings.push(`［${n}］${err.message}，略過`); }
        });
        if (!list.length) throw new Error(warnings.join('；'));
        const form = fb_formFromParsed(tplName.replace(/\.[^.]+$/, ''), list);
        fb_renderForm(form);
        $('fb-sheet-list').innerHTML = all.map(n => `<label class="mr-3 whitespace-nowrap"><input type="checkbox" class="fb-sheet-pick" value="${fb_esc(n)}"${names.includes(n) ? ' checked' : ''}> ${fb_esc(n)}</label>`).join('');
        $('fb-sheet-box').classList.toggle('hidden', all.length < 2);
        const rows = list.reduce((a, p) => a + p.rows.length, 0), nc = list.reduce((a, p) => a + p.template.columns.length, 0);
        fb_msg(`已由「${tplName}」產生表單：${list.length > 1 ? list.length + ' 個工作表、' : `工作表「${list[0].sheet}」、`}${nc} 欄、${rows} 個預設項目。`, warnings);
    };
    $('fb-xlsx-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try { tplBook = fb_openWorkbook(await file.arrayBuffer()); tplName = file.name; loadTemplate(); }
        catch (err) { alert('Excel 解析失敗：' + err.message); }
    };
    $('fb-sheet-list').addEventListener('change', () => {
        const names = [...document.querySelectorAll('.fb-sheet-pick:checked')].map(c => c.value);
        if (!names.length) return alert('至少要勾選一個工作表');
        try { loadTemplate(names); } catch (err) { alert('Excel 解析失敗：' + err.message); }
    });
    $('fb-tabs').addEventListener('click', e => { const b = e.target.closest('.fb-tab'); if (b) fb_showSheet(+b.dataset.i); });
    // 填表人把已在 Excel 填好的檔案匯入：每個工作表欄名須與表單一致，表單結構不變，只換掉各列
    $('fb-btn-fill-xlsx').onclick = () => $('fb-fill-xlsx-file').click();
    $('fb-fill-xlsx-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!fb_form || !file) return;
        try {
            const { found, notes, metadata } = fb_matchSheets(await file.arrayBuffer(), file.name, fb_form.sheets);
            const hit = Object.keys(found);
            if (!hit.length) return alert('欄位與表單不一致，未匯入：\n' + notes.join('\n'));
            fb_saveSheet();
            hit.forEach(n => { fb_form.tables[n] = found[n].rows; });
            Object.entries(metadata).forEach(([k, v]) => { if (v && $('fb-' + k)) $('fb-' + k).value = v; }); // 調查代號沒有輸入欄
            fb_form.cur = null; // 目前分頁也換成匯入的資料，不要再把畫面上的舊列存回去
            fb_showSheet(fb_form.sheets.findIndex(s => s.name === hit[0]));
            const n = hit.reduce((a, k) => a + found[k].rows.length, 0);
            const where = hit.map(k => found[k].sheet === k ? `「${k}」` : `「${k}」← 工作表「${found[k].sheet}」`).join('、');
            fb_msg(`已由「${file.name}」匯入 ${n} 列（${where}）。`, [...notes, ...hit.flatMap(k => found[k].warnings.map(w => hit.length > 1 ? `［${k}］${w}` : w))]);
        } catch (err) { alert('Excel 匯入失敗：' + err.message); }
    };
    $('fb-json-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try { fb_importJSON(await file.text()); } catch (err) { alert('JSON 匯入失敗：' + err.message); }
    };

    $('fb-tbody').addEventListener('input', fb_calc);
    ['fb-org', 'fb-year', 'fb-user'].forEach(id => $(id).addEventListener('input', fb_saveDraft));
    $('fb-tbody').addEventListener('paste', fb_onPaste);
    $('fb-tbody').addEventListener('click', e => {
        if (e.target.classList.contains('fb-del')) { e.target.closest('tr').remove(); fb_calc(); }
    });

    const params = new URLSearchParams(location.search), form = params.get('form'), survey = params.get('survey');
    // ?form=名稱&blank=1：範本庫的「查填用 Excel」，載入表單後直接下載空白查填檔
    const afterLoad = () => { if (params.get('blank')) fb_exportXlsx(true); else fb_offerDraft(); };
    if (survey) fb_loadSurvey(survey).then(afterLoad).catch(err => fb_msg('⚠️ ' + err.message));
    else if (form) fb_loadFixedForm(form).then(afterLoad).catch(err => fb_msg('⚠️ ' + err.message));
    else { $('fb-hub').classList.remove('hidden'); fb_showTab(location.hash.slice(1)); fb_loadLibrary(); fb_loadSurveys(); }
    $('fb-hub-new').onclick = () => $('fb-xlsx-file').click();
    $('fb-lib-q').addEventListener('input', fb_renderLibrary);
    $('fb-lib-cat').addEventListener('change', fb_renderLibrary);
    $('fb-sv-q').addEventListener('input', fb_renderSurveys);
    $('fb-sv-status').addEventListener('change', fb_renderSurveys);
    $('fb-hub-import').onclick = () => $('fb-json-file').click();
    document.querySelectorAll('.fb-hub-tab').forEach(b => b.addEventListener('click', () => { history.replaceState(null, '', '#' + b.dataset.tab); fb_showTab(b.dataset.tab); }));
    $('fb-hub').addEventListener('click', async e => {
        const btn = e.target.closest('.fb-lib-copy');
        if (!btn) return;
        const text = btn.textContent;
        await navigator.clipboard.writeText(btn.dataset.link).catch(() => {});
        btn.textContent = '已複製 ✓';
        setTimeout(() => { btn.textContent = text; }, 1500);
    });
});

}
