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

// "1,234" → 1234、"(500)" → -500、"－" 或空白 → ''；不是數字就原樣回傳讓使用者看到
function fb_cleanNum(s) {
    let t = String(s).replace(/[,\s]/g, '');
    if (t === '' || t === '-' || t === '－') return '';
    const neg = /^\(.*\)$/.test(t);
    if (neg) t = t.slice(1, -1);
    return isNaN(Number(t)) ? s.trim() : String(neg ? -Number(t) : Number(t));
}

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

// 每列的父列索引（-1 = 最上層）：往上找最近一個編號層級較高的列；沒編號的列跳過
function fb_parents(names) {
    const lv = names.map(fb_level);
    return lv.map((l, i) => {
        if (!l) return -1;
        for (let j = i - 1; j >= 0; j--) if (lv[j] && lv[j] < l) return j;
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
        nums.forEach(c => { sums[p][c.key] += parseFloat(rows[i][c.key]) || 0; });
    }
    return rows;
}

// 公式是否參照了別列（例如合計列 =SUM(B3:B9)）
const fb_refsOtherRow = (f, row) => [...f.toUpperCase().matchAll(/\$?[A-Z]+\$?(\d+)/g)].some(m => +m[1] !== row);

// 填好的 Excel 匯入前比對欄名，不一致就拒絕（回傳錯誤訊息，一致回傳 null）
function fb_headerMismatch(expected, got) {
    if (expected.length !== got.length) return `欄位數量不符：表單 ${expected.length} 欄，檔案 ${got.length} 欄\n表單：${expected.join('、')}\n檔案：${got.join('、')}`;
    const i = expected.findIndex((l, k) => l !== got[k]);
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
    let hr = r0;
    while (hr <= r1 && (filled(hr) === 0 || (c1 > c0 && filled(hr) === 1))) hr++; // ponytail: 單欄範本的表頭若只有一格會被當表名跳過
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

// 「前往 GitHub 發布」與分享連結跟著網站所在的帳號／repo 走，別人複製這個專案後不必改程式
// alice.github.io/local_gov_budget/… → github.com/alice/local_gov_budget；本機開發時用原作者的 repo
function fb_repoUrls(hostname, pathname) {
    const m = /^([^.]+)\.github\.io$/i.exec(hostname);
    if (!m) return { upload: 'https://github.com/cormort/local_gov_budget/upload/main/forms', site: 'https://cormort.github.io/local_gov_budget/' };
    const owner = m[1].toLowerCase();
    const first = pathname.split('/')[1] || '';
    const isProjectSite = first && !first.includes('.');            // /repo/… 是專案網站；/xxx.html 是 owner.github.io 本身
    const repo = isProjectSite ? first : hostname.toLowerCase();
    return {
        upload: `https://github.com/${owner}/${repo}/upload/main/forms`,
        site: `https://${hostname.toLowerCase()}/${isProjectSite ? repo + '/' : ''}`,
    };
}

if (typeof module !== 'undefined') module.exports = { fb_repoUrls, fb_translateFormula, fb_evalFormula, fb_parseTSV, fb_cleanNum, fb_headerBlock, fb_columnLevels, fb_theadRows, fb_level, fb_parents, fb_refsOtherRow, fb_headerMismatch, fb_rollup };
if (typeof document === 'undefined') { /* node 測試用 */ } else {

// ========== 2. 讀 Excel 範本 ==========
let fb_template = null;

function fb_readWorkbook(buf, fileName) {
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true, sheetStubs: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const cell = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
    const text = (r, c) => { const x = cell(r, c); return x ? String(x.w ?? x.v ?? '').trim() : ''; };

    // 表頭範圍：可能跨多列（合併儲存格），欄名以「－」串接各層，例如「預算數－原編」
    const merges = ws['!merges'] || [];
    const block = fb_headerBlock(text, merges, range.s.r, range.e.r, range.s.c, range.e.c);
    if (!block) throw new Error('找不到表頭列');
    const { hr, he } = block;

    const columns = [], warnings = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const levels = fb_columnLevels(text, merges, hr, he, c);
        if (!levels.length) continue;
        const label = levels.join('－');
        const key = XLSX.utils.encode_col(c);
        const col = { key, label, type: columns.length === 0 ? 'text' : 'num' };
        if (levels.length > 1) col.levels = levels;
        if (columns.length > 0) {
            // 取表頭下方第一個同列公式當該欄公式；跨列公式（合計、小計列）不算，改由層級小計與表尾合計計算
            let unsupported = null;
            for (let r = he + 1; r <= range.e.r; r++) {
                const x = cell(r, c);
                if (x && x.f) {
                    const expr = fb_translateFormula(x.f, r + 1);
                    if (expr) { col.formula = expr; break; }
                    if (!fb_refsOtherRow(x.f, r + 1)) unsupported ??= x.f;
                }
                if (x && x.t === 's' && x.v.trim()) col.type = 'text';
            }
            if (!col.formula && unsupported) warnings.push(`「${label}」公式 =${unsupported} 不支援，改為手填`);
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
    // 沒有編號又用跨列公式加總的列（合計、小計）略過：表單由表尾合計與層級小計自動算，留著會重複加總
    const rows = [], skipped = [];
    for (let r = he + 1; r <= range.e.r; r++) {
        const name = text(r, XLSX.utils.decode_col(columns[0].key));
        if (!name) continue;
        if (!fb_level(name) && columns.some(c => { const x = cell(r, XLSX.utils.decode_col(c.key)); return x?.f && fb_refsOtherRow(x.f, r + 1); })) { skipped.push(name); continue; }
        const row = {};
        columns.forEach(c => {
            if (c.formula) return;
            const x = cell(r, XLSX.utils.decode_col(c.key));
            row[c.key] = x ? (c.type === 'num' && typeof x.v === 'number' ? x.v : String(x.v ?? '')) : '';
        });
        rows.push(row);
    }
    if (skipped.length) warnings.push(`略過以公式加總的「${skipped.join('、')}」列（表單會自動計算合計）`);

    // 表頭上方的表名、「單位：千元」等文字列 → 顯示在表格上方
    const caption = [];
    for (let r = range.s.r; r < hr; r++) {
        const line = [];
        for (let c = range.s.c; c <= range.e.c; c++) if (text(r, c)) line.push(text(r, c));
        if (line.length) caption.push(line.join(' '));
    }

    const template = { title: fileName.replace(/\.[^.]+$/, ''), sheet: wb.SheetNames[0], columns };
    if (caption.length) template.caption = caption;
    return { template, rows, warnings };
}

// ========== 3. 畫表 + 計算 ==========
const fb_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

function fb_render(template, rows) {
    fb_template = template;
    document.getElementById('fb-publish').classList.add('hidden');
    document.getElementById('fb-title').value = template.title;
    document.getElementById('fb-heading').textContent = template.title;
    document.getElementById('fb-caption').textContent = (template.caption || []).join('\n');
    document.getElementById('fb-btn-fill-xlsx').classList.remove('hidden');
    const head = fb_theadRows(template.columns);
    document.getElementById('fb-thead').innerHTML = head.map((row, k) =>
        '<tr>' + (k === 0 ? `<th class="w-8" rowspan="${head.length}"></th>` : '') +
        row.map(h => `<th colspan="${h.colspan}" rowspan="${h.rowspan}">${fb_esc(h.text)}${h.col?.formula ? ' 🔒' : ''}</th>`).join('') +
        '</tr>').join('');
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

// 從游標所在儲存格開始，依 Excel 的欄位順序往右、往下填；公式欄跳過（照樣佔一欄，與 Excel 版面一致），列不夠就新增
function fb_onPaste(e) {
    const input = e.target.closest('input');
    const text = e.clipboardData?.getData('text/plain') || '';
    if (!input || !/[\t\n\r]/.test(text)) return; // 單一儲存格交給瀏覽器預設處理
    e.preventDefault();
    const cols = fb_template.columns;
    const startCol = cols.findIndex(c => input.classList.contains('v-' + c.key));
    let tr = input.closest('tr');
    fb_parseTSV(text).forEach((cells, r) => {
        if (r > 0) {
            if (!tr.nextElementSibling) fb_addRow();
            tr = tr.nextElementSibling;
        }
        cells.forEach((v, k) => {
            const c = cols[startCol + k];
            if (!c || c.formula) return;
            tr.querySelector('.v-' + c.key).value = c.type === 'num' ? fb_cleanNum(v) : v.trim();
        });
    });
    fb_calc();
}

function fb_calc() {
    if (!fb_template) return;
    const cols = fb_template.columns, totals = {};
    const trs = [...document.querySelectorAll('#fb-tbody tr')];
    const parent = fb_parents(trs.map(tr => tr.querySelector('.v-' + cols[0].key).value));
    const sums = trs.map(() => ({}));
    // 由下往上：子列一定在父列下方，算到父列時子列已加進 sums；合計只加最上層的列，避免重複加總
    for (let i = trs.length - 1; i >= 0; i--) {
        const tr = trs[i], el = k => tr.querySelector('.v-' + k);
        const get = k => parseFloat(el(k)?.value) || 0;
        const isParent = parent.includes(i);
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
            into[c.key] = (into[c.key] || 0) + v; // ponytail: 合計一律直接加總，比率類公式欄的合計無意義
        });
    }
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

// 固定表單範本：放到 repo 的 forms/ 資料夾，以 ?form=檔名 開啟
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
    document.getElementById('fb-share-url').value = FB_SITE_URL + '?form=' + encodeURIComponent(fb_template.title);
    document.getElementById('fb-publish').classList.remove('hidden');
}

const { upload: FB_UPLOAD_URL, site: FB_SITE_URL } = fb_repoUrls(location.hostname, location.pathname);

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

// ?form=名稱 → 讀 forms/名稱.json；沒有的話直接讀 forms/名稱.xlsx（可只上傳 Excel，不必先匯出範本）
async function fb_loadFixedForm(name) {
    if (/[\\/]|\.\./.test(name)) throw new Error('表單名稱不合法：' + name);
    const base = `forms/${encodeURIComponent(name)}`;
    let warnings = [];
    const res = await fetch(base + '.json', { cache: 'no-cache' });
    if (res.ok) fb_load(await res.json());
    else {
        const xres = await fetch(base + '.xlsx', { cache: 'no-cache' });
        if (!xres.ok) throw new Error(`找不到固定表單「${name}」（forms/ 裡沒有 ${name}.json 或 ${name}.xlsx）`);
        const parsed = fb_readWorkbook(await xres.arrayBuffer(), name + '.xlsx');
        fb_render(parsed.template, parsed.rows);
        warnings = parsed.warnings;
    }
    document.getElementById('fb-btn-xlsx').classList.add('hidden');
    document.getElementById('fb-btn-template').classList.add('hidden');
    document.getElementById('fb-link-aggregate').classList.add('hidden');
    fb_msg(`固定表單「${name}」：請填寫機關、填表人與各列數值後按「💾 匯出 JSON」回傳。`, warnings);
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
    // 填表人把已在 Excel 填好的檔案匯入：欄名須與目前表單完全一致，表單結構不變，只換掉各列
    $('fb-btn-fill-xlsx').onclick = () => $('fb-fill-xlsx-file').click();
    $('fb-fill-xlsx-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!file || !fb_template) return;
        try {
            const { template, rows } = fb_readWorkbook(await file.arrayBuffer(), file.name);
            const bad = fb_headerMismatch(fb_template.columns.map(c => c.label), template.columns.map(c => c.label));
            if (bad) return alert('欄位與表單不一致，未匯入：\n' + bad);
            fb_render(fb_template, rows);
            fb_msg(`已由「${file.name}」匯入 ${rows.length} 列。`);
        } catch (err) { alert('Excel 匯入失敗：' + err.message); }
    };
    $('fb-json-file').onchange = async e => {
        const file = e.target.files[0];
        e.target.value = '';
        if (!file) return;
        try { fb_importJSON(await file.text()); } catch (err) { alert('JSON 匯入失敗：' + err.message); }
    };

    $('fb-tbody').addEventListener('input', fb_calc);
    $('fb-tbody').addEventListener('paste', fb_onPaste);
    $('fb-tbody').addEventListener('click', e => {
        if (e.target.classList.contains('fb-del')) { e.target.closest('tr').remove(); fb_calc(); }
    });

    const form = new URLSearchParams(location.search).get('form');
    if (form) fb_loadFixedForm(form).catch(err => fb_msg('⚠️ ' + err.message));
});

}
