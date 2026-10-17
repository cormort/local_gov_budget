/**
 * xlsx-import.js — 上傳官方 Excel → 差異預覽 → 套用至目前資料
 *
 * 只是「借用」xlsx2webform 的上傳 xlsx 思路，不是搬它的 cell-grid 模型：
 * 官方檔案只會被拿來**補值**（預設寫 bud114），表格結構、層級、自動加總與審核鏈
 * 仍然由 special-fund.js 自己管。
 *
 * 三段式比對（重點）：
 *   exact    正規化後完全相同      → 自動勾選
 *   contains 一方包含另一方        → 需人工確認
 *   fuzzy    相似度 ≥ 門檻         → 需人工確認
 *   unmatched 找不到對應欄位       → 只列清單，永不寫入
 *
 * 模組同時給瀏覽器與 Node 測試用：require() 會拿到 core（純函式，不碰 DOM）。
 */
(function (root, factory) {
    var api = factory(root);
    root.LGBXlsxImport = api;
    if (typeof module === 'object' && module.exports) module.exports = api.core;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
    'use strict';

    // ===================== core：正規化 =====================
    var CN_NUM = '〇一二三四五六七八九十百壹貳參肆伍陸柒捌玖拾';

    // xlsx 表頭：去掉「(本年度預算數)」等年度後綴
    function normalizeHeader(h) {
        return String(h == null ? '' : h)
            .replace(/[（(]\s*(本年度預算數|上年度預算數|本年度決算數|前年度決算數|本年度概算數)\s*[)）]\s*$/, '')
            .trim();
    }

    // 列名：去掉「一、」「壹、」「1.」「(1)」「(一)」等編號前綴與尾端註記括號
    function normalizeName(s) {
        var t = String(s == null ? '' : s).replace(/\u3000/g, ' ').trim();
        t = t.replace(new RegExp('^[' + CN_NUM + ']+\\s*[、.．]\\s*'), '');
        t = t.replace(new RegExp('^[（(]\\s*[' + CN_NUM + ']+\\s*[)）]\\s*'), '');
        t = t.replace(/^[（(]?\s*\d+\s*[)）.．]\s*/, '');
        t = t.replace(/[（(][^（()）]*[)）]\s*$/, '');
        return t.trim();
    }

    // 比對用：去掉標點與空白、全形轉半形數字
    function unify(s) {
        return String(s == null ? '' : s)
            .replace(/[\s\u3000\u00b7\u2027・,，、。.()（）「」『』\[\]【】]/g, '')
            .replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xfee0); })
            .trim();
    }

    function similarity(a, b) {
        var m = a.length, n = b.length;
        if (!m || !n) return 0;
        var prev = new Array(n + 1), cur = new Array(n + 1), i, j;
        for (j = 0; j <= n; j++) prev[j] = j;
        for (i = 1; i <= m; i++) {
            cur[0] = i;
            for (j = 1; j <= n; j++) {
                cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
            }
            var tmp = prev; prev = cur; cur = tmp;
        }
        return 1 - prev[n] / Math.max(m, n);
    }

    function toNumber(v) {
        if (typeof v === 'number') return isFinite(v) ? v : null;
        if (typeof v === 'string') {
            var s = v.replace(/[,\s]/g, '');
            if (!s) return null;
            var n = Number(s);
            return isFinite(n) ? n : null;
        }
        return null;
    }

    // 元 → 千元，避免 0.1+0.2 這種浮點尾差
    function scaleValue(v, scale) {
        var n = v / (scale || 1);
        return Math.round(n * 1e6) / 1e6;
    }

    // ===================== core：解析官方檔 =====================
    /**
     * @param {Array<Array>} aoa       XLSX.utils.sheet_to_json(ws, {header:1}) 的結果
     * @param {Object} profile

     * @returns {{funds:Object, missingFunds:Array}}
     */
    function parseSheet(aoa, profile) {
        if (!Array.isArray(aoa) || !aoa.length) throw new Error('工作表是空的');
        var headers = (aoa[0] || []).map(normalizeHeader);
        var keyIdx = headers.indexOf(profile.keyColumn);
        if (keyIdx < 0) keyIdx = 0;

        var split = profile.sectionSplit || {};
        var useIdx = split.useColumn ? headers.indexOf(split.useColumn) : -1;
        var srcEnd = useIdx >= 0 ? useIdx : headers.length;

        function pick(row, from, to) {
            var out = {};
            for (var c = from; c < to && c < headers.length; c++) {
                if (c === keyIdx || !headers[c]) continue;
                var num = toNumber(row[c]);
                if (num === null) continue;
                out[headers[c]] = scaleValue(num, profile.scale);
            }
            return out;
        }

        var funds = {};
        var missing = [];
        Object.keys(profile.funds).forEach(function (code) {
            var official = profile.funds[code];
            var row = null;
            for (var r = 1; r < aoa.length; r++) {
                var name = String((aoa[r] || [])[keyIdx] == null ? '' : aoa[r][keyIdx]).replace(/\u3000/g, ' ').trim();
                if (unify(name) === unify(official)) { row = aoa[r]; break; }
            }
            if (!row) { missing.push(official); return; }
            funds[code] = {
                src: pick(row, 0, srcEnd),
                use: pick(row, useIdx >= 0 ? useIdx : srcEnd, headers.length)
            };
        });

        var columns = { src: [], use: [] };
        for (var i = 0; i < headers.length; i++) {
            if (i === keyIdx || !headers[i]) continue;
            if (i < srcEnd) columns.src.push(headers[i]);
            else columns.use.push(headers[i]);
        }

        return { funds: funds, columns: columns, missingFunds: missing, headers: headers };
    }

    // ===================== core：名稱 → 欄位 =====================
    function resolveColumn(name, columns, profile) {
        var target = normalizeName(name);
        var alias = (profile.nameAliases || {})[target];
        var want = unify(alias || target);
        var threshold = profile.fuzzyThreshold == null ? 0.75 : profile.fuzzyThreshold;
        var minLen = profile.minNameLength == null ? 3 : profile.minNameLength;

        var i;
        // 1) 完全相同（或命中人工別名）
        for (i = 0; i < columns.length; i++) {
            if (unify(normalizeName(columns[i])) === want) {
                return { column: columns[i], how: alias ? 'alias' : 'exact', score: 1 };
            }
        }
        // 名稱太短（如「其他」）不做寬鬆比對，避免撞上「其他收入」這種母項
        if (want.length < minLen) return { column: null, how: 'too-short' };

        // 2) 包含：官方欄名常夾舊名稱註記，或系統列名是官方欄名的子字串
        var bestContains = null, bestLen = 0;
        for (i = 0; i < columns.length; i++) {
            var u = unify(normalizeName(columns[i]));
            if (!u) continue;
            if (u.indexOf(want) >= 0 || want.indexOf(u) >= 0) {
                if (u.length > bestLen) { bestLen = u.length; bestContains = columns[i]; }
            }
        }
        if (bestContains) return { column: bestContains, how: 'contains', score: bestLen / Math.max(bestLen, want.length) };

        // 3) 模糊
        var best = null, bestScore = 0;
        for (i = 0; i < columns.length; i++) {
            var s = similarity(want, unify(normalizeName(columns[i])));
            if (s > bestScore) { bestScore = s; best = columns[i]; }
        }
        if (best && bestScore >= threshold) return { column: best, how: 'fuzzy', score: bestScore };
        return { column: null, how: 'unmatched', best: best, score: bestScore };
    }

    // ===================== core：資料 → 差異清單 =====================
    /**
     * @param {Object} data      collectData() 的結果（會讀 data.tables）
     * @param {Object} parsed    parseSheet() 的結果
     * @param {Object} profile
     */
    function mapData(data, parsed, profile) {
        var tables = (data && data.tables) || {};
        var valueKey = profile.valueKey;
        var ops = [], unmatched = [], ignored = [], missingTables = [];

        Object.keys(profile.funds).forEach(function (code) {
            var fund = parsed.funds[code];
            var officialName = profile.funds[code];
            if (!fund) return;
            ['src', 'use'].forEach(function (section) {
                var tableId = (profile.tableTemplates[section] || '').replace('{code}', code);
                var rows = tables[tableId];
                if (!Array.isArray(rows)) {
                    if (missingTables.indexOf(tableId) < 0) missingTables.push(tableId);
                    return;
                }
                var columns = parsed.columns[section] || [];
                var vals = fund[section] || {};

                rows.forEach(function (row, idx) {
                    var rawName = String((row && row.name) || '').trim();
                    if (!rawName) return;
                    // L0「甲、基金來源」等段落標題由 recalcPlanTable 自動加總，不覆寫
                    if ((parseInt(row.level, 10) || 0) === 0) return;
                    if ((profile.ignoreNames || []).indexOf(rawName) >= 0) {
                        ignored.push({ tableId: tableId, name: rawName });
                        return;
                    }
                    var res = resolveColumn(rawName, columns, profile);
                    if (!res.column) {
                        unmatched.push({
                            tableId: tableId, fund: officialName, name: rawName,
                            reason: res.how, best: res.best || null,
                            score: res.score == null ? null : Math.round(res.score * 1000) / 1000
                        });
                        return;
                    }
                    var value = vals[res.column];
                    if (value === undefined) {
                        unmatched.push({
                            tableId: tableId, fund: officialName, name: rawName,
                            reason: 'no-value', column: res.column
                        });
                        return;
                    }
                    var oldRaw = row[valueKey];
                    var oldNum = toNumber(oldRaw);
                    ops.push({
                        tableId: tableId, fund: officialName, index: idx, level: parseInt(row.level, 10) || 1,
                        name: rawName, column: res.column, how: res.how,
                        score: res.score == null ? null : Math.round(res.score * 1000) / 1000,
                        oldRaw: oldRaw == null ? '' : oldRaw, oldNum: oldNum,
                        value: value,
                        changed: oldNum === null ? value !== 0 : Math.abs(oldNum - value) > 1e-9
                    });
                });
            });
        });

        return {
            ops: ops,
            unmatched: unmatched,
            ignored: ignored,
            missingTables: missingTables,
            warnings: computeAutoSumWarnings(data, ops, profile),
            stats: summarize(ops, unmatched)
        };
    }

    function summarize(ops, unmatched) {
        var changed = ops.filter(function (o) { return o.changed; });
        var byHow = function (h) { return ops.filter(function (o) { return o.how === h; }).length; };
        return {
            matched: ops.length,
            changed: changed.length,
            confirmRequired: ops.filter(function (o) { return o.how !== 'exact' && o.how !== 'alias'; }).length,
            exact: byHow('exact'),
            alias: byHow('alias'),
            contains: byHow('contains'),
            fuzzy: byHow('fuzzy'),
            unmatched: unmatched.length,
            rounding: ops.filter(function (o) { return Math.abs(o.value - Math.round(o.value)) > 1e-9; }).length
        };
    }

    /**
     * 母列會被 recalcPlanTable 用**直屬子列**加總覆寫；若官方值與子列加總不一致，
     * 先在此提出警告（不會阻擋套用）。
     */
    function computeAutoSumWarnings(data, ops, profile) {
        var valueKey = profile.valueKey;
        var byTable = {};
        ops.forEach(function (o) {
            (byTable[o.tableId] = byTable[o.tableId] || []).push(o);
        });
        var warnings = [];
        Object.keys(byTable).forEach(function (tableId) {
            var rows = (data.tables && data.tables[tableId]) || [];
            var list = byTable[tableId];
            var opAt = {};
            list.forEach(function (o) { opAt[o.index] = o; });

            list.forEach(function (op) {
                // 只取「直屬子列」：與 recalcPlanTable 的 node.children 語意一致，
                // 否則像「徵收收入」這種中間層會被重複計算一次。
                var children = [];
                for (var i = op.index + 1; i < rows.length; i++) {
                    var lv = parseInt(rows[i] && rows[i].level, 10) || 0;
                    if (lv <= op.level) break;
                    if (lv !== op.level + 1) continue;
                    var childOp = opAt[i];
                    var v = childOp ? childOp.value : toNumber(rows[i] && rows[i][valueKey]);
                    if (v !== null && v !== undefined && v !== '') children.push(v);
                }
                if (!children.length) return;
                var sum = children.reduce(function (a, b) { return a + b; }, 0);
                if (Math.abs(sum - op.value) > 0.5) {
                    warnings.push({
                        tableId: tableId, name: op.name, official: op.value, childrenSum: sum
                    });
                }
            });
        });
        return warnings;
    }

    /** 把勾選的 ops 套進 data（回傳新物件，不動傳入的 data） */
    function applyOps(data, ops, profile) {
        var valueKey = profile.valueKey;
        var out = {
            meta: data.meta ? JSON.parse(JSON.stringify(data.meta)) : undefined,
            reviews: data.reviews ? JSON.parse(JSON.stringify(data.reviews)) : undefined,
            tables: {}
        };
        Object.keys(data.tables || {}).forEach(function (tid) {
            out.tables[tid] = (data.tables[tid] || []).map(function (row) { return Object.assign({}, row); });
        });
        ops.forEach(function (op) {
            var row = out.tables[op.tableId] && out.tables[op.tableId][op.index];
            if (!row) return;
            row[valueKey] = String(op.value);
        });
        return out;
    }

    function buildReport(parsed, result, profile, fileName) {
        return {
            profile: profile.id,
            profileLabel: profile.label,
            file: fileName || '',
            generatedAt: new Date().toISOString(),
            scale: profile.scale,
            valueKey: profile.valueKey,
            stats: result.stats,
            ops: result.ops.map(function (o) {
                return {
                    tableId: o.tableId, row: o.index, name: o.name, column: o.column,
                    how: o.how, score: o.score, old: o.oldRaw, value: o.value, changed: o.changed
                };
            }),
            unmatched: result.unmatched,
            ignored: result.ignored,
            warnings: result.warnings,
            missingTables: result.missingTables
        };
    }

    var core = {
        normalizeHeader: normalizeHeader,
        normalizeName: normalizeName,
        unify: unify,
        similarity: similarity,
        scaleValue: scaleValue,
        parseSheet: parseSheet,
        resolveColumn: resolveColumn,
        mapData: mapData,
        applyOps: applyOps,
        buildReport: buildReport
    };

    // ===================== UI（瀏覽器） =====================
    var state = null; // { parsed, result, fileName, data, profile, report, selected:Set }

    function fmt(n) {
        if (n === '' || n == null || isNaN(n)) return '—';
        return Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 });
    }

    function esc(s) {
        return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    function toast(msg) {
        if (typeof root.flashAutosave === 'function') root.flashAutosave(msg);
    }

    function tableTitle(tableId) {
        // tableConfigs 是 special-fund.js 的 const（全域語彙環境），不是 window 屬性，
        // 所以只能直接以變數名取用，並用 try/catch 容忍載入順序。
        try {
            /* global tableConfigs */
            if (typeof tableConfigs === 'object' && tableConfigs && tableConfigs[tableId]) {
                return tableConfigs[tableId].title || tableId;
            }
        } catch (e) { /* ignore */ }
        return tableId;
    }

    function ensureModal() {
        var el = document.getElementById('sf-xlsx-overlay');
        if (el) return el;
        el = document.createElement('div');
        el.id = 'sf-xlsx-overlay';
        el.className = 'sf-xlsx-overlay';
        el.hidden = true;
        el.innerHTML = ''
            + '<div class="sf-xlsx-panel" role="dialog" aria-modal="true" aria-labelledby="sf-xlsx-title">'
            + '  <div class="sf-xlsx-head">'
            + '    <div>'
            + '      <h3 id="sf-xlsx-title">匯入官方 Excel（差異預覽）</h3>'
            + '      <p class="sf-xlsx-sub" id="sf-xlsx-sub">只補值，不動表格結構</p>'
            + '    </div>'
            + '    <button type="button" class="sf-xlsx-close" id="sf-xlsx-close" aria-label="關閉">✕</button>'
            + '  </div>'
            + '  <div class="sf-xlsx-body">'
            + '    <div class="sf-xlsx-drop" id="sf-xlsx-drop">'
            + '      <div class="sf-xlsx-drop-main">📊 拖放或點選官方 <b>.xlsx</b> 檔</div>'
            + '      <div class="sf-xlsx-drop-hint">例：FNGBRB5300_*.xlsx（各基金業務計畫年度預算數）</div>'
            + '      <input type="file" id="sf-xlsx-file" accept=".xlsx,.xlsm" hidden>'
            + '    </div>'
            + '    <div class="sf-xlsx-summary" id="sf-xlsx-summary" hidden></div>'
            + '    <div class="sf-xlsx-toolbar" id="sf-xlsx-toolbar" hidden>'
            + '      <label class="sf-xlsx-chk"><input type="checkbox" id="sf-xlsx-only-changed" checked> 只顯示有變動的列（需確認項目一律顯示）</label>'
            + '      <button type="button" class="sf-xlsx-mini" id="sf-xlsx-check-confirm">勾選全部需確認項目</button>'
            + '      <button type="button" class="sf-xlsx-mini" id="sf-xlsx-check-none">全部取消</button>'
            + '    </div>'
            + '    <div class="sf-xlsx-list" id="sf-xlsx-list"></div>'
            + '  </div>'
            + '  <div class="sf-xlsx-foot">'
            + '    <span class="sf-xlsx-footnote">「完全相同」與 profile 內的「別名」預設勾選；包含／相似比對請逐筆確認。</span>'
            + '    <div class="sf-xlsx-actions">'
            + '      <button type="button" class="sf-xlsx-btn" id="sf-xlsx-report" disabled>⬇ 下載映射報告</button>'
            + '      <button type="button" class="sf-xlsx-btn" id="sf-xlsx-cancel">取消</button>'
            + '      <button type="button" class="sf-xlsx-btn primary" id="sf-xlsx-apply" disabled>套用至目前資料</button>'
            + '    </div>'
            + '  </div>'
            + '</div>';
        document.body.appendChild(el);

        el.querySelector('#sf-xlsx-close').addEventListener('click', close);
        el.querySelector('#sf-xlsx-cancel').addEventListener('click', close);
        el.addEventListener('click', function (ev) { if (ev.target === el) close(); });
        document.addEventListener('keydown', function (ev) {
            if (ev.key === 'Escape' && !el.hidden) close();
        });

        var drop = el.querySelector('#sf-xlsx-drop');
        var input = el.querySelector('#sf-xlsx-file');
        drop.addEventListener('click', function () { input.click(); });
        input.addEventListener('change', function () {
            if (input.files && input.files[0]) loadFile(input.files[0]);
            input.value = '';
        });
        ['dragenter', 'dragover'].forEach(function (t) {
            drop.addEventListener(t, function (ev) { ev.preventDefault(); drop.classList.add('hot'); });
        });
        ['dragleave', 'drop'].forEach(function (t) {
            drop.addEventListener(t, function (ev) { ev.preventDefault(); drop.classList.remove('hot'); });
        });
        drop.addEventListener('drop', function (ev) {
            var f = ev.dataTransfer && ev.dataTransfer.files && ev.dataTransfer.files[0];
            if (f) loadFile(f);
        });

        el.querySelector('#sf-xlsx-only-changed').addEventListener('change', renderList);
        el.querySelector('#sf-xlsx-check-confirm').addEventListener('click', function () {
            if (!state) return;
            state.result.ops.forEach(function (op, i) {
                if (op.how !== 'exact' && op.how !== 'alias') state.selected.add(i);
            });
            renderList();
        });
        el.querySelector('#sf-xlsx-check-none').addEventListener('click', function () {
            if (!state) return;
            state.selected.clear();
            renderList();
        });
        el.querySelector('#sf-xlsx-report').addEventListener('click', downloadReport);
        el.querySelector('#sf-xlsx-apply').addEventListener('click', applyNow);

        return el;
    }

    function close() {
        var el = document.getElementById('sf-xlsx-overlay');
        if (el) el.hidden = true;
        state = null;
    }

    function currentProfile() {
        var profiles = root.LGB_XLSX_PROFILES || {};
        var ids = Object.keys(profiles);
        if (!ids.length) throw new Error('找不到映射 profile（xlsx-profiles.js 未載入）');
        return profiles[ids[0]];
    }

    function readWorkbook(file, cb) {
        if (typeof root.XLSX === 'undefined') {
            toast('⚠ 需要網路載入 Excel 解析元件（SheetJS）');
            return;
        }
        var reader = new FileReader();
        reader.onload = function () {
            try {
                var wb = root.XLSX.read(new Uint8Array(reader.result), { type: 'array' });
                var profile = currentProfile();
                var sheetName = typeof profile.sheet === 'number'
                    ? wb.SheetNames[profile.sheet]
                    : profile.sheet;
                var ws = wb.Sheets[sheetName];
                if (!ws) throw new Error('找不到工作表：' + sheetName);
                var aoa = root.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
                cb(aoa, profile, file.name, sheetName);
            } catch (e) {
                toast('⚠ 解析失敗：' + e.message);
            }
        };
        reader.onerror = function () { toast('⚠ 讀取檔案失敗'); };
        reader.readAsArrayBuffer(file);
    }

    function loadFile(file) {
        if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
            toast('⚠ 只支援 .xlsx / .xlsm');
            return;
        }
        readWorkbook(file, function (aoa, profile, fileName, sheetName) {
            var parsed = core.parseSheet(aoa, profile);
            var data = typeof root.collectData === 'function' ? root.collectData() : { tables: {} };
            var result = core.mapData(data, parsed, profile);
            state = {
                parsed: parsed, result: result, data: data, profile: profile,
                fileName: fileName, sheetName: sheetName,
                report: core.buildReport(parsed, result, profile, fileName),
                selected: new Set()
            };
            result.ops.forEach(function (op, i) {
                if ((op.how === 'exact' || op.how === 'alias') && op.changed) state.selected.add(i);
            });
            renderSummary();
            renderList();
            var btn = document.getElementById('sf-xlsx-report');
            if (btn) btn.disabled = false;
        });
    }

    function renderSummary() {
        var el = document.getElementById('sf-xlsx-summary');
        var s = state.result.stats;
        var bits = [];
        bits.push('<b>' + esc(state.fileName) + '</b>');
        bits.push('工作表：' + esc(state.sheetName));
        bits.push('命中 ' + s.matched + ' 列');
        bits.push('變動 <b class="' + (s.changed ? 'warn' : 'ok') + '">' + s.changed + '</b>');
        if (s.alias) bits.push('別名 ' + s.alias);
        bits.push('需確認 ' + s.confirmRequired);
        bits.push('未匹配 ' + s.unmatched);
        if (s.rounding) bits.push('含小數 ' + s.rounding);
        var html = '<div class="sf-xlsx-summary-line">' + bits.join('　·　') + '</div>';
        if (state.result.missingTables.length) {
            html += '<div class="sf-xlsx-warn">缺少表格：' + state.result.missingTables.map(esc).join('、') + '</div>';
        }
        state.result.warnings.slice(0, 5).forEach(function (w) {
            html += '<div class="sf-xlsx-warn">⚠ ' + esc(tableTitle(w.tableId)) + '／' + esc(w.name) +
                '：官方 ' + fmt(w.official) + '，子列加總 ' + fmt(w.childrenSum) + '（套用後以子列加總為準）</div>';
        });
        el.innerHTML = html;
        el.hidden = false;
        document.getElementById('sf-xlsx-toolbar').hidden = false;
    }

    function renderList() {
        var box = document.getElementById('sf-xlsx-list');
        if (!state) { box.innerHTML = ''; return; }
        var onlyChanged = document.getElementById('sf-xlsx-only-changed').checked;
        var ops = state.result.ops;
        var groups = {};
        ops.forEach(function (op, i) {
            if (onlyChanged && !op.changed && (op.how === 'exact' || op.how === 'alias')) return;
            (groups[op.tableId] = groups[op.tableId] || []).push({ op: op, i: i });
        });

        var html = '';
        Object.keys(groups).forEach(function (tid) {
            html += '<div class="sf-xlsx-group"><div class="sf-xlsx-group-title">' + esc(tableTitle(tid)) +
                '<span class="sf-xlsx-group-count">' + groups[tid].length + ' 列</span></div>';
            groups[tid].forEach(function (item) {
                var op = item.op;
                var checked = state.selected.has(item.i);
                var cls = 'sf-xlsx-row' + (op.changed ? ' changed' : '') + (op.how !== 'exact' ? ' confirm' : '');
                html += '<label class="' + cls + '">'
                    + '<input type="checkbox" data-op="' + item.i + '"' + (checked ? ' checked' : '') + '>'
                    + '<span class="sf-xlsx-name">' + esc(op.name) + '</span>'
                    + '<span class="sf-xlsx-col">' + esc(op.column) + '</span>'
                    + '<span class="sf-xlsx-old">' + fmt(op.oldNum) + '</span>'
                    + '<span class="sf-xlsx-arrow">→</span>'
                    + '<span class="sf-xlsx-new">' + fmt(op.value) + '</span>'
                    + '<span class="sf-xlsx-tag ' + op.how + '">' + tagLabel(op.how) + '</span>'
                    + '</label>';
            });
            html += '</div>';
        });

        if (state.result.unmatched.length) {
            html += '<div class="sf-xlsx-group"><div class="sf-xlsx-group-title">未匹配（不會寫入）' +
                '<span class="sf-xlsx-group-count">' + state.result.unmatched.length + ' 列</span></div>';
            state.result.unmatched.forEach(function (u) {
                html += '<div class="sf-xlsx-row unmatched">'
                    + '<span class="sf-xlsx-name">' + esc(u.name) + '</span>'
                    + '<span class="sf-xlsx-col">' + (u.best ? '最接近：' + esc(u.best) + '（' + u.score + '）' : reasonLabel(u.reason)) + '</span>'
                    + '</div>';
            });
            html += '</div>';
        }
        box.innerHTML = html || '<div class="sf-xlsx-empty">沒有可顯示的列</div>';

        box.querySelectorAll('input[data-op]').forEach(function (inp) {
            inp.addEventListener('change', function () {
                var i = Number(inp.getAttribute('data-op'));
                if (inp.checked) state.selected.add(i); else state.selected.delete(i);
                updateApplyState();
            });
        });
        updateApplyState();
    }

    function tagLabel(how) {
        return { exact: '完全相同', alias: '別名', contains: '包含', fuzzy: '相似' }[how] || how;
    }

    function reasonLabel(reason) {
        return {
            'too-short': '名稱過短，不自動比對',
            'no-value': '官方檔在此基金無此欄位值',
            'unmatched': '找不到對應欄位'
        }[reason] || '未匹配';
    }

    function updateApplyState() {
        if (!state) return;
        var btn = document.getElementById('sf-xlsx-apply');
        var n = state.selected.size;
        btn.disabled = n === 0;
        btn.textContent = n ? '套用 ' + n + ' 筆' : '套用至目前資料';
    }

    function downloadReport() {
        if (!state) return;
        var report = core.buildReport(state.parsed, state.result, state.profile, state.fileName);
        var blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'mapping-report_' + (report.file || 'xlsx').replace(/\.[^.]+$/, '') + '_' +
            new Date().toISOString().slice(0, 19).replace(/[:T]/g, '') + '.json';
        a.click();
        setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    }

    function applyNow() {
        if (!state) return;
        var ops = state.result.ops.filter(function (op, i) { return state.selected.has(i); });
        if (!ops.length) return;
        var next = core.applyOps(state.data, ops, state.profile);
        if (typeof root.applyData !== 'function') { toast('⚠ 找不到 applyData()'); return; }
        root.applyData(next);
        if (typeof root.scheduleAutosave === 'function') root.scheduleAutosave();
        toast('✓ 已套用 ' + ops.length + ' 筆（' + state.fileName + '）');
        close();
    }

    function open() {
        var el = ensureModal();
        el.hidden = false;
    }

    function init() {
        var btn = document.getElementById('sf-btn-xlsx-import');
        if (!btn) return;
        var input = document.getElementById('sf-xlsx-import-file');
        btn.addEventListener('click', function () { open(); });
        if (input) {
            input.addEventListener('change', function () {
                if (input.files && input.files[0]) {
                    open();
                    loadFile(input.files[0]);
                }
                input.value = '';
            });
        }
    }

    if (typeof document !== 'undefined') {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', init);
        } else {
            init();
        }
    }

    return { core: core, open: open, close: close };
});
