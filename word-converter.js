'use strict';

// ============================================================
// word_table_to_json_converter 輸出 → 本專案結構化格式
// ============================================================
// 來源：使用者用 word_table_to_json_converter.html 從 Word 轉出的「原始 2D 陣列 JSON」
//   形如 [{ tableIndex, tableName, rowCount, data: [[...row...], ...] }, ...]
//
// 規則（user 確認）：
//  1. 第 1～2 列略過（為 thead，mammoth 把合併儲存格展平後通常佔 2 列）
//  2. 各列**依欄位位置**對應到結構化 key（不看表頭文字，避免年度錯位）：
//        cells[0] → dec113     （第 1 個數字欄）
//        cells[1] → bud114     （第 2 個數字欄）
//        cells[2] → name
//        cells[3] → desc
//        cells[4] → orig
//        cells[5] → dep_diff
//        cells[6] → dep_app
//        cells[7] → gov_diff
//        cells[8] → gov_app
//  3. 「壹、貳、參/叁、肆、伍、陸」整列代表新分基金的開始
//  4. 「甲、...」啟用 src 段，「乙、...」啟用 use 段
//  5. 層級依文字前綴推測；無前綴者預設為 L2「計畫名」，使用者可在 UI 內調整
//  6. Table 2+ 若有「主管機關」「先期審查機關」標籤列，視為 reviews.plan

const WORD_CONVERTER_FUND_MAP = {
    '壹、農業發展基金':            'plan_agri',
    '貳、林務發展及造林基金':       'plan_forest',
    '參、農業天然災害救助基金':     'plan_disaster',
    '叁、農業天然災害救助基金':     'plan_disaster',
    '肆、漁業發展基金':             'plan_fish',
    '伍、農產品受進口損害救助基金': 'plan_loss',
    '陸、農村再生基金':             'plan_renewal',
};

function _wcInferPlanLevel(name) {
    if (!name) return 2;
    if (/^甲、|^乙、/.test(name)) return 0;
    if (/^[一二三四五六七八九十百]+、/.test(name)) return 1;
    if (/^[(（][一二三四五六七八九十百]+[)）]/.test(name)) return 2;
    if (/^\d+\./.test(name)) return 3;
    if (/^[(（]\d+[)）]/.test(name)) return 4;
    return 2; // 預設「計畫名」
}

function _wcParseNum(s) {
    if (s == null) return '';
    const t = String(s).trim();
    if (t === '' || t === '-') return '';
    const n = parseFloat(t.replace(/[,\s]/g, ''));
    return isNaN(n) ? '' : String(n);
}

function _wcEnsureSectionHeader(rows, suffix) {
    const expected = suffix === 'src' ? '甲' : '乙';
    const label    = suffix === 'src' ? '甲、基金來源' : '乙、基金用途';
    if (!rows.length || !(rows[0].name || '').trim().startsWith(expected)) {
        return [{ level: 0, name: label + '：' }, ...rows];
    }
    // 強制首列為 L0
    if (parseInt(rows[0].level) !== 0) rows[0].level = 0;
    return rows;
}

// 主入口：解析整份 converter 輸出
// 回傳：{
//   funds: { plan_agri: { src: [...], use: [...] }, plan_forest: {...}, ... },
//   reviews: { plan: { org, gov } },
//   counts:  { plan_agri: N, ... },
//   warnings: ['...']
// }
function parseWordConverterJSON(raw) {
    if (!Array.isArray(raw)) throw new Error('輸入不是 converter 陣列格式（應為 [{ tableIndex, data }, ...]）');
    const mainTable = raw.find(t => t && t.tableIndex === 1) || raw[0];
    if (!mainTable || !Array.isArray(mainTable.data)) throw new Error('找不到主表格（Table 1）的 data 陣列');

    const out = {
        funds: {
            plan_agri:     { src: [], use: [] },
            plan_forest:   { src: [], use: [] },
            plan_disaster: { src: [], use: [] },
            plan_fish:     { src: [], use: [] },
            plan_loss:     { src: [], use: [] },
            plan_renewal:  { src: [], use: [] },
        },
        reviews: { plan: { org: '', gov: '' } },
        counts: {},
        warnings: []
    };

    let currentFund = null;
    let currentSection = null; // 'src' | 'use'

    for (let i = 2; i < mainTable.data.length; i++) {
        const cells = (mainTable.data[i] || []).map(c => String(c ?? '').trim());
        if (!cells.length) continue;
        const name = cells[2] || '';
        if (!name) continue;

        // 分基金標題列
        const fundId = WORD_CONVERTER_FUND_MAP[name];
        if (fundId) {
            currentFund = fundId;
            currentSection = null;
            continue;
        }
        if (!currentFund) {
            out.warnings.push(`第 ${i + 1} 列出現在任何基金之前，已略過：「${name}」`);
            continue;
        }

        // 甲/乙 段落切換
        if (/^甲、/.test(name)) currentSection = 'src';
        else if (/^乙、/.test(name)) currentSection = 'use';
        if (!currentSection) {
            out.warnings.push(`第 ${i + 1} 列在「甲、」之前出現於 ${currentFund}，已略過：「${name}」`);
            continue;
        }

        const row = {
            level: _wcInferPlanLevel(name),
            name: name,
            desc: cells[3] || '',
            dec113:   _wcParseNum(cells[0]),
            bud114:   _wcParseNum(cells[1]),
            orig:     _wcParseNum(cells[4]),
            dep_diff: _wcParseNum(cells[5]),
            dep_app:  _wcParseNum(cells[6]),
            gov_diff: _wcParseNum(cells[7]),
            gov_app:  _wcParseNum(cells[8]),
        };
        out.funds[currentFund][currentSection].push(row);
    }

    // 從 Table 2+ 擷取審核意見
    for (let t = 1; t < raw.length; t++) {
        const rows = raw[t]?.data || [];
        for (const r of rows) {
            const lbl = String(r?.[0] ?? '').trim();
            const val = String(r?.[1] ?? '').trim();
            if (lbl === '主管機關' && val) out.reviews.plan.org = val;
            else if (lbl === '先期審查機關' && val) out.reviews.plan.gov = val;
        }
    }

    // 強制每段最上方為 L0 標題列、統計列數
    Object.entries(out.funds).forEach(([fid, secs]) => {
        secs.src = _wcEnsureSectionHeader(secs.src, 'src');
        secs.use = _wcEnsureSectionHeader(secs.use, 'use');
        // 計數時扣掉 L0 標題列（純資料列才算）
        const cnt = (rows) => rows.filter(r => parseInt(r.level) !== 0).length;
        out.counts[fid] = cnt(secs.src) + cnt(secs.use);
    });

    return out;
}

// ============================================================
// .docx 直接解析（需 mammoth.js 已載入）
// ============================================================
// 行為與 word_table_to_json_converter.html 一致：
//   mammoth.convertToHtml → DOMParser 抓 <table> → 取每格 textContent
// 再餵給 parseWordConverterJSON 取得結構化資料。
async function parseDocxFile(file) {
    if (typeof mammoth === 'undefined' || !mammoth.convertToHtml) {
        throw new Error('mammoth.js 未載入，無法解析 .docx');
    }
    if (!file || !(file instanceof Blob)) {
        throw new Error('需要傳入 File / Blob');
    }
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    const parser = new DOMParser();
    const doc = parser.parseFromString(result.value || '', 'text/html');
    const tables = doc.querySelectorAll('table');
    if (!tables.length) {
        throw new Error('文件中沒有偵測到任何表格');
    }
    const raw = [];
    tables.forEach((tbl, idx) => {
        const rows = Array.from(tbl.querySelectorAll('tr')).map(tr =>
            Array.from(tr.querySelectorAll('th, td')).map(c => {
                const text = c.innerText != null ? c.innerText : (c.textContent || '');
                return String(text).trim();
            })
        ).filter(r => r.length > 0);
        raw.push({
            tableIndex: idx + 1,
            tableName: `表格 ${idx + 1}`,
            rowCount: rows.length,
            data: rows
        });
    });
    const parsed = parseWordConverterJSON(raw);
    // 附帶原始 mammoth messages（若有）方便除錯
    parsed.docxMessages = result.messages || [];
    return parsed;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = { parseWordConverterJSON, parseDocxFile, WORD_CONVERTER_FUND_MAP };
}
