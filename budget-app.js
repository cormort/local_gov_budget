'use strict';

// ========== 1. 全域配置與欄位定義 (對應您要求的專業名稱) ==========
const sectionConfigs = [
    { id: 'op', title: '一、營業基金', color: '#2563eb', fields: ['name','rev','cost','gross','exp','opprofit','nonrev','nonexp','nonprofit','pretax','tax','net'] },
    { id: 'wk', title: '二、作業基金', color: '#16a34a', fields: ['name','rev','cost','surplus','nonrev','nonexp','nonsurplus','net'] },
    { id: 'db', title: '三、債務基金', color: '#ea580c', fields: ['name','source','use','surplus','begin','remit','end'] },
    { id: 'sp', title: '四、特別收入基金', color: '#9333ea', fields: ['name','source','use','surplus','begin','remit','end'] },
    { id: 'cp', title: '五、資本計畫基金', color: '#0891b2', fields: ['name','source','use','surplus','begin','remit','end'] }
];

const fieldNames = {
    name: '基金名稱', rev: '收入', cost: '成本/費用', gross: '營業毛利(毛損)', exp: '營業費用',
    opprofit: '營業利益(損失)', nonrev: '外收入', nonexp: '外費用', nonprofit: '外利益(損失)',
    pretax: '稅前淨利(淨損)', tax: '所得稅費用(利益)', net: '本期淨利(淨損)', 
    surplus: '賸餘(短絀)', nonsurplus: '外賸餘(短絀)', source: '基金來源', use: '基金用途', 
    begin: '期初基金餘額', remit: '解繳公庫', end: '期末基金餘額'
};

// 輔助函式：根據不同類別切換精確的標題文字
function getFieldLabel(sectionId, fieldId) {
    if (sectionId === 'op') {
        const labels = { rev: '營業收入', cost: '營業成本', gross: '營業毛利(毛損)', exp: '營業費用', opprofit: '營業利益(損失)', nonrev: '營業外收入', nonexp: '營業外費用', nonprofit: '營業外利益(損失)', pretax: '稅前淨利(淨損)', tax: '所得稅費用(利益)', net: '本期淨利(淨損)' };
        return labels[fieldId] || fieldNames[fieldId];
    }
    if (sectionId === 'wk') {
        const labels = { rev: '業務收入', cost: '業務成本與費用', surplus: '業務賸餘(短絀)', nonrev: '業務外收入', nonexp: '業務外費用', nonsurplus: '業務外賸餘(短絀)', net: '本期賸餘(短絀)' };
        return labels[fieldId] || fieldNames[fieldId];
    }
    if (['db', 'sp', 'cp'].includes(sectionId) && fieldId === 'surplus') return '本期賸餘(短絀)';
    return fieldNames[fieldId];
}

// ========== 2. 匯出備份 (修正 CSS 跑掉的問題) ==========
function mgr_exportHTML() {
    try {
        // 先將所有 input 的值同步到 DOM 的 value 屬性，確保備份檔靜態呈現正確
        document.querySelectorAll('input').forEach(i => i.setAttribute('value', i.value));
        
        // 抓取 budget-style.css 的內容
        let inlineStyle = "";
        try {
            for (let sheet of document.styleSheets) {
                // 檢查是否為 budget-style.css (過濾掉 Google Fonts 等外部連結)
                if (sheet.href && (sheet.href.includes('budget-style.css') || sheet.href.includes('input.css'))) {
                    const rules = sheet.rules || sheet.cssRules;
                    inlineStyle += Array.from(rules).map(r => r.cssText).join("\n");
                }
            }
        } catch (e) { console.warn("CSS 抓取受限，採用基礎排版備份"); }

        let cloneDoc = document.documentElement.cloneNode(true);
        
        // 移除原有的外部 CSS 引用，改為內嵌 Style 標籤
        cloneDoc.querySelectorAll('link[href*="css"]').forEach(l => {
            if(!l.href.includes('fonts')) l.remove();
        });
        
        if (inlineStyle) {
            const styleTag = document.createElement('style');
            styleTag.textContent = inlineStyle;
            cloneDoc.querySelector('head').appendChild(styleTag);
        }

        const htmlContent = "<!DOCTYPE html>\n" + cloneDoc.outerHTML;
        const blob = new Blob([htmlContent], { type: "text/html" });
        const org = document.getElementById('mgr-org').value || '預算備份';
        saveAs(blob, `${org}.html`);
    } catch (err) { alert('匯出 HTML 失敗：' + err.message); }
}

// ========== 3. 匯整端邏輯 (參考 index (41).html) ==========
let agg_data = [];

function agg_processFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            let data;
            if (file.name.endsWith('.json')) {
                data = JSON.parse(e.target.result);
            } else {
                const doc = new DOMParser().parseFromString(e.target.result, 'text/html');
                data = {
                    metadata: { 
                        org: doc.getElementById('mgr-org')?.value || file.name.replace('.html',''),
                        year: doc.getElementById('mgr-year')?.value || '',
                        user: doc.getElementById('mgr-user')?.value || ''
                    },
                    sections: sectionConfigs.map(conf => ({
                        id: conf.id,
                        items: Array.from(doc.querySelectorAll(`#tbody-${conf.id} tr`)).map(tr => {
                            let item = {};
                            conf.fields.forEach(f => {
                                const inp = tr.querySelector('.v-'+f);
                                item[f] = inp?.getAttribute('value') || inp?.value || '';
                            });
                            return item;
                        }).filter(i => i.name)
                    }))
                };
            }
            agg_data.push(data);
            agg_render();
        } catch (err) { alert('檔案解析失敗'); }
    };
    reader.readAsText(file);
}

function agg_render() {
    const container = document.getElementById('agg-content');
    if (!agg_data.length) { container.classList.add('hidden'); return; }
    container.classList.remove('hidden');

    let stats = { govs: agg_data.length, funds: 0, totalRev: 0, profit: 0, loss: 0 };
    const num = v => parseFloat(v) || 0;

    agg_data.forEach(gov => {
        gov.sections?.forEach(sec => {
            sec.items?.forEach(item => {
                stats.funds++;
                // 參考 index(41) 的全口徑總額計算
                let rev = (sec.id === 'op' || sec.id === 'wk') ? (num(item.rev) + num(item.nonrev)) : num(item.source);
                let bal = num(item.net) || num(item.surplus) || (num(item.end) - num(item.begin));
                stats.totalRev += rev;
                if (bal >= 0) stats.profit++; else stats.loss++;
            });
        });
    });

    document.getElementById('agg-kpi').innerHTML = `
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>機關總數</div><div class="text-2xl font-bold text-blue-400">${stats.govs}</div></div>
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>基金總數</div><div class="text-2xl font-bold text-green-400">${stats.funds}</div></div>
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>總收入規模(億)</div><div class="text-2xl font-bold text-emerald-400">${(stats.totalRev / 100000).toFixed(2)}</div></div>
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>盈虧分佈</div><div class="text-sm">盈: ${stats.profit} / 虧: ${stats.loss}</div></div>
    `;

    document.getElementById('agg-list-body').innerHTML = agg_data.map((d, i) => `
        <tr class="border-b border-slate-700">
            <td class="p-3 text-slate-500">${i+1}</td>
            <td class="p-3 font-bold text-blue-300">${d.metadata.org}</td>
            <td class="p-3 text-sm">${d.metadata.year}年 / ${d.metadata.user}</td>
            <td class="p-3 text-right"><button class="text-red-400 hover:underline" onclick="agg_remove(${i})">移除</button></td>
        </tr>
    `).join('');
}

window.agg_remove = (idx) => { agg_data.splice(idx,1); agg_render(); };

// ========== 4. 填報端核心計算與自動化 ==========
function render() {
    const container = document.getElementById('sections-container');
    container.innerHTML = '';
    sectionConfigs.forEach(conf => {
        const div = document.createElement('div');
        div.className = 'section-card bg-white rounded-xl shadow-sm border-t-4 p-6 mb-6';
        div.style.borderTopColor = conf.color;
        div.innerHTML = `
            <div class="flex justify-between items-center mb-4">
                <h3 class="font-bold text-lg" style="color:${conf.color}">${conf.title}</h3>
                <button class="add-row-btn text-blue-600 font-bold hover:underline" data-section="${conf.id}">+ 新增</button>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full budget-table text-sm">
                    <thead><tr><th class="w-8"></th>${conf.fields.map(f => `<th>${getFieldLabel(conf.id, f)}</th>`).join('')}</tr></thead>
                    <tbody id="tbody-${conf.id}"></tbody>
                    <tfoot id="tfoot-${conf.id}" class="bg-slate-50 font-bold">
                        <tr><td>∑</td><td class="text-center">合計</td>${conf.fields.slice(1).map(f => `<td><input type="text" class="t-${f}" readonly value="0"></td>`).join('')}</tr>
                    </tfoot>
                </table>
            </div>`;
        container.appendChild(div);
    });
}

function mgr_addRow(type, data = {}) {
    const conf = sectionConfigs.find(c => c.id === type);
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="text-center"><button class="delete-btn text-red-400" data-type="${type}">✕</button></td>` +
        conf.fields.map(f => {
            const isRO = ['gross','opprofit','nonprofit','pretax','net','surplus','nonsurplus','end'].includes(f);
            return `<td><input type="${f==='name'?'text':'number'}" class="v-${f} ${f==='name'?'text-left':''}" value="${data[f]||''}" ${isRO?'readonly':''}></td>`;
        }).join('');
    document.getElementById(`tbody-${type}`).appendChild(tr);
}

function update(type) {
    const rows = document.querySelectorAll(`#tbody-${type} tr`);
    const conf = sectionConfigs.find(c => c.id === type);
    let totals = {}; conf.fields.slice(1).forEach(f => totals[f] = 0);
    
    rows.forEach(row => {
        const v = (c) => parseFloat(row.querySelector('.v-'+c)?.value) || 0;
        const r = (c) => row.querySelector('.v-'+c);
        
        if (type === 'op') {
            let g = v('rev') - v('cost'), op = g - v('exp'), np = v('nonrev') - v('nonexp');
            if(r('gross')) r('gross').value = g; if(r('opprofit')) r('opprofit').value = op;
            if(r('nonprofit')) r('nonprofit').value = np; if(r('pretax')) r('pretax').value = op+np;
            if(r('net')) r('net').value = op+np-v('tax');
        } else if (type === 'wk') {
            let s = v('rev') - v('cost'), ns = v('nonrev') - v('nonexp');
            if(r('surplus')) r('surplus').value = s; if(r('nonsurplus')) r('nonsurplus').value = ns;
            if(r('net')) r('net').value = s + ns;
        } else {
            let s = v('source') - v('use');
            if(r('surplus')) r('surplus').value = s; if(r('end')) r('end').value = v('begin')+s-v('remit');
        }
        
        conf.fields.slice(1).forEach(f => {
            let val = v(f);
            row.querySelector('.v-'+f)?.classList.toggle('text-red-600', val < 0);
            totals[f] += val;
        });
    });
    
    Object.keys(totals).forEach(f => {
        const tEl = document.querySelector(`#tfoot-${type} .t-${f}`);
        if (tEl) {
            tEl.value = totals[f].toLocaleString();
            tEl.classList.toggle('text-red-600', totals[f] < 0);
        }
    });
}

// ========== 5. 事件與初始化 ==========
function bindEvents() {
    document.getElementById('btn-manager').onclick = () => { document.getElementById('tab-manager').classList.remove('hidden'); document.getElementById('tab-aggregator').classList.add('hidden'); };
    document.getElementById('btn-aggregator').onclick = () => { document.getElementById('tab-manager').classList.add('hidden'); document.getElementById('tab-aggregator').classList.remove('hidden'); };
    document.getElementById('btn-export-html').onclick = mgr_exportHTML;
    document.getElementById('btn-agg-clear').onclick = () => { agg_data = []; agg_render(); };
    
    document.getElementById('sections-container').onclick = e => {
        if (e.target.classList.contains('add-row-btn')) mgr_addRow(e.target.dataset.section);
        if (e.target.classList.contains('delete-btn')) { e.target.closest('tr').remove(); update(e.target.dataset.type); }
    };

    document.getElementById('sections-container').oninput = e => {
        const tbody = e.target.closest('tbody');
        if (tbody) update(tbody.id.replace('tbody-', ''));
    };

    const dz = document.getElementById('agg-dropzone');
    dz.onclick = () => {
        const inp = document.createElement('input'); inp.type = 'file'; inp.multiple = true;
        inp.onchange = e => Array.from(e.target.files).forEach(f => agg_processFile(f));
        inp.click();
    };
    dz.ondragover = e => { e.preventDefault(); dz.classList.add('bg-slate-700'); };
    dz.ondrop = e => { e.preventDefault(); dz.classList.remove('bg-slate-700'); Array.from(e.dataTransfer.files).forEach(f => agg_processFile(f)); };
}

document.addEventListener('DOMContentLoaded', () => {
    render();
    bindEvents();
    sectionConfigs.forEach(c => mgr_addRow(c.id)); 
});
