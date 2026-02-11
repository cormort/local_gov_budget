'use strict';

// ========== 1. 全域配置與欄位定義 ==========
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

// ========== 2. 靜態備份功能 (含列印按鈕與日期) ==========
function mgr_exportHTML() {
    try {
        // 同步 input 值
        document.querySelectorAll('input').forEach(i => i.setAttribute('value', i.value));
        
        let inlineStyle = "";
        try {
            for (let sheet of document.styleSheets) {
                if (sheet.href && (sheet.href.includes('budget-style.css') || sheet.href.includes('input.css'))) {
                    const rules = sheet.rules || sheet.cssRules;
                    inlineStyle += Array.from(rules).map(r => r.cssText).join("\n");
                }
            }
        } catch (e) { console.warn("CSS 抓取受限"); }

        let cloneDoc = document.documentElement.cloneNode(true);
        
        // A. 移除所有功能性元件
        cloneDoc.querySelector('nav')?.remove();
        cloneDoc.getElementById('tab-aggregator')?.remove();
        cloneDoc.querySelectorAll('.flex.gap-2, #btn-clear, .excel-guide, script').forEach(el => el.remove());
        cloneDoc.querySelectorAll('.add-row-btn, .delete-btn, #autosave-indicator, #undo-btn').forEach(el => el.remove());

        // B. 轉換 input 為純文字
        cloneDoc.querySelectorAll('input').forEach(input => {
            const span = document.createElement('span');
            span.textContent = input.value || '';
            span.className = input.className + " inline-block";
            input.parentNode.replaceChild(span, input);
        });

        // C. 加入列印控制區 (按鈕與日期)
        const printHeader = document.createElement('div');
        const now = new Date();
        const dateStr = `${now.getFullYear()-1911}年${now.getMonth()+1}月${now.getDate()}日 ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}`;
        
        printHeader.className = "max-w-7xl mx-auto mb-6 flex justify-between items-end border-b pb-4 no-print";
        printHeader.innerHTML = `
            <div>
                <button onclick="window.print()" style="background:#2563eb; color:white; padding:8px 20px; border-radius:6px; font-weight:bold; cursor:pointer; border:none;">🖨️ 列印此報表</button>
                <p style="font-size:12px; color:#64748b; margin-top:8px;">提示：此為靜態備份檔，僅供檢視與列印。</p>
            </div>
            <div style="text-align:right; color:#64748b; font-size:14px;">
                產製日期：${dateStr}
            </div>
        `;
        const container = cloneDoc.getElementById('tab-manager');
        container.prepend(printHeader);
        container.style.marginTop = "20px";

        // D. 注入 CSS (包含列印隱藏邏輯)
        cloneDoc.querySelectorAll('link[href*="css"]').forEach(l => { if(!l.href.includes('fonts')) l.remove(); });
        const styleTag = document.createElement('style');
        styleTag.textContent = inlineStyle + `
            @media print { .no-print { display: none !important; } body { background: white; } .section-card { border: 1px solid #eee; break-inside: avoid; } }
            body { background: #f8fafc; padding-bottom: 50px; }
            .section-card { box-shadow: none !important; margin-bottom: 30px; }
            span.negative-value { color: #dc2626; font-weight: bold; }
        `;
        cloneDoc.querySelector('head').appendChild(styleTag);

        const htmlContent = "<!DOCTYPE html>\n" + cloneDoc.outerHTML;
        const org = document.getElementById('mgr-org').value || '預算報表';
        saveAs(new Blob([htmlContent], { type: "text/html" }), `靜態報表_${org}.html`);
    } catch (err) { alert('匯出失敗：' + err.message); }
}

// ========== 3. 匯整端邏輯 (參考 index 41) ==========
let agg_data = [];
function agg_processFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const doc = new DOMParser().parseFromString(e.target.result, 'text/html');
            const data = {
                metadata: { 
                    org: doc.getElementById('mgr-org')?.value || file.name.replace('.html',''),
                    year: doc.getElementById('mgr-year')?.value || '115',
                    user: doc.getElementById('mgr-user')?.value || '未知'
                },
                sections: sectionConfigs.map(conf => ({
                    id: conf.id,
                    items: Array.from(doc.querySelectorAll(`#tbody-${conf.id} tr`)).map(tr => {
                        let item = {};
                        conf.fields.forEach(f => {
                            const inp = tr.querySelector('.v-'+f) || tr.querySelector('.v-'+f.replace('v-',''));
                            item[f] = inp?.getAttribute('value') || inp?.textContent || '';
                        });
                        return item;
                    }).filter(i => i.name)
                }))
            };
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
    const num = v => parseFloat(String(v).replace(/,/g,'')) || 0;

    agg_data.forEach(gov => {
        gov.sections?.forEach(sec => {
            sec.items?.forEach(item => {
                stats.funds++;
                let rev = (sec.id === 'op' || sec.id === 'wk') ? (num(item.rev) + num(item.nonrev)) : num(item.source);
                let bal = num(item.net) || num(item.surplus) || (num(item.end) - num(item.begin));
                stats.totalRev += rev;
                if (bal >= 0) stats.profit++; else stats.loss++;
            });
        });
    });

    document.getElementById('agg-kpi').innerHTML = `
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>機關數</div><div class="text-2xl font-bold text-blue-400">${stats.govs}</div></div>
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>基金數</div><div class="text-2xl font-bold text-green-400">${stats.funds}</div></div>
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>總規模(億)</div><div class="text-2xl font-bold text-emerald-400">${(stats.totalRev / 100000).toFixed(2)}</div></div>
        <div class="kpi-card bg-slate-800 p-4 rounded-lg"><div>盈虧分佈</div><div class="text-sm">盈: ${stats.profit} / 虧: ${stats.loss}</div></div>
    `;

    document.getElementById('agg-list-body').innerHTML = agg_data.map((d, i) => `
        <tr class="border-b border-slate-700">
            <td class="p-3 text-slate-500">${i+1}</td>
            <td class="p-3 font-bold text-blue-300">${d.metadata.org}</td>
            <td class="p-3 text-sm text-slate-400">${d.metadata.year}年 / ${d.metadata.user}</td>
            <td class="p-3 text-right"><button class="text-red-400 text-sm" onclick="agg_remove(${i})">移除</button></td>
        </tr>
    `).join('');
}

window.agg_remove = (idx) => { agg_data.splice(idx,1); agg_render(); };

// ========== 4. 填報端核心計算與介面 ==========
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
                <button class="add-row-btn text-blue-600 font-bold hover:underline" data-section="${conf.id}">+ 新增基金</button>
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
            row.querySelector('.v-'+f)?.classList.toggle('negative-value', val < 0);
            totals[f] += val;
        });
    });
    
    Object.keys(totals).forEach(f => {
        const tEl = document.querySelector(`#tfoot-${type} .t-${f}`);
        if (tEl) {
            tEl.value = totals[f].toLocaleString();
            tEl.classList.toggle('negative-value', totals[f] < 0);
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
}

document.addEventListener('DOMContentLoaded', () => {
    render();
    bindEvents();
    sectionConfigs.forEach(c => mgr_addRow(c.id)); 
});
