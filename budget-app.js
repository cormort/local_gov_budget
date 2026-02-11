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

// ========== 2. 匯入功能 (修正：相容舊版結構) ==========
function mgr_handleImport(files) {
    const file = files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const content = e.target.result;
            if (content.trim().startsWith('{')) {
                const data = JSON.parse(content);
                mgr_populate(data);
                alert('JSON 匯入成功！');
            } else {
                // 處理 HTML 匯入
                const doc = new DOMParser().parseFromString(content, 'text/html');
                const getVal = (row, f) => {
                    const el = row.querySelector('.v-' + f);
                    if (!el) return '';
                    return el.tagName === 'INPUT' ? el.value : el.textContent;
                };
                const data = {
                    metadata: { 
                        org: doc.querySelector('#mgr-org')?.value || '', 
                        year: doc.querySelector('#mgr-year')?.value || '',
                        user: doc.querySelector('#mgr-user')?.value || ''
                    },
                    sections: sectionConfigs.map(conf => ({
                        id: conf.id,
                        items: Array.from(doc.querySelectorAll(`#tbody-${conf.id} tr`)).map(tr => {
                            let item = {};
                            conf.fields.forEach(f => item[f] = getVal(tr, f));
                            return item;
                        }).filter(i => i.name)
                    }))
                };
                mgr_populate(data);
                alert('HTML 匯入成功！');
            }
            document.getElementById('mgr-import-file').value = ''; 
        } catch (err) {
            console.error(err);
            alert('匯入失敗：檔案格式錯誤。');
        }
    };
    reader.readAsText(file);
}

function mgr_populate(data) {
    // 回填 Metadata
    if (data.metadata) {
        if(document.getElementById('mgr-org')) document.getElementById('mgr-org').value = data.metadata.org || '';
        if(document.getElementById('mgr-year')) document.getElementById('mgr-year').value = data.metadata.year || '';
        if(document.getElementById('mgr-user')) document.getElementById('mgr-user').value = data.metadata.user || '';
    }

    sectionConfigs.forEach(conf => {
        const tbody = document.getElementById(`tbody-${conf.id}`);
        if (!tbody) return;
        tbody.innerHTML = ''; 

        // 關鍵修正：判斷資料來源格式
        let items = [];
        if (data.sections && Array.isArray(data.sections)) {
            // 新版結構： { sections: [ { id: 'op', items: [...] } ] }
            const section = data.sections.find(s => s.id === conf.id);
            if (section) items = section.items;
        } else if (data[conf.id] && Array.isArray(data[conf.id])) {
            // 舊版結構： { op: [...], wk: [...] }
            items = data[conf.id];
        }

        if (items && items.length > 0) {
            items.forEach(item => mgr_addRow(conf.id, item));
        } else {
            mgr_addRow(conf.id);
        }
        update(conf.id);
    });
}

// ========== 3. 匯出功能 (防呆 + 清洗 + 格線) ==========
function mgr_exportHTML() {
    try {
        // 1. 取得當前畫面上的所有輸入框 (來源資料)
        const sourceInputs = document.querySelectorAll('#sections-container input');
        
        // 檢查是否有資料
        const hasData = Array.from(sourceInputs).some(input => input.value.trim() !== "" && !input.readOnly);
        if (!hasData) {
            if(!confirm('⚠️ 警告：目前報表似乎是空白的。\n是否仍要強制匯出？')) return;
        }

        // 2. 複製整份文件
        let cloneDoc = document.documentElement.cloneNode(true);
        const tabManager = cloneDoc.querySelector('#tab-manager');
        if (!tabManager) { alert('匯出錯誤：找不到報表容器'); return; }

        // 3. 【關鍵修正】強制數值同步：一對一將來源資料填入複製後的 HTML
        // 我們直接對應 Index，這比 setAttribute 更可靠
        const targetInputs = cloneDoc.querySelectorAll('#sections-container input');
        
        if (sourceInputs.length !== targetInputs.length) {
            console.warn('DOM 結構不一致，改用屬性同步模式');
            // 備用方案
            document.querySelectorAll('input').forEach(i => i.setAttribute('value', i.value));
            cloneDoc = document.documentElement.cloneNode(true);
        } else {
            // 主力方案：直接把畫面上的值 (source) 塞進備份檔 (target) 的文字內容中
            sourceInputs.forEach((src, idx) => {
                const tgt = targetInputs[idx];
                // 將值暫存於 data-value 屬性，確保轉 span 時抓得到
                tgt.setAttribute('value', src.value); 
                tgt.value = src.value;
            });
        }

        // 4. 清洗介面 (移除不需要的按鈕與文字)
        cloneDoc.querySelector('nav')?.remove();
        cloneDoc.querySelector('#tab-aggregator')?.remove();
        cloneDoc.querySelectorAll('.excel-guide, .flex.gap-2, #btn-clear, script, .add-row-btn, .delete-btn, #autosave-indicator, #undo-btn').forEach(el => el.remove());

        const unwantedTexts = ['預算填報工作站', '支援 Excel 貼上', '自動儲存'];
        cloneDoc.querySelectorAll('h1, h2, h3, p, div, span').forEach(el => {
            unwantedTexts.forEach(text => { if (el.textContent.includes(text)) el.remove(); });
        });

        // 5. 將 Input 轉為 Span (靜態文字)
        // 這裡直接讀取我們剛剛強制寫入的 attribute
        cloneDoc.querySelectorAll('input').forEach(input => {
            const val = input.getAttribute('value') || input.value || '';
            const span = document.createElement('span');
            
            // 處理千分位：如果是數字，加上千分位逗號讓報表更好看
            let displayVal = val;
            if (!isNaN(parseFloat(val)) && val.trim() !== '') {
                displayVal = parseFloat(val).toLocaleString();
            }

            span.textContent = displayVal;
            // 繼承樣式 (靠右、紅字) 並給予最小寬度避免格線塌陷
            span.className = input.className + " inline-block min-w-[1em] min-h-[1.2em] w-full";
            
            // 移除 input 邊框相關的 class，確保 span 乾淨
            span.classList.remove('border', 'border-b-2', 'outline-none');
            
            input.parentNode.replaceChild(span, input);
        });

        // 6. 建立報表標頭
        const now = new Date();
        const dateStr = `${now.getFullYear()-1911}年${now.getMonth()+1}月${now.getDate()}日 ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}`;
        const printHeader = document.createElement('div');
        printHeader.className = "no-print";
        printHeader.style.cssText = "max-width:1200px; margin:20px auto; display:flex; justify-content:space-between; align-items:flex-end; border-bottom:3px solid #000; padding-bottom:15px;";
        printHeader.innerHTML = `
            <div><button id="p-btn" style="background:#2563eb; color:white; padding:12px 28px; border-radius:6px; font-weight:bold; cursor:pointer; border:none; font-size:16px;">🖨️ 列印 / 儲存 PDF</button></div>
            <div style="text-align:right; color:#000; font-weight:bold; font-size:16px;">產製日期：${dateStr}</div>
        `;
        
        tabManager.prepend(printHeader);
        tabManager.classList.remove('hidden');
        tabManager.style.background = "white";
        tabManager.style.padding = "20px";

        // 7. 注入 CSS (強制黑格線)
        const styleTag = document.createElement('style');
        styleTag.textContent = `
            body { background: white !important; font-family: "Noto Sans TC", sans-serif; }
            .budget-table { 
                border-collapse: collapse !important; width: 100% !important; 
                border: 2px solid #000 !important; margin-bottom: 40px; background: white; 
            }
            .budget-table th, .budget-table td { 
                border: 1px solid #000 !important; padding: 8px 4px !important; 
                text-align: center; color: #000 !important; font-size: 14px; 
            }
            .budget-table th { background-color: #f1f5f9 !important; font-weight: bold; }
            .section-card { border: none !important; margin-bottom: 60px; box-shadow: none !important; background: white; }
            h3 { margin-bottom: 15px; font-size: 1.5rem; color: #000 !important; font-weight: bold; }
            .negative-value { color: #dc2626 !important; font-weight: bold; }
            @media print { .no-print { display: none !important; } .section-card { break-inside: avoid; } body { padding: 0; } }
        `;
        cloneDoc.querySelector('head').appendChild(styleTag);

        // 8. 注入列印腳本
        const printScript = document.createElement('script');
        printScript.textContent = `document.addEventListener('DOMContentLoaded',function(){document.getElementById('p-btn').onclick=function(){window.print()}});`;
        cloneDoc.querySelector('body').appendChild(printScript);

        // 9. 下載
        const htmlContent = "<!DOCTYPE html>\n" + cloneDoc.outerHTML;
        const org = document.getElementById('mgr-org').value || '預算報表';
        saveAs(new Blob([htmlContent], { type: "text/html" }), `正式報表_${org}.html`);

    } catch (err) { 
        console.error(err); 
        alert('匯出失敗：' + err.message); 
    }
}

// ========== 4. 匯整端邏輯 ==========
let agg_data = [];
function agg_processFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            let data;
            const content = e.target.result;
            // 匯整端也要支援 JSON 匯入
            if (file.name.endsWith('.json') || content.trim().startsWith('{')) {
                data = JSON.parse(content);
                // 正規化 JSON 結構給 Aggregator 使用
                if (!data.sections) {
                    data.sections = sectionConfigs.map(c => ({ id: c.id, items: data[c.id] || [] }));
                }
            } else {
                const doc = new DOMParser().parseFromString(content, 'text/html');
                const getVal = (row, f) => {
                    const el = row.querySelector('.v-'+f);
                    return el ? (el.tagName === 'INPUT' ? el.value : el.textContent) : '';
                };
                data = {
                    metadata: { org: doc.querySelector('#mgr-org')?.value || doc.querySelector('#mgr-org')?.textContent || file.name.replace('.html',''), year: '', user: '' },
                    sections: sectionConfigs.map(conf => ({
                        id: conf.id,
                        items: Array.from(doc.querySelectorAll(`#tbody-${conf.id} tr`)).map(tr => {
                            let item = {};
                            conf.fields.forEach(f => item[f] = getVal(tr, f));
                            return item;
                        }).filter(i => i.name)
                    }))
                };
            }
            agg_data.push(data);
            agg_render();
        } catch (err) { alert('檔案解析失敗：' + err.message); }
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
                let bal = num(item.net) || num(item.surplus);
                stats.totalRev += rev;
                if (bal >= 0) stats.profit++; else stats.loss++;
            });
        });
    });
    document.getElementById('agg-kpi').innerHTML = `
        <div class="bg-slate-800 p-4 rounded-lg"><div>機關數</div><div class="text-2xl font-bold text-blue-400">${stats.govs}</div></div>
        <div class="bg-slate-800 p-4 rounded-lg"><div>基金數</div><div class="text-2xl font-bold text-green-400">${stats.funds}</div></div>
        <div class="bg-slate-800 p-4 rounded-lg"><div>規模(億)</div><div class="text-2xl font-bold text-emerald-400">${(stats.totalRev / 100000).toFixed(2)}</div></div>
        <div class="bg-slate-800 p-4 rounded-lg"><div>盈虧</div><div class="text-sm">盈: ${stats.profit}/虧: ${stats.loss}</div></div>
    `;
    document.getElementById('agg-list-body').innerHTML = agg_data.map((d, i) => `
        <tr class="border-b border-slate-700"><td class="p-3 text-slate-500">${i+1}</td><td class="p-3 font-bold text-blue-300">${d.metadata.org || '未命名機關'}</td><td class="p-3 text-right"><button class="text-red-400 text-sm" onclick="window.agg_remove(${i})">移除</button></td></tr>
    `).join('');
}
window.agg_remove = (idx) => { agg_data.splice(idx,1); agg_render(); };

// ========== 5. 基礎功能與事件 ==========
function render() {
    const container = document.getElementById('sections-container');
    container.innerHTML = '';
    sectionConfigs.forEach(conf => {
        const div = document.createElement('div');
        div.className = 'section-card bg-white rounded-xl shadow-sm border-t-4 p-6 mb-6';
        div.style.borderTopColor = conf.color;
        div.innerHTML = `
            <div class="flex justify-between items-center mb-4"><h3 class="font-bold text-lg" style="color:${conf.color}">${conf.title}</h3><button class="add-row-btn text-blue-600 font-bold hover:underline" data-section="${conf.id}">+ 新增</button></div>
            <div class="overflow-x-auto"><table class="budget-table"><thead><tr><th class="w-8"></th>${conf.fields.map(f => `<th>${getFieldLabel(conf.id, f)}</th>`).join('')}</tr></thead><tbody id="tbody-${conf.id}"></tbody>
            <tfoot id="tfoot-${conf.id}" class="bg-slate-50 font-bold"><tr><td>∑</td><td class="text-center">合計</td>${conf.fields.slice(1).map(f => `<td><input type="text" class="t-${f}" readonly value="0"></td>`).join('')}</tr></tfoot></table></div>`;
        container.appendChild(div);
    });
}

function mgr_addRow(type, data = {}) {
    const conf = sectionConfigs.find(c => c.id === type);
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="text-center"><button class="delete-btn text-red-400" data-type="${type}">✕</button></td>` +
        conf.fields.map(f => `<td class="p-2"><input type="${f==='name'?'text':'number'}" class="v-${f} ${f==='name'?'text-left':''}" value="${data[f]||''}" ${['gross','opprofit','nonprofit','pretax','net','surplus','nonsurplus','end'].includes(f)?'readonly':''}></td>`).join('');
    document.getElementById(`tbody-${type}`).appendChild(tr);
}

function update(type) {
    const rows = document.querySelectorAll(`#tbody-${type} tr`), conf = sectionConfigs.find(c => c.id === type);
    let totals = {}; conf.fields.slice(1).forEach(f => totals[f] = 0);
    rows.forEach(row => {
        const v = (c) => parseFloat(row.querySelector('.v-'+c)?.value) || 0, r = (c) => row.querySelector('.v-'+c);
        if (type === 'op') {
            let g = v('rev') - v('cost'), op = g - v('exp'), np = v('nonrev') - v('nonexp');
            if(r('gross')) r('gross').value = g; if(r('opprofit')) r('opprofit').value = op; if(r('nonprofit')) r('nonprofit').value = np; if(r('pretax')) r('pretax').value = op+np; if(r('net')) r('net').value = op+np-v('tax');
        } else if (type === 'wk') {
            let s = v('rev') - v('cost'), ns = v('nonrev') - v('nonexp');
            if(r('surplus')) r('surplus').value = s; if(r('nonsurplus')) r('nonsurplus').value = ns; if(r('net')) r('net').value = s + ns;
        } else {
            let s = v('source') - v('use');
            if(r('surplus')) r('surplus').value = s; if(r('end')) r('end').value = v('begin')+s-v('remit');
        }
        conf.fields.slice(1).forEach(f => { let val = v(f); row.querySelector('.v-'+f)?.classList.toggle('negative-value', val < 0); totals[f] += val; });
    });
    Object.keys(totals).forEach(f => { const tEl = document.querySelector(`#tfoot-${type} .t-${f}`); if (tEl) { tEl.value = totals[f].toLocaleString(); tEl.classList.toggle('negative-value', totals[f] < 0); } });
}

function bindEvents() {
    document.getElementById('btn-manager').onclick = () => { document.getElementById('tab-manager').classList.remove('hidden'); document.getElementById('tab-aggregator').classList.add('hidden'); };
    document.getElementById('btn-aggregator').onclick = () => { document.getElementById('tab-manager').classList.add('hidden'); document.getElementById('tab-aggregator').classList.remove('hidden'); };
    
    document.getElementById('btn-export-html').onclick = mgr_exportHTML;
    document.getElementById('btn-import').onclick = () => document.getElementById('mgr-import-file').click();
    document.getElementById('mgr-import-file').onchange = (e) => mgr_handleImport(e.target.files);

    document.getElementById('btn-agg-clear').onclick = () => { agg_data = []; agg_render(); };
    document.getElementById('sections-container').onclick = e => {
        if (e.target.classList.contains('add-row-btn')) mgr_addRow(e.target.dataset.section);
        if (e.target.classList.contains('delete-btn')) { e.target.closest('tr').remove(); update(e.target.dataset.type); }
    };
    document.getElementById('sections-container').oninput = e => { const tbody = e.target.closest('tbody'); if (tbody) update(tbody.id.replace('tbody-', '')); };
    const dz = document.getElementById('agg-dropzone');
    dz.onclick = () => { const inp = document.createElement('input'); inp.type = 'file'; inp.multiple = true; inp.onchange = e => Array.from(e.target.files).forEach(f => agg_processFile(f)); inp.click(); };
}

document.addEventListener('DOMContentLoaded', () => { render(); bindEvents(); sectionConfigs.forEach(c => mgr_addRow(c.id)); });
