'use strict';

// ========== Configuration ==========
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

const sectionConfigs = [
    { id: 'op', title: '一、營業基金', color: '#2563eb', fields: ['name','rev','cost','gross','exp','opprofit','nonrev','nonexp','nonprofit','pretax','tax','net'] },
    { id: 'wk', title: '二、作業基金', color: '#16a34a', fields: ['name','rev','cost','surplus','nonrev','nonexp','nonsurplus','net'] },
    { id: 'db', title: '三、債務基金', color: '#ea580c', fields: ['name','source','use','surplus','begin','remit','end'] },
    { id: 'sp', title: '四、特別收入基金', color: '#9333ea', fields: ['name','source','use','surplus','begin','remit','end'] },
    { id: 'cp', title: '五、資本計畫基金', color: '#0891b2', fields: ['name','source','use','surplus','begin','remit','end'] }
];

const fieldNames = {
    name: '基金名稱',
    // 原始欄位映射 (用於程式碼邏輯)
    rev: '收入', cost: '成本/費用', gross: '營業毛利(毛損)', exp: '營業費用',
    opprofit: '營業利益(損失)', nonrev: '外收入', nonexp: '外費用', nonprofit: '外利益(損失)',
    pretax: '稅前淨利(淨損)', tax: '所得稅費用(利益)', net: '本期淨利/賸餘',
    surplus: '本期賸餘(短絀)', nonsurplus: '外賸餘(短絀)',
    source: '基金來源', use: '基金用途', begin: '期初基金餘額', remit: '解繳公庫', end: '期末基金餘額'
};

// ========== Helper: 取得精確標題 ==========
function getFieldLabel(sectionId, fieldId) {
    // 營業基金
    if (sectionId === 'op') {
        const labels = {
            rev: '營業收入', cost: '營業成本', gross: '營業毛利(毛損)', exp: '營業費用',
            opprofit: '營業利益(損失)', nonrev: '營業外收入', nonexp: '營業外費用',
            nonprofit: '營業外利益(損失)', pretax: '稅前淨利(淨損)', tax: '所得稅費用(利益)', net: '本期淨利(淨損)'
        };
        return labels[fieldId] || fieldNames[fieldId];
    }
    // 作業基金
    if (sectionId === 'wk') {
        const labels = {
            rev: '業務收入', cost: '業務成本與費用', surplus: '業務賸餘(短絀)',
            nonrev: '業務外收入', nonexp: '業務外費用', nonsurplus: '業務外賸餘(短絀)', net: '本期賸餘(短絀)'
        };
        return labels[fieldId] || fieldNames[fieldId];
    }
    // 債務、特別收入、資本計畫基金
    if (['db', 'sp', 'cp'].includes(sectionId)) {
        if (fieldId === 'surplus') return '本期賸餘(短絀)';
    }
    return fieldNames[fieldId];
}

// ========== Undo System ==========
const undoStack = [];
const MAX_UNDO = 50;

function saveState() {
    const state = collectData();
    undoStack.push(JSON.stringify(state));
    if (undoStack.length > MAX_UNDO) undoStack.shift();
    const undoBtn = document.getElementById('undo-btn');
    if (undoBtn) undoBtn.classList.toggle('show', undoStack.length > 1);
}

function undo() {
    if (undoStack.length <= 1) return;
    undoStack.pop();
    const prevState = undoStack[undoStack.length - 1];
    if (prevState) {
        mgr_populate(JSON.parse(prevState), false);
        showAutosave('↩ 已復原');
    }
    const undoBtn = document.getElementById('undo-btn');
    if (undoBtn) undoBtn.classList.toggle('show', undoStack.length > 1);
}

// ========== Debounce ==========
function debounce(fn, delay) {
    let timer;
    return function(...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

// ========== localStorage Auto-save ==========
const STORAGE_KEY = 'budget_data_v4';

function collectData() {
    return {
        metadata: {
            org: document.getElementById('mgr-org').value,
            year: document.getElementById('mgr-year').value,
            user: document.getElementById('mgr-user').value,
            savedAt: new Date().toISOString()
        },
        sections: sectionConfigs.map(conf => ({
            id: conf.id,
            items: Array.from(document.querySelectorAll(`#tbody-${conf.id} tr`)).map(tr => {
                let item = {};
                conf.fields.forEach(f => item[f] = tr.querySelector('.v-'+f)?.value || '');
                return item;
            })
        }))
    };
}

function saveToStorage() {
    try {
        const data = collectData();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
        showAutosave();
    } catch (e) {
        console.error('儲存失敗:', e);
    }
}

function loadFromStorage() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            const data = JSON.parse(saved);
            mgr_populate(data, false);
            return true;
        }
    } catch (e) {
        console.error('載入失敗:', e);
    }
    return false;
}

function showAutosave(msg = '✓ 已自動儲存') {
    const indicator = document.getElementById('autosave-indicator');
    if (!indicator) return;
    indicator.textContent = msg;
    indicator.classList.add('show');
    setTimeout(() => indicator.classList.remove('show'), 1500);
}

const debouncedSave = debounce(() => {
    saveState();
    saveToStorage();
}, 500);

// ========== Core Functions ==========
function switchTab(id) {
    document.getElementById('tab-manager').classList.toggle('hidden', id !== 'manager');
    document.getElementById('tab-aggregator').classList.toggle('hidden', id !== 'aggregator');
    document.getElementById('btn-manager').className = `px-6 py-2 rounded-full font-bold transition ${id === 'manager' ? 'bg-white shadow-sm' : 'text-slate-500'}`;
    document.getElementById('btn-aggregator').className = `px-6 py-2 rounded-full font-bold transition ${id === 'aggregator' ? 'bg-white shadow-sm' : 'text-slate-500'}`;
}

function render() {
    const container = document.getElementById('sections-container');
    container.innerHTML = '';
    sectionConfigs.forEach(conf => {
        const div = document.createElement('div');
        div.className = 'section-card';
        div.style.borderTopColor = conf.color;
        
        const headerHtml = `
            <div class="flex justify-between items-center mb-4">
                <h3 class="font-bold text-lg" style="color:${conf.color}">${conf.title}</h3>
                <button class="add-row-btn text-blue-600 text-xs font-bold hover:underline" data-section="${conf.id}">+ 手動新增</button>
            </div>`;
            
        const tableHtml = `
            <div class="overflow-x-auto">
                <table class="w-full budget-table">
                    <thead>
                        <tr>
                            <th class="w-8"></th>
                            ${conf.fields.map(f => `<th>${getFieldLabel(conf.id, f)}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody id="tbody-${conf.id}"></tbody>
                    <tfoot id="tfoot-${conf.id}" class="total-row">
                        <tr>
                            <td>∑</td>
                            <td class="text-center">合計</td>
                            ${conf.fields.slice(1).map(f => `<td><input type="text" class="t-${f}" readonly value="0"></td>`).join('')}
                        </tr>
                    </tfoot>
                </table>
            </div>`;
            
        div.innerHTML = headerHtml + tableHtml;
        container.appendChild(div);
    });
}

function mgr_addRow(type, data = {}) {
    const conf = sectionConfigs.find(c => c.id === type);
    if (!conf) return;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="text-center"><button class="delete-btn text-red-400 hover:text-red-600" data-type="${escapeHtml(type)}">✕</button></td>` +
        conf.fields.map(f => {
            const isRO = ['gross','opprofit','nonprofit','pretax','net','surplus','nonsurplus','end'].includes(f);
            const val = data[f] !== undefined ? data[f] : '';
            return `<td data-tip="${isRO ? '系統自動計算' : '支援從 Excel 貼上數據'}">
                <input type="${f === 'name' ? 'text' : 'number'}" class="v-${f} ${f === 'name' ? 'text-left' : ''}" value="${escapeHtml(val)}" ${isRO ? 'readonly' : ''}>
            </td>`;
        }).join('');
    document.getElementById(`tbody-${type}`).appendChild(tr);
}

function deleteRow(btn, type) {
    saveState();
    btn.closest('tr').remove();
    update(type);
    debouncedSave();
}

function handleInput(type) {
    update(type);
    debouncedSave();
}

function update(type) {
    const rows = document.querySelectorAll(`#tbody-${type} tr`);
    const conf = sectionConfigs.find(c => c.id === type);
    if (!conf) return;
    let totals = {};
    conf.fields.slice(1).forEach(f => totals[f] = 0);
    
    rows.forEach(row => {
        const v = (c) => parseFloat(row.querySelector('.v-'+c)?.value) || 0;
        const r = (c) => row.querySelector('.v-'+c);
        
        if (type === 'op') {
            let g = v('rev') - v('cost'), op = g - v('exp'), np = v('nonrev') - v('nonexp');
            if (r('gross')) r('gross').value = g;
            if (r('opprofit')) r('opprofit').value = op;
            if (r('nonprofit')) r('nonprofit').value = np;
            if (r('pretax')) r('pretax').value = op + np;
            if (r('net')) r('net').value = op + np - v('tax');
        } else if (type === 'wk') {
            let s = v('rev') - v('cost'), ns = v('nonrev') - v('nonexp');
            if (r('surplus')) r('surplus').value = s;
            if (r('nonsurplus')) r('nonsurplus').value = ns;
            if (r('net')) r('net').value = s + ns;
        } else {
            let s = v('source') - v('use');
            if (r('surplus')) r('surplus').value = s;
            if (r('end')) r('end').value = v('begin') + s - v('remit');
        }
        
        conf.fields.slice(1).forEach(f => {
            let val = v(f);
            const el = row.querySelector('.v-'+f);
            if (el) el.classList.toggle('negative-value', val < 0);
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

// ========== Import/Export ==========
function mgr_handleImport(files) {
    const file = files[0];
    if (!file) return;
    const r = new FileReader();
    r.onload = (e) => {
        try {
            const content = e.target.result;
            if (file.name.endsWith('.json')) {
                const data = JSON.parse(content);
                mgr_populate(data);
            } else {
                const doc = new DOMParser().parseFromString(content, 'text/html');
                const data = {
                    metadata: { org: doc.getElementById('mgr-org')?.value, year: doc.getElementById('mgr-year')?.value },
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
                mgr_populate(data);
            }
        } catch (err) {
            alert('匯入失敗：檔案格式錯誤\n' + err.message);
        }
    };
    r.readAsText(file);
}

function mgr_populate(data, shouldSaveState = true) {
    if (shouldSaveState) saveState();
    document.getElementById('mgr-org').value = data.metadata?.org || '';
    document.getElementById('mgr-year').value = data.metadata?.year || '115';
    document.getElementById('mgr-user').value = data.metadata?.user || '';
    sectionConfigs.forEach(conf => {
        const tbody = document.getElementById(`tbody-${conf.id}`);
        if (!tbody) return;
        tbody.innerHTML = '';
        const section = data.sections?.find(s => s.id === conf.id);
        if (section && section.items?.length > 0) {
            section.items.forEach(i => mgr_addRow(conf.id, i));
        } else {
            mgr_addRow(conf.id);
        }
        update(conf.id);
    });
    if (shouldSaveState) debouncedSave();
}

function mgr_exportXLSX() {
    try {
        const sheetData = [];
        sectionConfigs.forEach(conf => {
            sheetData.push([conf.title]);
            // 匯出時使用精確標題
            sheetData.push(conf.fields.map(f => getFieldLabel(conf.id, f)));
            Array.from(document.querySelectorAll(`#tbody-${conf.id} tr`)).forEach(tr => {
                sheetData.push(conf.fields.map((f, idx) => {
                    const val = tr.querySelector('.v-'+f)?.value || '';
                    return idx === 0 ? val : { v: parseFloat(val) || 0, t: 'n', z: '#,##0' };
                }));
            });
            const footRow = ["合計"];
            conf.fields.slice(1).forEach(f => footRow.push({
                v: parseFloat(document.querySelector(`#tfoot-${conf.id} .t-${f}`)?.value.replace(/,/g, '')) || 0,
                t: 'n', z: '#,##0'
            }));
            sheetData.push(footRow, []);
        });
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "預算總表");
        XLSX.writeFile(wb, "預算報表.xlsx");
    } catch (err) {
        alert('匯出失敗：' + err.message);
    }
}

function mgr_exportJSON() {
    try {
        const data = collectData();
        const orgName = data.metadata.org || '未命名';
        saveAs(new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }), `${orgName}_預算.json`);
    } catch (err) {
        alert('匯出失敗：' + err.message);
    }
}

function mgr_exportHTML() {
    try {
        document.querySelectorAll('input').forEach(i => i.setAttribute('value', i.value));
        saveAs(new Blob(["<!DOCTYPE html>\n" + document.documentElement.outerHTML], { type: "text/html" }), "預算備份.html");
    } catch (err) {
        alert('匯出失敗：' + err.message);
    }
}

function mgr_clearAll() {
    if (confirm('確定清空所有資料？\n（此操作無法復原）')) {
        localStorage.removeItem(STORAGE_KEY);
        location.reload();
    }
}

// ========== Aggregator ==========
let agg_data = [];

function agg_setupDrag() {
    const dropzone = document.getElementById('agg-dropzone');
    if (!dropzone) return;

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, e => {
            e.preventDefault();
            e.stopPropagation();
        });
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => dropzone.classList.add('dragover'));
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => dropzone.classList.remove('dragover'));
    });

    dropzone.addEventListener('drop', e => {
        const files = e.dataTransfer.files;
        Array.from(files).forEach(file => agg_processFile(file));
    });

    dropzone.addEventListener('click', () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json,.html';
        input.multiple = true;
        input.addEventListener('change', e => {
            Array.from(e.target.files).forEach(file => agg_processFile(file));
        });
        input.click();
    });
}

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
                    metadata: { org: doc.getElementById('mgr-org')?.value || file.name },
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
        } catch (err) {
            alert('處理檔案失敗：' + err.message);
        }
    };
    reader.readAsText(file);
}

function agg_render() {
    document.getElementById('agg-content').classList.remove('hidden');
    let totalOrgs = agg_data.length;
    let totalFunds = 0;
    let totalNet = 0;

    agg_data.forEach(d => {
        d.sections?.forEach(s => {
            totalFunds += s.items?.length || 0;
            s.items?.forEach(item => {
                totalNet += parseFloat(item.net) || parseFloat(item.end) || 0;
            });
        });
    });

    document.getElementById('agg-kpi').innerHTML = `
        <div class="kpi-card"><div class="text-slate-400 text-sm">已匯入機關</div><div class="kpi-value text-blue-400">${totalOrgs}</div></div>
        <div class="kpi-card"><div class="text-slate-400 text-sm">基金總數</div><div class="kpi-value text-green-400">${totalFunds}</div></div>
        <div class="kpi-card"><div class="text-slate-400 text-sm">淨額合計</div><div class="kpi-value ${totalNet < 0 ? 'text-red-400' : 'text-emerald-400'}">${totalNet.toLocaleString()}</div></div>
        <div class="kpi-card"><div class="text-slate-400 text-sm">匯整時間</div><div class="kpi-value text-purple-400 text-lg">${new Date().toLocaleTimeString()}</div></div>
    `;

    document.getElementById('agg-list-body').innerHTML = agg_data.map((d, i) => `
        <tr class="border-b border-slate-700">
            <td class="py-3 text-slate-400">${i + 1}</td>
            <td class="py-3 font-bold">${escapeHtml(d.metadata?.org || '未命名')}</td>
            <td class="py-3">${d.sections?.reduce((sum, s) => sum + (s.items?.length || 0), 0) || 0} 筆基金</td>
            <td class="py-3"><button class="agg-remove-btn text-red-400 hover:text-red-300" data-index="${i}">移除</button></td>
        </tr>
    `).join('');
}

function agg_clearAll() {
    agg_data = [];
    document.getElementById('agg-content').classList.add('hidden');
}

// ========== Event Binding ==========
function bindStaticEvents() {
    document.getElementById('btn-manager').addEventListener('click', () => switchTab('manager'));
    document.getElementById('btn-aggregator').addEventListener('click', () => switchTab('aggregator'));
    document.getElementById('btn-import').addEventListener('click', () => document.getElementById('mgr-import-file').click());
    document.getElementById('mgr-import-file').addEventListener('change', function() { mgr_handleImport(this.files); });
    document.getElementById('btn-export-json').addEventListener('click', mgr_exportJSON);
    document.getElementById('btn-export-html').addEventListener('click', mgr_exportHTML);
    document.getElementById('btn-export-xlsx').addEventListener('click', mgr_exportXLSX);
    document.getElementById('btn-clear').addEventListener('click', mgr_clearAll);
    document.getElementById('btn-agg-clear').addEventListener('click', agg_clearAll);
    document.getElementById('undo-btn')?.addEventListener('click', undo);
    ['mgr-org', 'mgr-year', 'mgr-user'].forEach(id => {
        document.getElementById(id)?.addEventListener('input', debouncedSave);
    });
}

function bindDynamicEvents() {
    const container = document.getElementById('sections-container');
    container.addEventListener('click', e => {
        const addBtn = e.target.closest('.add-row-btn');
        if (addBtn) { mgr_addRow(addBtn.dataset.section); return; }
        const delBtn = e.target.closest('.delete-btn');
        if (delBtn) { deleteRow(delBtn, delBtn.dataset.type); return; }
    });
    container.addEventListener('input', e => {
        if (e.target.tagName === 'INPUT') {
            const tbody = e.target.closest('tbody');
            if (tbody?.id?.startsWith('tbody-')) handleInput(tbody.id.replace('tbody-', ''));
        }
    });
    document.getElementById('agg-list-body').addEventListener('click', e => {
        const removeBtn = e.target.closest('.agg-remove-btn');
        if (removeBtn) {
            agg_data.splice(parseInt(removeBtn.dataset.index), 1);
            agg_render();
        }
    });
}

function bindKeyboardEvents() {
    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
            if (document.activeElement.tagName !== 'INPUT') { e.preventDefault(); undo(); }
        }
    });

    document.addEventListener('keydown', e => {
        const el = document.activeElement;
        if (!el || el.tagName !== 'INPUT' || el.readOnly) return;
        if (e.key === 'Enter') {
            e.preventDefault();
            const tr = el.closest('tr'), tbody = tr.parentNode;
            const colIdx = Array.from(tr.children).indexOf(el.closest('td'));
            let nextRow = tr.nextElementSibling;
            if (!nextRow) { 
                mgr_addRow(tbody.id.replace('tbody-', ''));
                nextRow = tbody.lastElementChild;
            }
            nextRow.children[colIdx]?.querySelector('input:not([readonly])')?.focus();
        }
    });

    document.addEventListener('paste', e => {
        const el = document.activeElement;
        if (!el || el.readOnly || el.tagName !== 'INPUT') return;
        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (!text.includes('\t') && !text.includes('\n')) return;
        e.preventDefault();
        saveState();
        const rowsData = text.split(/\r\n|\n|\r/).filter(r => r.trim());
        const startTr = el.closest('tr'), tbody = startTr.parentNode;
        const colIdx = Array.from(startTr.children).indexOf(el.closest('td'));
        const rowIdx = Array.from(tbody.children).indexOf(startTr);
        
        rowsData.forEach((rowStr, i) => {
            let targetTr = tbody.children[rowIdx + i];
            if (!targetTr) { 
                mgr_addRow(tbody.id.replace('tbody-', ''));
                targetTr = tbody.lastElementChild;
            }
            rowStr.split('\t').forEach((val, j) => {
                const input = targetTr.children[colIdx + j]?.querySelector('input:not([readonly])');
                if (input) {
                    input.value = val.trim().replace(/,/g, '');
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                }
            });
        });
        debouncedSave();
    });
}

// ========== Initialization ==========
document.addEventListener('DOMContentLoaded', () => {
    render();
    agg_setupDrag();
    bindStaticEvents();
    bindDynamicEvents();
    bindKeyboardEvents();
    if (!loadFromStorage()) sectionConfigs.forEach(c => mgr_addRow(c.id));
    saveState();
});
