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
    { id: 'op', title: '\u4e00\u3001\u71df\u696d\u57fa\u91d1', color: '#2563eb', fields: ['name','rev','cost','gross','exp','opprofit','nonrev','nonexp','nonprofit','pretax','tax','net'] },
    { id: 'wk', title: '\u4e8c\u3001\u4f5c\u696d\u57fa\u91d1', color: '#16a34a', fields: ['name','rev','cost','surplus','nonrev','nonexp','nonsurplus','net'] },
    { id: 'db', title: '\u4e09\u3001\u50b5\u52d9\u57fa\u91d1', color: '#ea580c', fields: ['name','source','use','surplus','begin','remit','end'] },
    { id: 'sp', title: '\u56db\u3001\u7279\u5225\u6536\u5165\u57fa\u91d1', color: '#9333ea', fields: ['name','source','use','surplus','begin','remit','end'] },
    { id: 'cp', title: '\u4e94\u3001\u8cc7\u672c\u8a08\u756b\u57fa\u91d1', color: '#0891b2', fields: ['name','source','use','surplus','begin','remit','end'] }
];

const fieldNames = {
    name:'\u57fa\u91d1\u540d\u7a31', rev:'\u4e00\u3001\u6536\u5165', cost:'\u4e00\u3001\u6210\u672c\u8cbb\u7528', gross:'\u4e00\u3001\u6bdb\u5229', exp:'\u4e00\u3001\u8cbb\u7528',
    opprofit:'\u4e00\u3001\u5229\u76ca', nonrev:'\u4e00\u3001\u5916\u6536\u5165', nonexp:'\u4e00\u3001\u5916\u8cbb\u7528', nonprofit:'\u4e00\u3001\u5916\u5229\u76ca',
    pretax:'\u7a05\u524d', tax:'\u6240\u5f97\u7a05', net:'\u672c\u671f\u76c8\u8667', surplus:'\u672c\u671f\u8cf8\u9918',
    nonsurplus:'\u5916\u8cf8\u9918', source:'\u4f86\u6e90', use:'\u7528\u9014', begin:'\u671f\u521d', remit:'\u89e3\u7e73', end:'\u671f\u672b'
};

// ========== Undo System ==========
const undoStack = [];
const MAX_UNDO = 50;

function saveState() {
    const state = collectData();
    undoStack.push(JSON.stringify(state));
    if (undoStack.length > MAX_UNDO) undoStack.shift();
    document.getElementById('undo-btn').classList.toggle('show', undoStack.length > 1);
}

function undo() {
    if (undoStack.length <= 1) return;
    undoStack.pop();
    const prevState = undoStack[undoStack.length - 1];
    if (prevState) {
        mgr_populate(JSON.parse(prevState), false);
        showAutosave('\u21a9 \u5df2\u5fa9\u539f');
    }
    document.getElementById('undo-btn').classList.toggle('show', undoStack.length > 1);
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
const STORAGE_KEY = 'budget_data_v3';

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
        console.error('\u5132\u5b58\u5931\u6557:', e);
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
        console.error('\u8f09\u5165\u5931\u6557:', e);
    }
    return false;
}

function showAutosave(msg = '\u2713 \u5df2\u81ea\u52d5\u5132\u5b58') {
    const indicator = document.getElementById('autosave-indicator');
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
        // No inline onclick - use data-section attribute + event delegation
        div.innerHTML = `<div class="flex justify-between items-center mb-4"><h3 class="font-bold text-lg" style="color:${conf.color}">${conf.title}</h3><button class="add-row-btn text-blue-600 text-xs font-bold hover:underline" data-section="${conf.id}">+ \u624b\u52d5\u65b0\u589e</button></div>
            <div class="overflow-x-auto"><table class="w-full budget-table"><thead><tr><th class="w-8"></th>${conf.fields.map(f => `<th>${fieldNames[f]}</th>`).join('')}</tr></thead><tbody id="tbody-${conf.id}"></tbody><tfoot id="tfoot-${conf.id}" class="total-row"><tr><td>\u2211</td><td class="text-center">\u5408\u8a08</td>${conf.fields.slice(1).map(f => `<td><input type="text" class="t-${f}" readonly value="0"></td>`).join('')}</tr></tfoot></table></div>`;
        container.appendChild(div);
    });
}

function mgr_addRow(type, data = {}) {
    const conf = sectionConfigs.find(c => c.id === type);
    if (!conf) return;
    const tr = document.createElement('tr');
    // No inline onclick/oninput - use data attributes + event delegation
    tr.innerHTML = `<td class="text-center"><button class="delete-btn text-red-400 hover:text-red-600" data-type="${escapeHtml(type)}">\u2715</button></td>` +
        conf.fields.map(f => {
            const isRO = ['gross','opprofit','nonprofit','pretax','net','surplus','nonsurplus','end'].includes(f);
            return `<td data-tip="${isRO ? '\u7cfb\u7d71\u81ea\u52d5\u8a08\u7b97' : '\u652f\u63f4\u5f9e Excel \u8cbc\u4e0a\u6578\u64da'}"><input type="${f === 'name' ? 'text' : 'number'}" class="v-${f} ${f === 'name' ? 'text-left' : ''}" value="${escapeHtml(data[f])}" ${isRO ? 'readonly' : ''}></td>`;
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
                if (!data || typeof data !== 'object') throw new Error('Invalid JSON structure');
                if (!data.sections || !Array.isArray(data.sections)) throw new Error('Invalid sections data');
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
            alert('\u532f\u5165\u5931\u6557\uff1a\u6a94\u6848\u683c\u5f0f\u932f\u8aa4\n' + err.message);
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
            sheetData.push(conf.fields.map(f => fieldNames[f]));
            Array.from(document.querySelectorAll(`#tbody-${conf.id} tr`)).forEach(tr => {
                sheetData.push(conf.fields.map((f, idx) => {
                    const val = tr.querySelector('.v-'+f)?.value || '';
                    return idx === 0 ? val : { v: parseFloat(val) || 0, t: 'n', z: '#,##0' };
                }));
            });
            const footRow = ["\u5408\u8a08"];
            conf.fields.slice(1).forEach(f => footRow.push({
                v: parseFloat(document.querySelector(`#tfoot-${conf.id} .t-${f}`)?.value.replace(/,/g, '')) || 0,
                t: 'n', z: '#,##0'
            }));
            sheetData.push(footRow, []);
        });
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "\u9810\u7b97\u7e3d\u8868");
        XLSX.writeFile(wb, "\u9810\u7b97\u5831\u8868.xlsx");
    } catch (err) {
        alert('\u532f\u51fa\u5931\u6557\uff1a' + err.message);
    }
}

function mgr_exportJSON() {
    try {
        const data = collectData();
        const orgName = data.metadata.org || '\u672a\u547d\u540d';
        saveAs(new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }), `${orgName}_\u9810\u7b97.json`);
    } catch (err) {
        alert('\u532f\u51fa\u5931\u6557\uff1a' + err.message);
    }
}

function mgr_exportHTML() {
    try {
        document.querySelectorAll('input').forEach(i => i.setAttribute('value', i.value));
        saveAs(new Blob(["<!DOCTYPE html>\n" + document.documentElement.outerHTML], { type: "text/html" }), "\u9810\u7b97\u5099\u4efd.html");
    } catch (err) {
        alert('\u532f\u51fa\u5931\u6557\uff1a' + err.message);
    }
}

function mgr_clearAll() {
    if (confirm('\u78ba\u5b9a\u6e05\u7a7a\u6240\u6709\u8cc7\u6599\uff1f\n\uff08\u6b64\u64cd\u4f5c\u7121\u6cd5\u5fa9\u539f\uff09')) {
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
                if (!data || typeof data !== 'object') throw new Error('Invalid JSON structure');
                if (!data.sections || !Array.isArray(data.sections)) throw new Error('Invalid sections data');
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
            alert('\u8655\u7406\u6a94\u6848\u5931\u6557\uff1a' + err.message);
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
        <div class="kpi-card"><div class="text-slate-400 text-sm">\u5df2\u532f\u5165\u6a5f\u95dc</div><div class="kpi-value text-blue-400">${totalOrgs}</div></div>
        <div class="kpi-card"><div class="text-slate-400 text-sm">\u57fa\u91d1\u7e3d\u6578</div><div class="kpi-value text-green-400">${totalFunds}</div></div>
        <div class="kpi-card"><div class="text-slate-400 text-sm">\u6de8\u984d\u5408\u8a08</div><div class="kpi-value ${totalNet < 0 ? 'text-red-400' : 'text-emerald-400'}">${totalNet.toLocaleString()}</div></div>
        <div class="kpi-card"><div class="text-slate-400 text-sm">\u532f\u6574\u6642\u9593</div><div class="kpi-value text-purple-400 text-lg">${new Date().toLocaleTimeString()}</div></div>
    `;

    // No inline onclick - use data-index + event delegation
    document.getElementById('agg-list-body').innerHTML = agg_data.map((d, i) => `
        <tr class="border-b border-slate-700">
            <td class="py-3 text-slate-400">${i + 1}</td>
            <td class="py-3 font-bold">${escapeHtml(d.metadata?.org || '\u672a\u547d\u540d')}</td>
            <td class="py-3">${d.sections?.reduce((sum, s) => sum + (s.items?.length || 0), 0) || 0} \u7b46\u57fa\u91d1</td>
            <td class="py-3"><button class="agg-remove-btn text-red-400 hover:text-red-300" data-index="${i}">\u79fb\u9664</button></td>
        </tr>
    `).join('');
}

function agg_clearAll() {
    agg_data = [];
    document.getElementById('agg-content').classList.add('hidden');
}

// ========== Event Binding (replaces all inline handlers) ==========
function bindStaticEvents() {
    // Tab switching
    document.getElementById('btn-manager').addEventListener('click', () => switchTab('manager'));
    document.getElementById('btn-aggregator').addEventListener('click', () => switchTab('aggregator'));

    // Import
    document.getElementById('btn-import').addEventListener('click', () => {
        document.getElementById('mgr-import-file').click();
    });
    document.getElementById('mgr-import-file').addEventListener('change', function() {
        mgr_handleImport(this.files);
    });

    // Export buttons
    document.getElementById('btn-export-json').addEventListener('click', mgr_exportJSON);
    document.getElementById('btn-export-html').addEventListener('click', mgr_exportHTML);
    document.getElementById('btn-export-xlsx').addEventListener('click', mgr_exportXLSX);

    // Clear
    document.getElementById('btn-clear').addEventListener('click', mgr_clearAll);

    // Aggregator clear
    document.getElementById('btn-agg-clear').addEventListener('click', agg_clearAll);

    // Undo
    document.getElementById('undo-btn').addEventListener('click', undo);

    // Metadata auto-save
    ['mgr-org', 'mgr-year', 'mgr-user'].forEach(id => {
        document.getElementById(id)?.addEventListener('input', debouncedSave);
    });
}

function bindDynamicEvents() {
    const container = document.getElementById('sections-container');

    // Event delegation for add-row and delete buttons
    container.addEventListener('click', e => {
        const addBtn = e.target.closest('.add-row-btn');
        if (addBtn) {
            mgr_addRow(addBtn.dataset.section);
            return;
        }
        const delBtn = e.target.closest('.delete-btn');
        if (delBtn) {
            deleteRow(delBtn, delBtn.dataset.type);
            return;
        }
    });

    // Event delegation for input changes (replaces oninput="handleInput(...)")
    container.addEventListener('input', e => {
        if (e.target.tagName === 'INPUT') {
            const tbody = e.target.closest('tbody');
            if (tbody && tbody.id && tbody.id.startsWith('tbody-')) {
                const type = tbody.id.replace('tbody-', '');
                handleInput(type);
            }
        }
    });

    // Aggregator list - event delegation for remove buttons
    document.getElementById('agg-list-body').addEventListener('click', e => {
        const removeBtn = e.target.closest('.agg-remove-btn');
        if (removeBtn) {
            const idx = parseInt(removeBtn.dataset.index, 10);
            agg_data.splice(idx, 1);
            agg_render();
        }
    });
}

// ========== Keyboard Navigation ==========
function bindKeyboardEvents() {
    // Ctrl+Z global undo
    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
            const el = document.activeElement;
            if (!el || (el.tagName !== 'INPUT' && el.tagName !== 'TEXTAREA')) {
                e.preventDefault();
                undo();
            }
        }
    });

    // Enter key navigation
    document.addEventListener('keydown', e => {
        const el = document.activeElement;
        if (!el || el.tagName !== 'INPUT' || el.readOnly) return;

        const td = el.closest('td');
        const tr = el.closest('tr');
        const tbody = tr?.parentNode;
        if (!td || !tr || !tbody || !tbody.id?.startsWith('tbody-')) return;

        const type = tbody.id.replace('tbody-', '');
        const cells = Array.from(tr.querySelectorAll('td'));
        const colIdx = cells.indexOf(td);
        const rows = Array.from(tbody.children);
        const rowIdx = rows.indexOf(tr);

        if (e.key === 'Enter') {
            e.preventDefault();
            let nextRow = rows[rowIdx + 1];
            if (!nextRow) {
                mgr_addRow(type);
                nextRow = tbody.lastElementChild;
            }
            const nextInput = nextRow?.children[colIdx]?.querySelector('input:not([readonly])');
            if (nextInput) nextInput.focus();
        }
    });

    // Paste handler
    document.addEventListener('paste', e => {
        const el = document.activeElement;
        if (!el || el.readOnly || el.tagName !== 'INPUT') return;
        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (!text.includes('\t') && !text.includes('\n')) return;
        e.preventDefault();
        saveState();
        const rowsData = text.split(/\r\n|\n|\r/).filter(r => r.trim());
        const startTr = el.closest('tr'), tbody = startTr?.parentNode;
        if (!tbody || !tbody.id?.startsWith('tbody-')) return;
        const type = tbody.id.replace('tbody-', '');
        const colIdx = Array.from(startTr.children).indexOf(el.closest('td'));
        const rowIdx = Array.from(tbody.children).indexOf(startTr);
        rowsData.forEach((rowStr, i) => {
            let targetTr = tbody.children[rowIdx + i];
            if (!targetTr) { mgr_addRow(type); targetTr = tbody.lastElementChild; }
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

    const loaded = loadFromStorage();
    if (!loaded) {
        sectionConfigs.forEach(c => mgr_addRow(c.id));
    }
    saveState();
});
