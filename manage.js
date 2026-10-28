// 範本與調查管理：用 GitHub Contents API 直接改 repo
//   forms/（範本檔）、forms/index.json（範本的分類、說明、隱藏）、surveys/index.json（調查專案）
// 權杖只放 sessionStorage（關分頁即清除）；網站本身沒有後端，寫入都是瀏覽器直接呼叫 api.github.com

const MG_KEY = 'local_gov_budget.token';
const mg_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const mg_repoApi = fb_repoUrls(location.hostname, location.pathname).api.replace(/\/contents\/forms$/, '');
const mg_siteUrl = fb_repoUrls(location.hostname, location.pathname).site;
const mg_path = name => 'forms/' + encodeURIComponent(name);
const mg_today = () => fb_today();
let mg_token = '';
let mg_files = [];      // GitHub API 列出的 forms/ 檔案（含 sha）
let mg_index = { templates: {} }, mg_indexSha = null;
let mg_sv = { surveys: {} }, mg_svSha = null;

// ========== GitHub API ==========
async function mg_api(path, opts = {}) {
    const res = await fetch(mg_repoApi + path, {
        ...opts,
        headers: { Accept: 'application/vnd.github+json', Authorization: `Bearer ${mg_token}`, ...(opts.body ? { 'Content-Type': 'application/json' } : {}) },
    });
    if (res.status === 404 && opts.allow404) return null;
    if (!res.ok) {
        const msg = (await res.json().catch(() => ({}))).message || res.statusText;
        throw new Error(res.status === 401 ? '權杖無效或已過期' : res.status === 403 ? `沒有寫入權限（${msg}）` : `GitHub ${res.status}：${msg}`);
    }
    return res.status === 204 ? null : res.json();
}

const mg_b64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s); };
const mg_b64text = text => mg_b64(new TextEncoder().encode(text));
const mg_unb64text = b64 => new TextDecoder().decode(Uint8Array.from(atob(b64.replace(/\s/g, '')), c => c.charCodeAt(0)));

// 寫檔（path 是 repo 內路徑，各段已編碼）：sha 有給就是覆蓋，沒給就是新增
const mg_putPath = (path, contentB64, message, sha) =>
    mg_api('/contents/' + path, { method: 'PUT', body: JSON.stringify({ message, content: contentB64, ...(sha ? { sha } : {}) }) });
const mg_put = (file, contentB64, message, sha) => mg_putPath(mg_path(file), contentB64, message, sha);
const mg_delete = (file, sha, message) =>
    mg_api('/contents/' + mg_path(file), { method: 'DELETE', body: JSON.stringify({ message, sha }) });
const mg_svPath = FB_SURVEY_INDEX.split('/').map(encodeURIComponent).join('/');

// ========== 讀取 ==========
async function mg_load() {
    const [list, idx, sv] = await Promise.all([
        mg_api('/contents/forms', { allow404: true }),
        mg_api('/contents/' + mg_path(FB_INDEX), { allow404: true }),
        mg_api('/contents/' + mg_svPath, { allow404: true }),
    ]);
    mg_files = (list || []).filter(x => x.type === 'file');
    mg_indexSha = idx ? idx.sha : null;
    mg_index = idx ? JSON.parse(mg_unb64text(idx.content)) : { templates: {} };
    mg_index.templates ||= {};
    mg_svSha = sv ? sv.sha : null;
    mg_sv = sv ? JSON.parse(mg_unb64text(sv.content)) : { surveys: {} };
    mg_sv.surveys ||= {};
    mg_render();
    mg_renderSurveys();
}

const mg_templateNames = () => fb_libraryEntries(mg_files).map(e => e.name);
const mg_usedBy = form => Object.entries(mg_sv.surveys).filter(([, s]) => s.form === form).map(([id]) => id);
const mg_badge = (st, text) => `<span class="badge ${FB_BADGE[st] || 'text-slate-600'}">${text}</span>`;
const mg_copyBtn = url => `<div class="flex gap-1 items-center"><code class="text-xs text-slate-500 break-all">${mg_esc(decodeURIComponent(url))}</code>
    <button class="mg-copy btn-link whitespace-nowrap" data-url="${mg_esc(url)}">複製</button></div>`;
// 有用過的值做成下拉建議
const mg_options = (values, id) => {
    document.getElementById(id).innerHTML = [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant')).map(c => `<option value="${mg_esc(c)}">`).join('');
};

// ========== 範本 ==========
function mg_render() {
    const t = mg_index.templates;
    const q = document.getElementById('mg-q').value, onlyHidden = document.getElementById('mg-only-hidden').checked;
    const all = fb_libraryEntries(mg_files);
    const entries = all.filter(e => fb_libraryMatch({ ...e, meta: t[e.name] }, q) && (!onlyHidden || t[e.name]?.hidden));
    mg_options(Object.values(t).map(m => m.category), 'mg-categories');
    const field = (name, key, label, extra = '') => `<label class="text-xs text-slate-500">${label}
        <input type="text" class="mg-meta block w-full border rounded p-1 text-sm text-slate-800" data-name="${mg_esc(name)}" data-key="${key}" value="${mg_esc(t[name]?.[key])}" ${extra}></label>`;
    document.getElementById('mg-list').innerHTML = entries.length ? entries.map(e => {
        const m = t[e.name] || {}, used = mg_usedBy(e.name);
        return `<div class="border rounded-md p-4 mb-3 ${m.hidden ? 'bg-slate-50' : 'bg-white'}">
        <div class="flex flex-wrap items-start gap-2 mb-3">
            <div class="mr-auto min-w-0">
                <span class="font-bold text-slate-800">${mg_esc(e.name)}</span>
                ${m.hidden ? mg_badge('', '隱藏') : ''}
                <span class="text-xs text-slate-400 ml-1">${mg_esc(mg_files.filter(f => f.name.replace(/\.[^.]+$/, '') === e.name && f.name !== FB_INDEX).map(f => f.name).join('、'))}</span>
                <div class="text-xs text-slate-500">${used.length ? `用於調查：${used.map(mg_esc).join('、')}` : '尚未用於任何調查'}</div>
            </div>
            <div class="whitespace-nowrap">
                <button class="mg-new-from btn-link" data-name="${mg_esc(e.name)}">以此發起調查</button>
                <a href="advanced.html?form=${encodeURIComponent(e.name)}" target="_blank" class="btn-link text-slate-600 ml-2">預覽</a>
                <button class="mg-rename btn-link text-slate-600 ml-2" data-name="${mg_esc(e.name)}">改名</button>
                <button class="mg-delete btn-link text-red-600 ml-2" data-name="${mg_esc(e.name)}">刪除</button>
            </div>
        </div>
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3">
            ${field(e.name, 'category', '分類', 'list="mg-categories" placeholder="未分類"')}
            <div class="col-span-2">${field(e.name, 'description', '說明')}</div>
            <label class="text-xs text-slate-500 flex items-end gap-1 pb-1"><input type="checkbox" class="mg-meta" data-name="${mg_esc(e.name)}" data-key="hidden"${m.hidden ? ' checked' : ''}> 隱藏（不列在範本庫）</label>
        </div></div>`;
    }).join('') : `<p class="text-slate-500 p-3">${all.length ? '沒有符合的範本。' : 'forms/ 裡還沒有範本，請從上面上傳。'}</p>`;
    document.getElementById('mg-count').textContent = `${entries.length}／${all.length} 份`;
    document.getElementById('mg-new-form').innerHTML = mg_templateNames().map(n => `<option>${mg_esc(n)}</option>`).join('');
}

// 範本卡片上的值收回 mg_index（空白的欄位不存）；查詢時沒顯示的範本保留原值
function mg_collect() {
    const t = { ...mg_index.templates }, shown = new Set([...document.querySelectorAll('.mg-meta')].map(el => el.dataset.name));
    shown.forEach(n => delete t[n]);
    document.querySelectorAll('.mg-meta').forEach(el => {
        const v = el.type === 'checkbox' ? el.checked : el.value.trim();
        if (!v) return;
        (t[el.dataset.name] ||= {})[el.dataset.key] = v;
    });
    mg_index.templates = t;
}

async function mg_saveIndex(message = '範本庫：更新分類與說明') {
    const r = await mg_put(FB_INDEX, mg_b64text(JSON.stringify(mg_index, null, 2) + '\n'), message, mg_indexSha);
    mg_indexSha = r.content.sha;
}

// 上傳前先確認是能用的表單，壞檔不發布
async function mg_check(file) {
    const bad = fb_badName(file.name.replace(/\.[^.]+$/, ''));
    if (bad) throw new Error('檔名不能當範本名稱：' + bad);
    if (/\.json$/i.test(file.name)) {
        const form = fb_normalize(JSON.parse(await file.text()));
        return `${form.sheets.length} 個工作表`;
    }
    const p = fb_readWorkbook(await file.arrayBuffer(), file.name);
    return `工作表「${p.sheet}」${p.template.columns.length} 欄`;
}

async function mg_upload(files) {
    const done = [], bad = [];
    for (const file of files) {
        try {
            const info = await mg_check(file);
            const name = file.name.replace(/\.[^.]+$/, '');
            const old = mg_files.find(f => f.name === file.name);
            const twin = mg_files.find(f => f.name !== file.name && f.name.replace(/\.[^.]+$/, '') === name);
            if (old && !confirm(`「${file.name}」已存在，要覆蓋成新版嗎？（用它的調查專案會改用新版）`)) continue;
            await mg_put(file.name, mg_b64(new Uint8Array(await file.arrayBuffer())), `範本庫：${old ? '更新' : '新增'} ${file.name}`, old?.sha);
            done.push(`${file.name}（${info}${twin && /\.xlsx$/i.test(file.name) ? `；網頁表單仍以同名 ${twin.name} 為準，這份 Excel 用於首頁的空白 Excel 與貼表` : ''}）`);
        } catch (err) { bad.push(`${file.name}：${err.message}`); }
    }
    await mg_load();
    mg_status([done.length ? `已上傳 ${done.join('、')}` : '', ...bad].filter(Boolean).join('；'), !!bad.length);
}

// 範本改名 = 用新檔名寫入同樣內容、刪掉舊檔、說明資料搬到新名稱；用它的調查專案一併改指向新名稱
async function mg_rename(name) {
    const to = prompt(`把範本「${name}」改名為：（範本的預覽網址會變；用它的調查專案會自動改用新名稱）`, name)?.trim();
    if (!to || to === name) return;
    const bad = fb_badName(to);
    if (bad) return alert(bad);
    if (mg_templateNames().includes(to)) return alert(`已經有「${to}」`);
    mg_collect();
    mg_collectSv();
    for (const f of mg_files.filter(f => f.name.replace(/\.[^.]+$/, '') === name && f.name !== FB_INDEX)) {
        const src = await mg_api('/contents/' + mg_path(f.name));
        const ext = f.name.slice(name.length);
        await mg_put(to + ext, src.content.replace(/\s/g, ''), `範本庫：${f.name} 改名為 ${to + ext}`);
        await mg_delete(f.name, f.sha, `範本庫：${f.name} 改名為 ${to + ext}`);
    }
    if (mg_index.templates[name]) { mg_index.templates[to] = mg_index.templates[name]; delete mg_index.templates[name]; await mg_saveIndex(`範本庫：${name} 改名為 ${to}`); }
    const used = mg_usedBy(name);
    if (used.length) { used.forEach(id => { mg_sv.surveys[id].form = to; }); await mg_saveSurveys(`調查專案：範本 ${name} 改名為 ${to}`); }
    await mg_load();
    mg_status(`已改名為「${to}」${used.length ? `，調查 ${used.join('、')} 已改用新名稱` : ''}`);
}

async function mg_remove(name) {
    const used = mg_usedBy(name);
    if (used.length) return alert(`範本「${name}」還有調查專案在用：${used.join('、')}\n請先刪除這些專案，或把範本設為「隱藏」。`);
    const files = mg_files.filter(f => f.name.replace(/\.[^.]+$/, '') === name && f.name !== FB_INDEX);
    if (!confirm(`確定刪除範本「${name}」（${files.map(f => f.name).join('、')}）？\n之後要復原只能從 GitHub 的歷史紀錄找回。只是不想列出的話，請改用「隱藏」。`)) return;
    mg_collect();
    for (const f of files) await mg_delete(f.name, f.sha, `範本庫：刪除 ${f.name}`);
    if (mg_index.templates[name]) { delete mg_index.templates[name]; await mg_saveIndex(`範本庫：刪除 ${name} 的說明`); }
    await mg_load();
    mg_status(`已刪除範本「${name}」`);
}

// ========== 調查專案 ==========
const mg_lines = text => text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

function mg_renderSurveys() {
    const all = fb_surveyEntries(mg_sv, true), today = mg_today();
    const q = document.getElementById('mg-sv-q').value, st = document.getElementById('mg-sv-status').value;
    const entries = all.filter(e => fb_libraryMatch(e, q, '', st, today));
    mg_options(Object.values(mg_sv.surveys).map(s => s.unit), 'mg-units');
    const forms = mg_templateNames();
    const input = (id, key, label, type = 'text', extra = '') => `<label class="text-xs text-slate-500">${label}
        <input type="${type}" class="mg-sv block w-full border rounded p-1 text-sm text-slate-800" data-id="${mg_esc(id)}" data-key="${key}" value="${mg_esc(mg_sv.surveys[id][key])}" ${extra}></label>`;
    document.getElementById('mg-sv-list').innerHTML = entries.length ? entries.map(e => {
        const m = e.meta, status = fb_libraryStatus(m, today), url = mg_siteUrl + '?survey=' + encodeURIComponent(e.name);
        const missing = m.form && !forms.includes(m.form);
        return `<div class="border rounded-md p-4 mb-3 ${m.hidden ? 'bg-slate-50' : 'bg-white'}">
        <div class="flex flex-wrap items-start gap-2 mb-3">
            <div class="mr-auto min-w-0">
                <span class="font-bold text-slate-800">${mg_esc(m.title || e.name)}</span>
                ${status ? mg_badge(status, FB_STATUS[status]) : ''}${m.hidden ? mg_badge('', '隱藏') : ''}
                <span class="text-xs text-slate-400 ml-1">代號 ${mg_esc(e.name)}</span>
                ${missing ? `<div class="text-xs text-red-600 font-bold">⚠️ 範本「${mg_esc(m.form)}」不存在，填報人會打不開</div>` : ''}
                ${mg_copyBtn(url)}
            </div>
            <div class="whitespace-nowrap">
                <a href="advanced.html?survey=${encodeURIComponent(e.name)}" target="_blank" class="btn-link">開啟</a>
                <a href="form-aggregate.html?survey=${encodeURIComponent(e.name)}" target="_blank" class="btn-link ml-2">彙整</a>
                <button class="mg-sv-end btn-link text-slate-600 ml-2" data-id="${mg_esc(e.name)}" title="截止日設為今天">提前結束</button>
                <button class="mg-sv-del btn-link text-red-600 ml-2" data-id="${mg_esc(e.name)}">刪除</button>
            </div>
        </div>
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3">
            <div class="col-span-2">${input(e.name, 'title', '調查名稱')}</div>
            <label class="text-xs text-slate-500">範本<select class="mg-sv block w-full border rounded p-1 text-sm text-slate-800" data-id="${mg_esc(e.name)}" data-key="form">
                ${[...new Set([...forms, m.form].filter(Boolean))].map(n => `<option${n === m.form ? ' selected' : ''}>${mg_esc(n)}</option>`).join('')}</select></label>
            ${input(e.name, 'unit', '發起單位', 'text', 'list="mg-units"')}
            ${input(e.name, 'start', '發起日期', 'date')}
            ${input(e.name, 'deadline', '截止日', 'date')}
            ${input(e.name, 'year', '年度')}
            <label class="text-xs text-slate-500 flex items-end gap-1 pb-1"><input type="checkbox" class="mg-sv" data-id="${mg_esc(e.name)}" data-key="hidden"${m.hidden ? ' checked' : ''}> 隱藏（首頁不列，連結照樣能用）</label>
            <div class="col-span-2 md:col-span-4">${input(e.name, 'purpose', '用途')}</div>
            <label class="text-xs text-slate-500 col-span-2 md:col-span-4">應繳機關（每行一個；${(m.respondents || []).length} 個）
                <textarea rows="2" class="mg-sv block w-full border rounded p-1 text-sm text-slate-800" data-id="${mg_esc(e.name)}" data-key="respondents">${mg_esc((m.respondents || []).join('\n'))}</textarea></label>
        </div></div>`;
    }).join('') : `<p class="text-slate-500 p-3">${all.length ? '沒有符合的調查。' : '還沒有調查專案，請從上面「發起新調查」。'}</p>`;
    document.getElementById('mg-sv-count').textContent = `${entries.length}／${all.length} 個`;
}

// 專案卡片上的值收回 mg_sv（空白欄位不存）；查詢時沒顯示的專案保留原值
function mg_collectSv() {
    document.querySelectorAll('.mg-sv').forEach(el => {
        const s = mg_sv.surveys[el.dataset.id], key = el.dataset.key;
        if (!s) return;
        const v = el.type === 'checkbox' ? el.checked : key === 'respondents' ? mg_lines(el.value) : el.value.trim();
        if (v && (!Array.isArray(v) || v.length)) s[key] = v; else delete s[key];
    });
}

async function mg_saveSurveys(message = '調查專案：更新') {
    const r = await mg_putPath(mg_svPath, mg_b64text(JSON.stringify(mg_sv, null, 2) + '\n'), message, mg_svSha);
    mg_svSha = r.content.sha;
}

async function mg_createSurvey() {
    const $ = id => document.getElementById(id);
    const id = $('mg-new-id').value.trim(), form = $('mg-new-form').value;
    const bad = fb_badName(id);
    if (bad) return alert('調查代號：' + bad);
    if (mg_sv.surveys[id]) return alert(`已經有代號「${id}」的調查`);
    if (!form) return alert('請先選範本');
    mg_collectSv();
    const s = { title: $('mg-new-title').value.trim() || id, form };
    ['unit', 'purpose', 'year', 'start', 'deadline'].forEach(k => { const v = $('mg-new-' + k).value.trim(); if (v) s[k] = v; });
    const resp = mg_lines($('mg-new-resp').value);
    if (resp.length) s.respondents = resp;
    mg_sv.surveys[id] = s;
    await mg_saveSurveys(`調查專案：新增 ${id}`);
    ['id', 'title', 'unit', 'purpose', 'year', 'start', 'deadline', 'resp'].forEach(k => { $('mg-new-' + k).value = ''; });
    mg_renderSurveys();
    mg_render(); // 範本卡片的「用於調查」
    mg_status(`已建立調查「${s.title}」，查填網址：${mg_siteUrl}?survey=${id}（約 1 分鐘後生效）`);
}

async function mg_endSurvey(id) {
    if (!confirm(`把調查「${mg_sv.surveys[id].title || id}」的截止日設為今天？`)) return;
    mg_collectSv();
    mg_sv.surveys[id].deadline = mg_today();
    await mg_saveSurveys(`調查專案：${id} 提前結束`);
    mg_renderSurveys();
    mg_status(`已結束調查「${id}」`);
}

async function mg_deleteSurvey(id) {
    if (!confirm(`確定刪除調查「${mg_sv.surveys[id].title || id}」？\n查填網址會失效（範本不受影響）。只是不想列出的話，請改用「隱藏」。`)) return;
    mg_collectSv();
    delete mg_sv.surveys[id];
    await mg_saveSurveys(`調查專案：刪除 ${id}`);
    mg_renderSurveys();
    mg_render();
    mg_status(`已刪除調查「${id}」`);
}

// ========== 畫面 ==========
function mg_status(text, warn = false) {
    const el = document.getElementById('mg-status');
    el.textContent = text;
    el.className = 'text-sm ' + (warn ? 'text-red-600' : 'text-slate-500');
}

// 從範本庫「以此範本發起調查」進來（?form=範本）：選好範本、捲到發起新調查
function mg_startFrom(form) {
    if (!form || !mg_templateNames().includes(form)) return;
    document.getElementById('mg-new-form').value = form;
    document.getElementById('surveys').scrollIntoView();
    document.getElementById('mg-new-id').focus();
}

async function mg_connect(token) {
    mg_token = token;
    mg_status('連線中…');
    try {
        const repo = await mg_api('');
        if (!repo.permissions?.push) throw new Error('這個權杖沒有寫入權限（Contents 要設 Read and write）');
        try { sessionStorage.setItem(MG_KEY, token); } catch (e) { /* 無痕模式等存不了就每次重貼 */ }
        document.getElementById('mg-main').classList.remove('hidden');
        document.getElementById('mg-logout').classList.remove('hidden');
        document.getElementById('mg-token').value = '';
        await mg_load();
        mg_status(`已連線 ${repo.full_name}`);
        mg_startFrom(new URLSearchParams(location.search).get('form'));
    } catch (err) {
        mg_token = '';
        document.getElementById('mg-main').classList.add('hidden');
        mg_status('⚠️ ' + err.message, true);
    }
}

// 每個寫入動作：鎖住畫面避免重複按，錯誤顯示在狀態列
const mg_run = fn => async (...a) => {
    document.body.style.pointerEvents = 'none';
    try { await fn(...a); } catch (err) { mg_status('⚠️ ' + err.message, true); }
    finally { document.body.style.pointerEvents = ''; }
};

document.addEventListener('DOMContentLoaded', () => {
    const $ = id => document.getElementById(id);
    $('mg-repo-name').textContent = mg_repoApi.replace('https://api.github.com/repos/', '');
    $('mg-connect').onclick = () => { const t = $('mg-token').value.trim(); if (t) mg_connect(t); };
    $('mg-token').addEventListener('keydown', e => { if (e.key === 'Enter') $('mg-connect').click(); });
    $('mg-logout').onclick = () => {
        try { sessionStorage.removeItem(MG_KEY); } catch (e) { /* 同上 */ }
        mg_token = '';
        $('mg-main').classList.add('hidden');
        $('mg-logout').classList.add('hidden');
        mg_status('已清除權杖');
    };
    $('mg-reload').onclick = mg_run(mg_load);
    $('mg-save').onclick = mg_run(async () => { mg_collect(); await mg_saveIndex(); mg_render(); mg_status('已儲存範本分類與說明（約 1 分鐘後首頁更新）'); });
    $('mg-upload').onchange = mg_run(async e => { const files = [...e.target.files]; e.target.value = ''; await mg_upload(files); });
    $('mg-sv-save').onclick = mg_run(async () => { mg_collectSv(); await mg_saveSurveys(); mg_renderSurveys(); mg_status('已儲存調查專案（約 1 分鐘後首頁更新）'); });
    $('mg-new-create').onclick = mg_run(mg_createSurvey);
    // 標準機關名單（surveys/orgs.json，本站靜態檔，不需要權杖）
    $('mg-new-std').onclick = mg_run(async () => {
        const res = await fetch(FB_ORGS, { cache: 'no-cache' });
        if (!res.ok) throw new Error(`讀不到 ${FB_ORGS}（${res.status}）`);
        const list = (await res.json()).sortOrder || [];
        const box = $('mg-new-resp');
        if (box.value.trim() && !confirm(`應繳機關已有內容，要換成標準名單（${list.length} 個）嗎？`)) return;
        box.value = list.join('\n');
        mg_status(`已帶入標準名單 ${list.length} 個機關，請刪掉這次不需要的`);
    });
    // 搜尋前先把卡片上改到一半的值收回，免得重畫時不見
    $('mg-q').addEventListener('input', () => { mg_collect(); mg_render(); });
    $('mg-only-hidden').addEventListener('change', () => { mg_collect(); mg_render(); });
    $('mg-sv-q').addEventListener('input', () => { mg_collectSv(); mg_renderSurveys(); });
    $('mg-sv-status').addEventListener('change', () => { mg_collectSv(); mg_renderSurveys(); });
    document.addEventListener('click', async e => {
        const c = e.target.closest('.mg-copy');
        if (!c) return;
        await navigator.clipboard.writeText(c.dataset.url).catch(() => {});
        c.textContent = '已複製 ✓';
        setTimeout(() => { c.textContent = '複製'; }, 1500);
    });
    $('mg-list').addEventListener('click', mg_run(async e => {
        const b = e.target.closest('button');
        if (b?.classList.contains('mg-new-from')) mg_startFrom(b.dataset.name);
        if (b?.classList.contains('mg-rename')) await mg_rename(b.dataset.name);
        if (b?.classList.contains('mg-delete')) await mg_remove(b.dataset.name);
    }));
    $('mg-sv-list').addEventListener('click', mg_run(async e => {
        const b = e.target.closest('button');
        if (b?.classList.contains('mg-sv-end')) await mg_endSurvey(b.dataset.id);
        if (b?.classList.contains('mg-sv-del')) await mg_deleteSurvey(b.dataset.id);
    }));
    let saved = '';
    try { saved = sessionStorage.getItem(MG_KEY) || ''; } catch (e) { /* 同上 */ }
    if (saved) mg_connect(saved);
});
