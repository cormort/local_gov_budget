// 範本庫管理：用 GitHub Contents API 直接改 repo 的 forms/（上傳、覆蓋、改名、刪除）與 forms/index.json（分類、說明、年度、截止日、隱藏）
// 權杖只放 sessionStorage（關分頁即清除）；網站本身沒有後端，寫入都是瀏覽器直接呼叫 api.github.com

const MG_KEY = 'local_gov_budget.token';
const mg_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const mg_repoApi = fb_repoUrls(location.hostname, location.pathname).api.replace(/\/contents\/forms$/, '');
const mg_path = name => 'forms/' + encodeURIComponent(name);
let mg_token = '';
let mg_files = [];      // GitHub API 列出的 forms/ 檔案（含 sha）
let mg_index = { templates: {} }, mg_indexSha = null;

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

// 寫檔：sha 有給就是覆蓋，沒給就是新增
const mg_put = (file, contentB64, message, sha) =>
    mg_api('/contents/' + mg_path(file), { method: 'PUT', body: JSON.stringify({ message, content: contentB64, ...(sha ? { sha } : {}) }) });
const mg_delete = (file, sha, message) =>
    mg_api('/contents/' + mg_path(file), { method: 'DELETE', body: JSON.stringify({ message, sha }) });

// ========== 讀取 ==========
async function mg_load() {
    const [list, idx] = await Promise.all([
        mg_api('/contents/forms', { allow404: true }),
        mg_api('/contents/' + mg_path(FB_INDEX), { allow404: true }),
    ]);
    mg_files = (list || []).filter(x => x.type === 'file');
    mg_indexSha = idx ? idx.sha : null;
    mg_index = idx ? JSON.parse(mg_unb64text(idx.content)) : { templates: {} };
    mg_index.templates ||= {};
    mg_render();
}

function mg_render() {
    const entries = fb_libraryEntries(mg_files), t = mg_index.templates;
    const cats = [...new Set(Object.values(t).map(m => m.category).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
    document.getElementById('mg-categories').innerHTML = cats.map(c => `<option value="${mg_esc(c)}">`).join('');
    const input = (name, key, extra = '') => `<input type="text" class="mg-meta text-left border rounded" data-name="${mg_esc(name)}" data-key="${key}" value="${mg_esc(t[name]?.[key])}" ${extra}>`;
    document.getElementById('mg-list').innerHTML = entries.length ? entries.map(e => `<tr>
        <td class="text-left px-2 font-bold whitespace-nowrap">${mg_esc(e.name)}</td>
        <td class="text-left px-2 text-xs">${mg_esc(mg_files.filter(f => f.name.replace(/\.[^.]+$/, '') === e.name && f.name !== FB_INDEX).map(f => f.name).join('、'))}</td>
        <td>${input(e.name, 'category', 'list="mg-categories" placeholder="未分類"')}</td>
        <td>${input(e.name, 'description')}</td>
        <td style="min-width:70px">${input(e.name, 'year')}</td>
        <td><input type="date" class="mg-meta border rounded" data-name="${mg_esc(e.name)}" data-key="deadline" value="${mg_esc(t[e.name]?.deadline)}"></td>
        <td class="text-center"><input type="checkbox" class="mg-meta" data-name="${mg_esc(e.name)}" data-key="hidden"${t[e.name]?.hidden ? ' checked' : ''}></td>
        <td class="whitespace-nowrap px-2">
            <a href="./?form=${encodeURIComponent(e.name)}" target="_blank" class="text-blue-600 text-sm font-bold">開啟</a>
            <button class="mg-rename text-slate-600 text-sm font-bold ml-2" data-name="${mg_esc(e.name)}">改名</button>
            <button class="mg-delete text-red-600 text-sm font-bold ml-2" data-name="${mg_esc(e.name)}">刪除</button>
        </td></tr>`).join('') : '<tr><td colspan="8" class="text-slate-500 p-3">forms/ 裡還沒有範本，請從上面上傳。</td></tr>';
}

// ========== 寫入 ==========
// 表格上的值收回 mg_index（空白的欄位不存，整列都空就刪掉那筆）
function mg_collect() {
    const t = {};
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
    const name = file.name.replace(/\.[^.]+$/, '');
    if (!name || name === 'index' || /[\\/]|\.\./.test(name)) throw new Error('檔名不能當表單名稱');
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
            if (old && !confirm(`「${file.name}」已存在，要覆蓋成新版嗎？（分享連結不變）`)) continue;
            if (twin && /\.xlsx$/i.test(file.name)) bad.push(`${file.name}：已有同名的 ${twin.name}（.json 優先），這份 Excel 不會被使用`);
            await mg_put(file.name, mg_b64(new Uint8Array(await file.arrayBuffer())), `範本庫：${old ? '更新' : '新增'} ${file.name}`, old?.sha);
            done.push(`${file.name}（${info}）`);
        } catch (err) { bad.push(`${file.name}：${err.message}`); }
    }
    await mg_load();
    mg_status([done.length ? `已上傳 ${done.join('、')}` : '', ...bad].filter(Boolean).join('；'), !!bad.length);
}

// 改名 = 用新檔名寫入同樣內容、刪掉舊檔、說明資料搬到新名稱；舊連結會失效
async function mg_rename(name) {
    const to = prompt(`把「${name}」改名為：（舊的分享連結會失效）`, name)?.trim();
    if (!to || to === name) return;
    if (to === 'index' || /[\\/]|\.\./.test(to)) return alert('這個名稱不能用');
    if (fb_libraryEntries(mg_files).some(e => e.name === to)) return alert(`已經有「${to}」`);
    mg_collect();
    for (const f of mg_files.filter(f => f.name.replace(/\.[^.]+$/, '') === name && f.name !== FB_INDEX)) {
        const src = await mg_api('/contents/' + mg_path(f.name));
        const ext = f.name.slice(name.length);
        await mg_put(to + ext, src.content.replace(/\s/g, ''), `範本庫：${f.name} 改名為 ${to + ext}`);
        await mg_delete(f.name, f.sha, `範本庫：${f.name} 改名為 ${to + ext}`);
    }
    if (mg_index.templates[name]) { mg_index.templates[to] = mg_index.templates[name]; delete mg_index.templates[name]; await mg_saveIndex(`範本庫：${name} 改名為 ${to}`); }
    await mg_load();
    mg_status(`已改名為「${to}」，新連結：./?form=${to}`);
}

async function mg_remove(name) {
    const files = mg_files.filter(f => f.name.replace(/\.[^.]+$/, '') === name && f.name !== FB_INDEX);
    if (!confirm(`確定刪除「${name}」（${files.map(f => f.name).join('、')}）？\n分享連結會失效；之後要復原只能從 GitHub 的歷史紀錄找回。\n只是不想列在首頁的話，請改用「隱藏」。`)) return;
    mg_collect();
    for (const f of files) await mg_delete(f.name, f.sha, `範本庫：刪除 ${f.name}`);
    if (mg_index.templates[name]) { delete mg_index.templates[name]; await mg_saveIndex(`範本庫：刪除 ${name} 的說明`); }
    await mg_load();
    mg_status(`已刪除「${name}」`);
}

// ========== 畫面 ==========
function mg_status(text, warn = false) {
    const el = document.getElementById('mg-status');
    el.textContent = text;
    el.className = 'text-sm ' + (warn ? 'text-red-600' : 'text-slate-500');
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
    $('mg-save').onclick = mg_run(async () => { mg_collect(); await mg_saveIndex(); mg_status('已儲存分類與說明（約 1 分鐘後首頁更新）'); });
    $('mg-upload').onchange = mg_run(async e => { const files = [...e.target.files]; e.target.value = ''; await mg_upload(files); });
    $('mg-list').addEventListener('click', mg_run(async e => {
        const b = e.target.closest('button');
        if (b?.classList.contains('mg-rename')) await mg_rename(b.dataset.name);
        if (b?.classList.contains('mg-delete')) await mg_remove(b.dataset.name);
    }));
    let saved = '';
    try { saved = sessionStorage.getItem(MG_KEY) || ''; } catch (e) { /* 同上 */ }
    if (saved) mg_connect(saved);
});
