// 範本庫：列出有空白 Excel（forms/名稱.xlsx）的範本，下載後寄出；貼表時把空白 Excel 跟回傳檔一起拖入、勾選為範本

const home_esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

document.addEventListener('DOMContentLoaded', async () => {
    const list = document.getElementById('home-list');
    const index = await fetch('forms/index.json', { cache: 'no-cache' }).then(r => r.ok ? r.json() : {}).catch(() => ({}));
    const names = Object.entries(index.templates || {}).filter(([, m]) => !m.hidden).map(([n]) => n);
    // ponytail: 只列 forms/index.json 有登記的；範本多到一頁放不下再加搜尋
    const has = await Promise.all(names.map(n => fetch('forms/' + encodeURIComponent(n) + '.xlsx', { method: 'HEAD', cache: 'no-cache' }).then(r => r.ok).catch(() => false)));
    const items = names.filter((_, i) => has[i]).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
    list.innerHTML = items.length ? items.map(n => {
        const m = index.templates[n], q = encodeURIComponent(n);
        return `<li class="flex flex-wrap items-center gap-2 py-3">
            <div class="mr-auto min-w-0"><p class="font-bold text-slate-800 break-all">${home_esc(n)}</p>
                ${m.description ? `<p class="text-sm text-slate-500">${home_esc(m.description)}</p>` : ''}</div>
            <a href="forms/${q}.xlsx" download class="btn btn-primary">下載空白 Excel</a>
        </li>`;
    }).join('') : '<li class="py-2 text-slate-500">還沒有空白 Excel 範本。到「管理範本」上傳。</li>';
});
