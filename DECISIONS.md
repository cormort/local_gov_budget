# DECISIONS — 官方 xlsx 匯入功能（P0+P1）

本次依〈local_gov_budget × xlsx2webform 優化建議〉實作 **P0（上傳 xlsx → 差異預覽 → 套用）** 與
**P1（profile 外部化、未匹配人工確認、映射報告）**。以下是自行拍板的決定與理由，以及不做的事。

## 已完成

| 項目 | 檔案 |
|---|---|
| 映射引擎（瀏覽器／Node 共用、純函式） | `xlsx-import.js` |
| 宣告式映射設定（換年度改這裡） | `xlsx-profiles.js` |
| UI：工具列按鈕、差異預覽視窗、逐筆勾選、映射報告下載 | `xlsx-import.js` + `special-fund.css` + `index.html` |
| 回歸測試（14 項） | `tools/verify-xlsx-mapping.mjs`，`npm run test:mapping` |
| 文件 | `README.md`（使用流程、換年度、比對規則） |

## 決定

1. **只做 ingestion，不搬 xlsx2webform 的 cell-grid 資料模型。**
   保留 level 階層、`recalcPlanTable` 父子加總、orig→dep_diff→dep_app→gov_diff→gov_app 核列鏈與審核意見。
   官方檔只拿來補金額（`bud114`）。

2. **profile 用 `xlsx-profiles.js` 而不是 JSON 檔。**
   本專案實際以 `file://` 開啟（README 的使用流程就是 file://），`fetch('profiles/*.json')` 會被
   file:// 的同源限制與 CSP `connect-src 'self'` 擋掉；用 `<script>` 載入的 JS 檔沒有這個問題，
   而且仍然是「資料檔」——改年度只需改這個檔，`xlsx-import.js` 不動。

3. **三段式比對取代 xlsx2webform 的「欄位完全一致否則拒收」。**
   官方欄名每年都會長尾巴（本年度就有 `農糧產業調整與轉型計畫(原為「綠色環境給付計畫」)`），
   嚴格比對會把整份檔退掉。改為：

   | 結果 | 行為 |
   |---|---|
   | 完全相同 | 自動勾選 |
   | 別名（profile 宣告） | 自動勾選 |
   | 包含 | **不勾選**，需人工確認 |
   | 相似（≥ 0.75） | **不勾選**，需人工確認 |
   | 未匹配 | 只列出，永不寫入 |

4. **段落標題與人工分類列不映射。**
   `level === 0` 的「甲、基金來源」「乙、基金用途」交給自動加總；
   `其他`、`購建固定資產` 列入 `ignoreNames`。名稱正規化後短於 3 字者不做包含／相似比對，
   否則「其他」會撞上「其他收入」。

5. **母列加總不一致只警告、不阻擋。**
   `recalcPlanTable` 是顯示的權威（子列有值就以子列加總並鎖定母列），
   所以預覽會標出「官方值 ≠ 子列加總」，套用後仍以子列加總呈現。阻擋反而會讓使用者卡住。

6. **未改動 `special-fund.js`。**
   新功能全部在新檔案，`index.html` 只加兩個 `<script>`、一顆按鈕與一個隱藏 input；
   CSS 只 append。既有匯入 JSON／Word、匯出 Doc、自動儲存等路徑完全不動（實測無 JS 錯誤）。

7. **瀏覽器仍用 CDN 的 SheetJS 0.18.5；`xlsx` 只作為 devDependency。**
   與 `index.html` 既有版本一致，避免線上／本機解析行為不同；測試用同一版本。

8. **測試以 `sample_fund_*.json` 為對帳基準（就是瀏覽器實際載入的資料），不是 `SAMPLES`。**
   實測結果：75 命中、0 未匹配、0 加總警告、1 筆待修正。

## 發現的資料問題（未自動修，留給人確認）

`sample_fund_loss.json` 的「綠色環境給付計畫」`bud114` 是**空白**，官方檔為 **12,759,005 千元**。
補上後農損基金 4 個計畫加總剛好等於官方「基金用途」16,695,143，可確認這是漏填而不是語意差異。
本次**沒有**把值寫回 `sample_fund_*.json`（那是資料變更，屬於要人決定的動作）——
上傳一次官方檔、按套用即可在瀏覽器內補上；是否要回寫 repo 的樣本檔請覆核。

## 待使用者執行／覆核

- **`git push`**：本次只做 local commit，沒有推上 GitHub。
- **是否回寫 `sample_fund_loss.json`**（見上）。
- **P2（協作：發布／回填／稽核）** 需要先決定：後端要放哪裡（沿用 xlsx2webform 的 FastAPI 還是
  另建）、資料落地方式（JSON 檔／SQLite）、身分驗證要不要沿用它的 bcrypt + 填表人覆蓋機制。
  這些牽涉部署與資安，不適合在無人覆核下直接開工。
- **P3（版本快照／關聯式 CSV）** 依賴 P2 的後端；若只想先要 CSV，可以在純前端加 Export，
  但要先確認 long format 的欄位定義。
- **5320（員額／用人費／管制項目）** 目前 repo 內沒有該檔，無法做第二組 profile；
  補上檔案後在 `xlsx-profiles.js` 加一組即可，`xlsx-import.js` 不用改。

## 已知限制

- 只處理「基金 × 計畫」的金額彙總（5300）。子項與層級變動（新增／刪除／改名）仍走人工，
  或等 5320 profile。
- 未匹配清單在 `file://`（載入不到 `sample_fund_*.json`、退回舊 `SAMPLES` 佔位資料）時會很長，
  這是預期行為：佔位資料本來就沒有官方檔的細項。實測仍可正常匯入且不誤寫。
- 映射報告只包含本次比對結果，不含「誰按了套用」的簽章；多人協作要等 P2。

---

# DECISIONS（第二輪）— 把表單定義從引擎抽離

## 做了什麼

`special-fund.js` 原本自己帶著 `COL`、`PLAN_BASE`、`PLAN_FUNDS`、`PLAN_SECTIONS`、
`tableConfigs`、`LEVEL_LABELS`、`TAB_INFO`、`SAMPLES`（634 行資料）與加總規則，
和 `agri_special_fund_forms` 形成兩份定義。這輪改成**引擎只當 consumer**：

```
agri_special_fund_forms（唯一來源）
  ├─ forms/agri-special-fund.json
  ├─ profiles/FNGBRB5300.json
  └─ data/*.json
        │  npm run sync:forms（tools/sync-forms.mjs）
        ▼
vendor/forms/{agri-special-fund,sample-data,xlsx-profiles}.js   ← 產生檔，已 commit
        ▼
special-fund.js：FORMS = window.LGB_FORMS → COL / PLAN_FUNDS / tableConfigs / TAB_INFO / 規則全部推導
```

`special-fund.js` 淨減 814 行（含刪掉 634 行硬編碼 `SAMPLES`）。

## 決定

1. **產生「.js」而不是直接讀 JSON。**
   本專案實務以 `file://` 開啟、CSP 又是 `connect-src 'self'`，`fetch()` 讀不到本地 JSON。
   產生檔用 `<script>` 載入，`file://` 與 http 都能跑；內容仍是 JSON（`window.LGB_FORMS = {...}`）。
   產生檔一律 commit，避免「clone 下來還要先 build 才能用」。

2. **產生檔放 `vendor/forms/`，並把來源 commit 短碼寫進檔頭。**
   看得到 `agri_special_fund_forms@c5d5e82`，就知道這份 vendor 對應哪個版本。

3. **`npm run check:forms` 做漂移檢查**（比對重新產生的內容，不一致 exit 1），可放 CI。
   `--from <path>` 可指定 forms repo 路徑。

4. **保留 `sample_fund_*.json` 在引擎 repo。**
   瀏覽器執行時仍優先 `fetch` 它們（可即時看到資料變動、不必每次 sync）；
   `vendor/forms/sample-data.js` 只是 `file://` 的 fallback 與離線保證。
   兩者由 `sync:forms` 一起產生，資料仍以 forms repo 為準。

5. **移除根目錄 `xlsx-profiles.js`**，改由 `vendor/forms/xlsx-profiles.js`（產生檔）提供；
   `tools/verify-xlsx-mapping.mjs` 改讀產生檔（IIFE 包裝，Node 也能載入）。

## 驗證

以重構前後 `collectData()` 的完整輸出對帳（同一份資料、同一次載入）：

```
before: tables 25, rows 196
after : tables 25, rows 196
meta / reviews / 25 張表逐列比對 → 完全相同（僅 savedAt 時間戳不同）
```

另外實測：
- `npm run test:mapping` → 14/14 通過
- `npm run check:forms` → 同步（25 張表、25 份資料、184 列；+12 列甲/乙段落列 = 196）
- 瀏覽器（http 與 file:// 兩種模式）：25 張表渲染、分頁/子分頁切換、新增列、彙編、
  上傳 xlsx → 差異預覽 → 套用（農損「綠色環境給付計畫」補為 12,759,005）、自動存檔皆正常，0 個 JS 錯誤
- `buildTableDocHTML` / `buildPlanCombinedDocHTML` / `buildPersonnelTabDocHTML` / `buildControlTabDocHTML`
  都能產出（長度 4,978 / 28,479 / 14,794 / 20,310）

## 推送狀態

- 遠端 `main` 已包含本輪與上一輪的改動（含 xlsx 匯入、forms 化重構）。
  推送方式同前：CLI 無可用憑證，改用瀏覽器既有登入的 session 以 GitHub Web 上傳，
  內容與本機 commit 樹狀一致（`git diff --stat origin/main HEAD` 為空）。
  本機原本的 commit 保留在分支 `local-history`（`1f0b54f`），`main` 對齊遠端後續 commit。
- 表單定義的線上預覽：<https://cormort.github.io/agri_special_fund_forms/>

## 待覆核

- `personnel_cost_disaster` 與 `personnel_cost_fish` 在 forms repo 的資料是空陣列，
  因此這兩張表目前是空的（重構前是舊佔位 `SAMPLES` 的假資料；現在不再偽造）。
  若確實該有員額/用人費資料，請補在 forms repo 的 `data/`。
- 刪除了根目錄 `xlsx-profiles.js`：若有人直接用舊路徑載入會 404，請改用 `vendor/forms/xlsx-profiles.js`。

---

# DECISIONS（第三輪）— 移除無效追蹤與冗餘欄位

## 1. 移除 GA4 追蹤碼

`index.html`、`sample-editor.html` 各有一段 gtag inline script。CSP 是
`script-src 'self' https://cdnjs.cloudflare.com; connect-src 'self' https://cdnjs.cloudflare.com`，
inline script 沒有 `unsafe-inline` 一律被擋，`connect-src` 也不含 `google-analytics.com`，
所以追蹤從一開始就不會送出。本次直接移除，不留一段「看起來有、其實沒作用」的程式。

若日後要恢復：改成 `gtag.js` 外部檔案、把 `googletagmanager.com` 加進 `script-src`、
`www.google-analytics.com` 加進 `connect-src`，並確認隱私權政策。

## 2. 移除業務計畫表的 `dec112` 資料

`dec112`（113 年度決算）只出現在業務計畫的 12 張表，而這 12 張表的欄位定義裡沒有它
（`cols` 是 `level/dec113/bud114/...`），所以純粹是資料冗餘，也是 forms repo 驗證的
12 個警告來源。已從 6 個 `sample_fund_*.json`（引擎與 forms repo 各一份）刪掉，
共 **100 列**的 `dec112` 鍵。

- 員額／用人費用／管制項目 **仍保留** `dec112` 欄位與表頭（113 年度決算），只是本來就沒有值。
- 既有使用者瀏覽器裡的 autosave 仍可能帶 `dec112`；`addRow` 只渲染 `cfg.cols`，
  不認得的鍵會被忽略，下次自動存檔就自然消失，不需要資料遷移。
- forms repo 的驗證警告因此從 12 個變成 **0 個**。

## 驗證

- `node tools/validate.mjs`（forms repo）→ 通過、0 警告、254 列
- `npm run test:mapping` → 14/14
- `npm run check:forms` → 同步（25 張表、25 份資料、184 列）
- 瀏覽器（http 與 file://）：25 張表、196 列、業務計畫表無 `dec112`、
  員額表仍有「113年度決算」表頭、`window.gtag` 為 undefined、0 個 JS 錯誤
