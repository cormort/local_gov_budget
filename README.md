# local_gov_budget — 農業特別收入基金 細項預算查填

純前端的預算書表查填工具：6 個分基金 × 25 張表（業務計畫甲／乙、員額、用人費用、其他管制項目），
可線上編輯、自動加總、填寫核列數、上傳官方 Excel 對帳、匯出 JSON 與 Word（.doc）。

> **表單定義不在這個 repo。** 欄位、表格、基金清單與加總規則的來源是
> [agri_special_fund_forms](https://github.com/cormort/agri_special_fund_forms)，
> 由 `npm run sync:forms` 產生 `vendor/forms/*.js` 後由本專案載入。
> 表單的唯讀線上預覽：<https://cormort.github.io/agri_special_fund_forms/>

---

## 目錄

- [快速開始](#快速開始)
- [主要功能](#主要功能)
- [架構與資料流](#架構與資料流)
- [檔案結構](#檔案結構)
- [資料模型](#資料模型)
- [使用流程](#使用流程)
- [開發指令與驗證](#開發指令與驗證)
- [已知限制](#已知限制)
- [授權](#授權)

---

## 快速開始

需求：現代瀏覽器；**第一次載入需要連得到 CDN**（SheetJS／FileSaver／mammoth／Google Fonts）。

```bash
# 方式 1：直接開檔（file://）—— 預設資料用內嵌的 vendor/forms/sample-data.js
open index.html

# 方式 2：本機靜態伺服器（會改抓根目錄的 sample_fund_*.json，可即時看到資料檔變動）
python3 -m http.server 8000
# → http://127.0.0.1:8000/index.html

# 方式 3：部署到任何靜態空間（GitHub Pages／內部網站皆可，沒有後端需求）
```

`file://` 與 http 的行為差異只有一個：http 會優先用 `sample_fund_*.json`，取不到才用內嵌資料；
其餘功能完全相同。

---

## 主要功能

### 表單

| 分頁 | 內容 | 表數 |
|---|---|---|
| 業務計畫 | 每基金「甲、基金來源」「乙、基金用途」各一張 | 12 |
| 員額用人費 | 甲、員額（六基金共用）＋ 乙、用人費用（每基金一張） | 7 |
| 其他管制項目 | 每基金一張 | 6 |
| 彙編（各分頁的「柒、彙編」） | 跨基金即時彙總（唯讀） | — |

### 編輯

- 單元格直接編輯；數字欄自動加千分號，負值標紅
- **Enter** 移到下一列；**Ctrl/⌘+V** 可從 Excel 貼上整塊資料
- 列的拖曳排序、`＋` 插入列、`✕` 刪除列
- 層級下拉（甲/乙、一、二、…、(1)/(2)），業務計畫的縮排與加總都依層級判定

### 自動計算

- **核列**：主管核列 = 原編數 + 主管增減；院擬核列 = 主管核列 + 院擬增減（即時、唯讀）
- **母列加總**：母列 = 直屬子列的加總（`dec113`、`bud114`、`orig`、`dep_diff`、`gov_diff` 五欄）。
  子列全部沒值時保留母列原值並開放編輯；`收購糧食`、`糧食銷售` 這類「子列是數量／單價」的列不套加總
  （關鍵字清單在 forms repo 的 `rules.noAutoSumKeywords`）

### 上傳官方 Excel（差異預覽）

業務計畫頁工具列的「📊 匯入官方 xlsx」：

1. 拖入官方彙總檔（例：`FNGBRB5300_*.xlsx`）
2. 系統先列出「舊值 → 新值」，不會直接覆蓋
3. 比對結果分四種：**完全相同／別名** 預設勾選；**包含／相似** 需人工確認；**未匹配** 只列出、永不寫入
4. 按「套用」才寫入並自動存檔；可下載 `mapping-report_*.json` 留存對帳紀錄

只補金額（`bud114`），不動表格結構、層級與審核意見。映射規則在 forms repo 的
`profiles/FNGBRB5300.json`。

### 匯入／匯出

| 動作 | 格式 | 說明 |
|---|---|---|
| 匯出 JSON | `.json` | 依分頁輸出，檔名 `{基金}_{年度}_{分頁}.json` |
| 匯出 Doc | `.doc` | Word 可直接開（HTML 包裝），含表頭、審核意見 |
| 匯入 JSON | `.json` | 單檔覆蓋該分頁；或「合併多份 JSON」（用人費／管制按同名加總，業務計畫逐基金追加） |
| 匯入 Word | `.docx` | 解析 Word 表格後帶入業務計畫（另一支「Word 轉 JSON」吃轉換器輸出） |

### 自動存檔

編輯後 800ms 自動寫入瀏覽器 `localStorage`（key：`sf_special_fund_v4`），重新整理會自動還原；
版本升級時會自動遷移舊 key（v1～v3）。清除：右上「🗑️ 清空全部」。

---

## 架構與資料流

```
agri_special_fund_forms（表單定義唯一來源）
  ├─ forms/agri-special-fund.json   欄位 / 25 張表 / 基金 / 加總規則
  ├─ profiles/FNGBRB5300.json       官方 xlsx 映射
  └─ data/sample_fund_*.json        預設資料
            │  npm run sync:forms（tools/sync-forms.mjs）
            ▼
vendor/forms/{agri-special-fund,sample-data,xlsx-profiles}.js   ← 產生檔（已 commit）
            │  window.LGB_FORMS / LGB_SAMPLE_DATA / LGB_XLSX_PROFILES
            ▼
special-fund.js  推導 COL、PLAN_FUNDS、tableConfigs、TAB_INFO 與加總規則
            │
            ├─ index.html / special-fund.css      畫面
            ├─ xlsx-import.js                     上傳 xlsx → 差異預覽 → 套用
            └─ word-converter.js + mammoth        解析 .docx
```

為什麼產生「.js」而不是直接讀 JSON：本專案實務上以 `file://` 開啟，
`fetch()` 讀不到本地 JSON，CSP 也設成 `connect-src 'self'`；用 `<script>` 載入產生檔可同時支援
`file://` 與 http。產生檔檔頭會寫來源 repo 的 commit 短碼，方便追朔。

---

## 檔案結構

```
local_gov_budget/
├── index.html                  # 主畫面（三個分頁 + 六基金子分頁）
├── special-fund.js             # 引擎：表格渲染、加總、核列、存檔、匯入匯出、Doc 產生
├── special-fund.css            # 樣式（含 xlsx 差異預覽視窗）
├── xlsx-import.js              # 上傳 xlsx → 差異預覽（三段式比對）→ 套用
├── word-converter.js           # .docx → JSON（搭配 mammoth.browser.min.js）
├── sample-editor.html/.js      # 「編輯預設樣本」頁：調整 sample_fund_*.json
├── vendor/forms/               # 產生檔（勿手改）：表單定義、預設資料、xlsx profile
├── tools/
│   ├── sync-forms.mjs          # 從 forms repo 產生 vendor/forms（--check 可做漂移檢查）
│   └── verify-xlsx-mapping.mjs # xlsx 映射回歸測試（14 項）
├── sample_fund_*.json          # 六個分基金的預設資料（http 模式優先載入）
├── FNGBRB5300_*.xlsx           # 官方彙總檔範例（測試與匯入用）
├── budget-app.js               # 早期功能模組，目前未被 index.html 載入（保留參考）
├── budget-style.css            # Tailwind 產生的樣式
├── src/input.css, tailwind.config.js
├── package.json                # 僅開發用（Tailwind 建置、xlsx 測試、forms 同步）
├── docs/legacy-excel-rebuild.md# 附錄：2026-05 之前的手工重建流程
├── DECISIONS.md                # 決策紀錄（含未做事項）
└── LICENSE                     # Apache-2.0
```

---

## 資料模型

### 欄位

| 內部 key | 表頭 | 型別 | 說明 |
|---|---|---|---|
| `level` | 層級 | level | 0＝甲/乙、1＝一、二、、2＝計畫名／(一)、3＝1.／2.、4＝(1)／(2) |
| `name` | 項目/科目名稱 | text | 列名 |
| `desc` | 摘要/說明 | textarea | |
| `dec112` | 113 年度決算 | number | 資料保留，但表單目前**不顯示**這欄 |
| `dec113` | 114 年度決算 | number | |
| `bud114` | 115 年度預算 | number | 本表單的主要金額欄（官方檔 `valueKey`） |
| `apr114` | 115 年 4 月底在職 | number | 僅員額表 |
| `orig` | 原編數 | number | |
| `dep_diff` | 主管增減(-) | number | |
| `dep_app` | 主管核列 | number | 自動：`orig + dep_diff` |
| `gov_diff` | 院擬增減(-) | number | |
| `gov_app` | 院擬核列 | number | 自動：`dep_app + gov_diff` |

內部 key 沿用 `dec112/dec113/bud114/apr114` 是為了與舊資料（autosave、既有 JSON）相容，
**表頭文字與年度字樣都在 forms repo**；要換年度請改 forms repo 的 `columns.*.label` 與 `meta.year`。

**JSON 匯入只認內部 key**（`name`、`bud114`…），中文表頭不能用來匯入。下面的對照只是閱讀用：

| 中文表頭（常見寫法） | 內部 key |
|---|---|
| 113 年度決算 ／ 113 決算 | `dec112` |
| 114 年度決算 ／ 114 決算 | `dec113` |
| 115 年度預算 ／ 115 預算 | `bud114` |
| 115 年 4 月底在職 | `apr114` |
| 原編數 ／ 原編金額 ／ 原編員額 | `orig` |
| 主管增減 ／ 主管請增減 | `dep_diff` |
| 院擬增減 ／ 行政院初審增減 | `gov_diff` |

### 單位

| 表 | 單位 |
|---|---|
| 業務計畫（甲／乙）、用人費用、管制項目 | 新臺幣千元 |
| 員額 | 人 |

### 年度字樣出現在三個地方

| 位置 | 現值 | 改哪裡 |
|---|---|---|
| 頁面標題（nav） | 116 年度 | `index.html` |
| 表單標題 | 116 年度 | forms repo `forms/agri-special-fund.json` 的 `meta.year` 與各表 `title` |
| 年度輸入框 | 116 | 使用者在畫面上改（會寫進匯出檔名） |

金額欄標籤（`115 年度預算`）對應官方檔的「本年度預算數」；年度標題與欄位標籤目前不一致是刻意的
（編製 116 年度、金額基準是官方 115 年度預算數），換年度時請三處一起改。

---

## 使用流程

### 編輯與審核

1. 選分頁（業務計畫／員額用人費／管制項目）與基金子分頁
2. 直接改數字，或從 Excel 貼上整塊（`Ctrl/⌘+V`）
3. 填「原編數」→ 主管填「主管增減」→ 自動算出「主管核列」→ 院方填「院擬增減」→ 自動算出「院擬核列」
4. 填審核意見（每個分頁一份，六基金共用）
5. 匯出 JSON 或 Doc；資料同時已自動存在瀏覽器

### 更新官方金額（建議流程）

1. 取得官方彙總檔（FNGBRB5300）
2. 業務計畫頁 →「📊 匯入官方 xlsx」→ 拖入檔案
3. 檢查差異清單（只列有變動與需確認的列）→ 按「套用 N 筆」
4. 需要留紀錄時先按「⬇ 下載映射報告」
5. 長期保存：匯出 JSON 或 Doc

> 匯入後再改資料，請記得**重新匯出 JSON**；localStorage 只是本機暫存。

### 重置與備份

- 各表「重置為 Word 預設項目」：回到該表的預設資料
- 右上「🗑️ 清空全部」：清掉 localStorage 與畫面
- `sample-editor.html`：調整 `sample_fund_*.json` 並匯出，供長期更新預設值

---

## 開發指令與驗證

```bash
npm install                 # 只有開發需要（xlsx 測試、Tailwind）

npm run test:mapping        # xlsx 映射回歸測試（14 項）
npm run sync:forms          # 從 ../agri_special_fund_forms 重新產生 vendor/forms/*.js
npm run sync:forms -- --from /path/to/agri_special_fund_forms
npm run check:forms         # 漂移檢查：vendor 與 forms repo 不一致就 exit 1（可放 CI）
npm run build:css           # 重新產生 budget-style.css（Tailwind）
```

### 改表單或換年度

1. 改 [agri_special_fund_forms](https://github.com/cormort/agri_special_fund_forms)
   （定義、profile 或資料）
2. 在本地跑 `npm run sync:forms`
3. 重新整理頁面；`npm run check:forms` 應通過

`special-fund.js` 不該再出現任何表單定義或預設資料；若需要新增欄位／表格，全部在 forms repo 完成。

### 驗證清單（手動）

- 25 張表都渲染出列、六個基金子分頁可切換
- 修改母列子列後，母列數字隨之更新且變唯讀
- 「📊 匯入官方 xlsx」→ 差異預覽 → 套用後數字更新、重新整理仍在
- 匯出 JSON／Doc 可開啟（Doc 用 Word 或 LibreOffice）

---

## 已知限制

- **純前端、單機**：資料只存在瀏覽器 `localStorage`，換裝置或多人同時編輯要靠匯出／匯入 JSON。
  要做「發布連結、多人回填、稽核」需要後端（見 `DECISIONS.md` 的 P2）。
- **沒有版本歷史**：只能靠每次匯出的 JSON／Doc 留檔。
- **兩張表目前是空的**：`personnel_cost_disaster`、`personnel_cost_fish` 在 forms repo 是空陣列；
  以前是用佔位資料填充，現在不再偽造。要補請改 forms repo 的 `data/`。
- **`dec112`（113 年度決算）不顯示**：資料仍存在於 `sample_fund_*.json`，表單沒有這一欄。
- **依賴 CDN**：SheetJS／FileSaver／mammoth 由 cdnjs 載入，離線時「匯入 xlsx」不可用，其餘功能仍可操作。
- **GA4 追蹤目前無效**：`index.html` 的 GA 是 inline script，但 CSP `script-src` 未允許 `unsafe-inline`，
  且 `connect-src` 未包含 `google-analytics.com`；若需要追蹤請改成外部檔案並放寬 CSP。
- **單頁載入 25 張表**：資料量大時首次渲染會慢（目前 196 列，尚無感）。
- **合併多份 JSON 的規則**：用人費／管制項目按「同名加總」，業務計畫逐基金追加；
  同名但語意不同的列會被誤加總，合併前請先確認。

官方檔結構與早期的手工重建流程（pandas 腳本、欄名對照、÷1000 等）整理在
[`docs/legacy-excel-rebuild.md`](docs/legacy-excel-rebuild.md)，保留作為參考。

---

## 授權

Apache-2.0，見 [`LICENSE`](LICENSE)。官方預算數字取自政府公開資訊，引用請保留原始檔名與年度。

## 相關專案

- [agri_special_fund_forms](https://github.com/cormort/agri_special_fund_forms) — 表單定義、預設資料、官方檔與映射 profile
- 線上表單預覽：<https://cormort.github.io/agri_special_fund_forms/>
