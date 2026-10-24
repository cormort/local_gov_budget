# local_gov_budget — 自訂表單查填

用 GitHub Pages 取代線上表單：上傳 Excel 範本產生表單 → 各機關填寫並匯出 JSON 回傳 → 承辦彙整並匯出 Excel。
全部是靜態頁面，沒有後端；填報資料只存在填表人的瀏覽器與回傳的 JSON 檔。

網站：https://cormort.github.io/local_gov_budget/（首頁就是表單；舊網址 `form-builder.html` 會自動導回首頁並保留 `?form=`）

| 頁面 | 用途 |
|---|---|
| `index.html`（首頁） | 承辦入口：①建立新表單 ②發出查填 ③彙整回傳檔 ④管理範本庫，下方是範本庫。上傳 Excel 範本產生表單、匯出範本、填寫並匯出 JSON |
| `?form=名稱` | 固定表單：讀 `forms/名稱.json`，沒有則讀 `forms/名稱.xlsx` |
| `manage.html` | 範本庫管理：貼上 GitHub 權杖後上傳、覆蓋、改名、刪除 `forms/` 的範本，編輯發起單位、分類、用途、說明、年度、發起日期、截止日、是否隱藏（存在 `forms/index.json`），依日期自動標示調查狀態 |
| `form-aggregate.html` | 匯入多份回傳 JSON 或已填好的 Excel → 繳交清單／依項目彙總／明細 → 匯出 Excel |

發布表單與完整流程見 [`forms/README.md`](forms/README.md)。

**想要有自己的一份？** 從申請 GitHub 帳號、Fork 到開啟 GitHub Pages 的步驟見 [`docs/setup.md`](docs/setup.md)。

## Excel 範本規則

- 範本有多個工作表時，勾選要填寫的工作表（預設第一個沒隱藏的），每個工作表是一個分頁；匯出的 JSON、查填用 Excel、彙整都依工作表分開
- 匯入已填 Excel、彙整時依工作表名稱對應；改過名的工作表只認名稱開頭 4 字相同、欄名一致的（避免把「113非營業(供參)」這類往年資料當成今年的表）
- 常見格式：第 1 列表名、第 2 列單位、第 3 列起是欄位（可多層合併）。表頭上方的表名、單位列會略過並顯示在表格上方：只有一格字、每格都是「標籤：值」（單位：千元、填表機關：），或第一欄橫向合併超過半寬的列都算；第一欄是項目名稱，表頭下方的資料列成為預設項目
- 表頭可用合併儲存格做多層（例：「預算數」橫跨「原編／增減／核列」），欄名串成「預算數－原編」，表單顯示對應的多層表頭；只有一格有字的列（表名、「單位：千元」）會略過
- 同一列的公式（`+ - * /`、括號、`SUM(B2:D2)`）→ 自動計算欄；跨列、跨工作表或其他函數會提示並改為手填
- 有文字資料的欄位是文字欄；空白欄預設為數值欄

## 開發

```bash
npm install
npm test            # 公式轉換與求值的自我檢查
npm run build:css   # 由 src/input.css 重建 budget-style.css（掃描根目錄所有 *.html、*.js）
python3 -m http.server 8000   # http://localhost:8000/
```

## 授權

見 [LICENSE](LICENSE)。
