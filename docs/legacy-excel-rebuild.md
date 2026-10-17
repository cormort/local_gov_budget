# 附錄：以手工腳本重建 sample_fund_*.json（legacy）

> **這份文件描述的是 2026-05 之前的手工流程**，保留作為官方檔案結構的參考。
> 現在有兩條更省事的路徑，優先使用它們：
>
> 1. **上傳官方 xlsx**：介面上的「📊 匯入官方 xlsx」，差異預覽後只補金額（見 README）。
> 2. **改 forms repo**：預設資料在 [agri_special_fund_forms](https://github.com/cormort/agri_special_fund_forms) 的
>    `data/sample_fund_*.json`，改完在引擎端跑 `npm run sync:forms`。
>
> 下列內容若提到 `special-fund.js` 的行號或 `SAMPLES`，指的是重構前的版本，
> 現在表單定義與預設資料都已改由 `vendor/forms/*.js`（產生檔）提供。

---

## 📊 數據同步

### 官方 Excel 文件結構

#### FNGBRB5300_20260520173956.xlsx（業務計畫預算汇总檔）

**Sheet**: `sheet1`

**結構**：
- **列（Row）**：各基金及汇总合計
  - Row 0：農業特別收入基金（汇总）
  - Row 1-6：各子基金
    - 　農業發展基金
    - 　漁業發展基金
    - 　林務發展及造林基金
    - 　農業天然災害救助基金
    - 　農產品受進口損害救助基金
    - 　農村再生基金

- **列（Column）**：約 100+ 欄
  - `基金名稱`：基金名稱
  - `{計畫名稱}(本年度預算數)`：各計畫的 115 年度預算
  - 格式範例：
    - `提升農業經營及發展計畫(本年度預算數)`
    - `促進農地利用及農業競爭力計畫(本年度預算數)`
    - `森林遊樂及林業鐵路經營管理計畫(本年度預算數)`

**數據提取規則**：

| 項目 | 規則 |
|---|---|
| 計畫名稱 | 欄位名稱去掉 `(本年度預算數)` 後綴 |
| 基金名稱 | 行標籤，去掉前綴「　」（全寬空格） |
| 預算金額 | 直接讀取單元格值 |
| 單位轉換 | **÷ 1000**（原檔為元，轉為千元） |
| 過濾條件 | 值 > 0 且 != NaN |

**提取範例**：

```
FNGBRB5300.xlsx
行 1（農業發展基金） × 欄「提升農業經營及發展計畫(本年度預算數)」
= 18683875000（元）
÷ 1000 = 18,683,875（千元） ← 寫入 SAMPLES
```

**有效計畫列表**（按 SAMPLES 對應）：

| SAMPLES Key | 計畫名稱（xlsx 欄名） | 單位 |
|---|---|---|
| plan_agri_use | 提升農業經營及發展計畫 | 千元 |
|  | 促進農地利用及農業競爭力計畫 |  |
|  | 增進農民所得及福利計畫 |  |
|  | 輔導菸農轉型與檳榔廢園轉作計畫 |  |
|  | 處理農會漁會信用部計畫 |  |
| plan_fish_use | 漁業發展補助計畫 |  |
| plan_forest_use | 獎勵輔導造林計畫 |  |
|  | 森林遊樂及林業鐵路經營管理計畫 |  |
|  | 山坡地開發利用回饋金繳交管理計畫 |  |
| plan_disaster_use | 農業天然災害救助計畫 |  |
| plan_loss_use | 調整產業或防範措施計畫 |  |
|  | 進口損害救助及穩價計畫 |  |
|  | 農糧產業調整與轉型計畫(原為「綠色環境給付計畫」) |  |
|  | 農產品進口管理計畫 |  |
| plan_renewal_use | 農村再生規劃及人力培育計畫 |  |
|  | 農村再生建設及發展計畫 |  |

#### FNGBRB5320_120260520174511.xlsx（詳細預算表）

**Sheet**: `Sheet1`

**用途**：員額、用人費用、管制項目的詳細預算（目前未完全整合，保留備用）

**結構**：680+ 行，業務計畫交叉表格式

---

### 官方資料更新流程

當收到新的官方預算檔案時：

#### 步驟 1：提取業務計畫數據

```python
import pandas as pd
import json

# 讀取官方汇总檔
df = pd.read_excel('FNGBRB5300_20260520173956.xlsx', sheet_name='sheet1')

# 基金行號對應
fund_rows = {
    '　農業發展基金': 1,
    '　漁業發展基金': 2,
    '　林務發展及造林基金': 3,
    '　農業天然災害救助基金': 4,
    '　農產品受進口損害救助基金': 5,
    '　農村再生基金': 6
}

# 提取所有 "(本年度預算數)" 結尾的欄
budget_cols = [col for col in df.columns if '本年度預算數' in col]

# 逐基金逐計畫提取
for fund_name, row_idx in fund_rows.items():
    for col in budget_cols:
        plan_name = col.replace('(本年度預算數)', '').strip()
        val = df.iloc[row_idx][col]
        
        # 關鍵轉換：除以 1000
        if pd.notna(val) and val > 0:
            val_k = int(val) // 1000
            print(f"{plan_name}: {val_k} (千元)")
```

#### 步驟 2：對應 SAMPLES 並更新

根據「計畫名稱」找出對應的 SAMPLES 行，更新 `bud114` 值

```javascript
// special-fund.js L100 範例
{ level: 3, name: '一、提升農業經營及發展計畫', bud114: 18683875 },
//                                              ↑ 從 xlsx 讀取，÷1000
```

#### 步驟 3：驗證與測試

```bash
node --check special-fund.js  # 語法檢查
```

檢查項目：
- [ ] 所有數值為正整數（無小數）
- [ ] 單位正確（千元）
- [ ] 無 `NaN` 或 `null`
- [ ] 計畫名稱拼寫一致

#### 步驟 4：清除快取

- 使用者需在瀏覽器清除 localStorage 或點「重置為 Word 預設項目」按鈕

#### 步驟 5：驗證結果

在瀏覽器開啟應用，點「重置為 Word 預設項目」，檢查各表的預設值是否正確更新

---

## 🔬 從 Excel 重建 sample_fund_*.json 流程

當有新的官方決算 / 預算 Excel 時，可重跑下列步驟把 6 份 `sample_fund_{code}.json` 整批刷新。每份 JSON 對應一個基金，內含 `plan_{id}_src`、`plan_{id}_use`、`personnel_cost_{code}`、`control_{code}`（`headcount` 維持占位資料；Excel 無此來源）。

### 1. 輸入檔案

兩個資料夾、共 14 個 xlsx：

```
114年度決算/
├── {基金}-餘.xlsx      # 基金來源、用途及餘絀表（業務計畫）
└── {基金}-費.xlsx      # 各項費用彙計表（用人費 + 服務費 + …）
115年度預算/
├── FNGBRB5300_*.xlsx   # 各基金 × 各業務計畫的 115 年預算總額
└── FNGBRB5320_*.xlsx   # 計畫 × 科目交叉表（目前不使用）
```

基金檔名前綴對應：

| 檔名前綴 | code | name |
|---|---|---|
| 農發 | agri | 農業發展基金 |
| 林務 | forest | 林務發展及造林基金 |
| 天災 | disaster | 農業天然災害救助基金 |
| 漁發 | fish | 漁業發展基金 |
| 農損 | loss | 農產品受進口損害救助基金 |
| 再生 | renewal | 農村再生基金 |

### 2. 解析 `{基金}-餘.xlsx`（業務計畫）

**檔案結構**：
- 列：4 行表頭 + 1 行空白 + 「基金來源」段（多個 L1/L2/L3 項目）+ 「基金用途」段 + 「本期賸餘（短絀）」+ 期初/期末基金餘額
- 欄：
  - Col 1 = 項目名稱
  - Col 2 = 本年度預算數（114 預算）— 本系統未使用
  - Col 4 = 本年度決算數（114 決算）→ `dec113`
  - Col 8 = 上年度決算數（113 決算）→ `dec112`

**定位段落**：以 Col 1 字面找出 `基金來源`、`基金用途`、`本期賸餘（短絀）` 三列的 row index：
- 來源段 = `(基金來源, 基金用途)` 之間
- 用途段 = `(基金用途, 本期賸餘)` 之間

**層級偵測（關鍵步驟）**：不要靠「子項加總是否等於父項」推論——數字累加遇到 0 元或四捨五入會錯。改讀 openpyxl 的 `cell.alignment.indent` 屬性，這是 Excel 儲存格的「縮排」格式：

```python
import openpyxl
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
indent = int(ws.cell(row=r, column=1).alignment.indent)  # 0/1/2/3
```

對應規則：

| `indent` 值 | sample JSON `level` | 範例 |
|---|---|---|
| 0 | 段落標題 — 不寫入（L0 由前端自動補入「甲、…」「乙、…」） | 基金來源、基金用途 |
| 1 | L1 | 徵收及依法分配收入 |
| 2 | L2 | 徵收收入 |
| 3 | L3 | 農地變更回饋金收入 |

### 3. 解析 `FNGBRB5300_*.xlsx`（115 年預算對照）

**檔案結構**：1 列表頭 + 7 列資料（總基金 + 6 子基金），約 100 個欄位。

- 第 1 列為欄位標籤，每個都長這樣：`{項目名稱}(本年度預算數)`，例如 `徵收收入(本年度預算數)`、`提升農業經營及發展計畫(本年度預算數)`
- Col 1 是基金名稱，子基金前方有全形空格 `　`

**建索引**：

```python
budget_cols = {}          # 項目名稱 → column index
for c in range(1, ws.max_column + 1):
    h = ws.cell(row=1, column=c).value
    m = re.match(r'^(.+?)\(本年度預算數\)$', str(h))
    if m: budget_cols[m.group(1).strip()] = c

fund_row = {}             # fund code → row index
for r in range(2, ws.max_row + 1):
    nm = str(ws.cell(row=r, column=1).value).strip().lstrip('　 ').strip()
    # ...逐基金比對 name 取得 row
```

**查表**：把 `{基金}-餘.xlsx` 抓到的項目名稱直接拿來 lookup `budget_cols`，找到就把該 fund row × 該 col 的值寫進 `bud114`；找不到就留空（多半發生在 L2/L3 細項，因為 FNGBRB5300 只有 L1 總額）。

### 4. 解析 `{基金}-費.xlsx`（用人費 + 管制項目）

**檔案結構**：表頭 5 行後，按 indent 層級列出費用科目，最後一列是「合計」。

- Col 1 = 科目名稱
- Col 2 = 預算數（114 預算）
- Col 3 = 決算數（114 決算）→ `dec113`

**personnel_cost_{code}**：找到 `用人費用`（indent=0 或 1）那列後，往下找所有 indent 比它大 1 的列當直系子項：

```python
for i, fr in enumerate(fei_rows):
    if fr['name'] == '用人費用':
        base = fr['indent']
        for j in range(i+1, len(fei_rows)):
            if fei_rows[j]['indent'] <= base: break
            if fei_rows[j]['indent'] == base + 1:
                personnel.append(...)
```

**control_{code}**：用名單白名單抓典型管制項目（水電費／郵電費／旅運費／一般服務費／專業服務費／修理保養及保固費／印刷裝訂及公告費／媒體政策及業務宣導費／推展費／保險費／捐助、補助與獎助／補貼、獎勵、慰問、照護與救濟），不論在哪個層級都收進來。

### 5. 單位換算

Excel 內所有金額單位皆為**新臺幣元**，sample JSON 統一存**新臺幣千元**：

```python
def to_qian(v):
    n = float(str(v).replace(',', ''))
    q = n / 1000
    return int(round(q)) if abs(q - round(q)) < 0.0001 else round(q, 2)
```

整數除得盡就用 int（避免 `18683875.0` 這種寫法），有小數則保留兩位。

### 6. 輸出 JSON 結構

最終每份 `sample_fund_{code}.json`：

```json
{
  "meta": { "fund": "...", "year": "116", "org": "農業部", "user": "..." },
  "tables": {
    "plan_{id}_src": [
      { "level": 1, "name": "徵收及依法分配收入",
        "dec112": "1706933.81", "dec113": "1853995.49", "bud114": "630000" },
      { "level": 2, "name": "徵收收入", "dec112": "...", "dec113": "...", "bud114": "..." },
      { "level": 3, "name": "農地變更回饋金收入", ... }
    ],
    "plan_{id}_use": [ ... ],
    "personnel_cost_{code}": [
      { "name": "正式員額薪資", "dec113": "198" },
      { "name": "加（夜）班費",  "dec113": "2809.55" }
    ],
    "control_{code}": [
      { "name": "水電費", "dec113": "24816.55" }
    ],
    "headcount": [ ... 原占位資料保留 ... ]
  }
}
```

**注意事項**：
- 不要在 sample 裡硬塞「甲、基金來源：」「乙、基金用途：」這種 L0 列；它們會在前端載入時由 `ensurePlanSectionHeaders()` 自動補入並做加總。
- `orig`、`dep_diff`、`dep_app`、`gov_diff`、`gov_app` 五個欄位在 sample 中留空（這些是 116 年新的審查流程欄位，無歷史 Excel 來源）。

### 7. 驗證

把 6 份 sample JSON 載入到頁面（用「合併多份 JSON」按鈕一次選 6 個），檢查：
- 業務計畫 tab 切到各基金子 tab，看 L0 列「甲/乙」是否自動加總（114 決算欄應與 Excel 的「基金來源」「基金用途」總額一致，除以 1000 後）
- 員額用人費 tab 看 6 個基金的用人費用子項是否正確
- 管制項目 tab 看 12 項白名單項目是否帶入
- 匯出業務計畫 doc，確認是一張合併 6 基金的大表

---
