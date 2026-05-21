# 農業特別收入基金 115 年度預算編製系統

## 📋 專案概述

本專案為農業部農業特別收入基金 115 年度預算編製工具。提供了一套完整的網頁應用，用於管理與編輯 6 個特別收入基金的業務計畫、員額、用人費用及管制項目預算數據。

**基金涵蓋**:
- 農業發展基金
- 漁業發展基金
- 林務發展及造林基金
- 農業天然災害救助基金
- 農產品受進口損害救助基金
- 農村再生基金

---

## 🎯 主要功能

### 1. 業務計畫預算表（Plan Tables）

**6 個基金 × 2 個截面 = 12 張表**

每張表結構：
- **甲、基金來源**：各基金的收入項目及預算
- **乙、基金用途**：各基金的支出計畫項目及預算

**層級結構**：
- L0：區段標題（甲/乙）
- L1-L4：逐級細化的計畫項目

**自動加總**：
- 無子列值時：保留父列預設值（bud114）
- 有子列值時：自動加總父列數值，鎖定為只讀

**預設值來源**：官方 FNGBRB5300 預算汇总檔（xlsx）

| 計畫 | L0 甲、基金來源 | L0 乙、基金用途 |
|---|---:|---:|
| 農業發展基金 | 25,088,694 | 26,027,696 |
| 漁業發展基金 | 980 | 2,300 |
| 林務發展及造林基金 | 1,696,479 | 1,659,962 |
| 農業天然災害救助基金 | 6,215,490 | 6,215,490 |
| 農產品受進口損害救助基金 | 16,095,143 | 16,695,143 |
| 農村再生基金 | 2,913,206 | 17,234,940 |

### 2. 員額表（Headcount）

表格 ID：`headcount`

內容：編制內員額、管理會委員、顧問人員、兼任人員、資本支出

欄位：
- 113 年度決算（dec112）
- 114 年度決算（dec113）
- **115 年度預算**（bud114）
- 115 年 4 月底在職（apr114）
- 原編數（orig）、主管增減、主管核列、院擬增減、院擬核列

### 3. 用人費用表（Personnel Cost）

表格 ID：`personnel_cost`

預設值（8 筆）：
| 科目 | 115 年度預算（千元） |
|---|---:|
| 一、正式員額薪資 | 457 |
| 二、聘僱及兼職人員薪資 | 690 |
| 三、加(夜)班費 | 6,132 |
| 　(一)延長工時加班費 | 4,200 |
| 四、津貼 | 61 |
| 五、獎金 | 86 |
| 六、退休及卹償金 | 42 |
| 八、福利費 | 81 |
| 九、提繳費 | 1 |

### 4. 管制項目表（Control）

表格 ID：`control`

預設值（11 筆）：
| 科目 | 115 年度預算（千元） |
|---|---:|
| 一、水電費 | 30,091 |
| 二、國內旅費 | 18,409 |
| 三、國外旅費 | 135 |
| 五、印刷裝訂費 | 43,623 |
| 六、媒體政策及業務宣導費 | 70,386 |
| 七、推展費 | 228,084 |
| 八、一般服務費（不含計時與計件人員酬金） | 874,115 |
| 九、契約勞力（約用人員） | 1,380 |
| 十、委託調查研究費 | 85,972 |
| 十三、用品消耗 | 12,580,064 |
| 十四、其他費用 | 11,476,160 |
| 十五、補助與捐助 | 114,566,529 |
| 十六、公務車輛 | 61,963 |

---

## 📁 文件結構

```
local_gov_budget/
├── index.html                        # 主入口頁面
├── special-fund.html                 # 特別收入基金編製頁面
├── special-fund.js                   # 核心邏輯 (1767 行)
│   ├── SAMPLES 預設資料定義 (L93)
│   ├── 表格欄位定義 COL (L8-22)
│   ├── 自動加總邏輯 recalcPlanTable (L552-600)
│   ├── 資料載入/還原 (L603-668)
│   ├── JSON 導入/導出 (L672-750+)
│   └── 事件監聽與互動 (L1470+)
├── special-fund.css                  # 樣式表
├── budget-app.js                     # 其他功能模組（預留）
├── budget-style.css                  # 額外樣式
│
├── 📊 資料檔案
├── FNGBRB5300_20260520173956.xlsx   # 【官方汇总預算】各基金業務計畫 115 年度預算數
├── FNGBRB5320_120260520174511.xlsx  # 【官方詳細表】員額、用人費用、管制項目
├── budget_data.json                  # 預算資料（已廢棄，不建議使用）
├── budget_data_converted.json        # 部分轉換資料（不完整）
│
├── tailwind.config.js                # Tailwind CSS 配置
├── package.json                      # 依賴項（Tailwind 構建）
└── README.md                         # 本檔案
```

---

## 💾 數據存儲與管理

### 1. 預設值（SAMPLES）

**位置**：`special-fund.js` 第 93 行 `const SAMPLES`

**結構**：嵌套的 JavaScript 對象與陣列

```javascript
const SAMPLES = {
    plan_agri: [
        { level: 2, name: '甲、基金來源：', bud114: 25088694 },
        { level: 3, name: '一、一般業務計畫' },
        ...
    ],
    plan_fish: [...],
    ...
    personnel_cost: [...],
    control: [...]
};
```

**數據來源**：
- **業務計畫（plan_*_src/use）**：FNGBRB5300_20260520173956.xlsx（官方汇总預算檔）
- **員額/用人費/管制項目**：部分來自 budget_data.json，單位為千元

### 2. 會話資料（Browser LocalStorage）

**自動保存**：當用戶編輯任何欄位時，會自動儲存到瀏覽器 localStorage

**結構**：
```javascript
{
    meta: {
        fund: "農業發展基金",      // 基金名稱
        year: "115",               // 年度
        lastSaved: "2026-05-21"    // 最後儲存時間
    },
    reviews: {
        plan: "主管審核意見",
        personnel: "人事審核意見",
        control: "管制項目審核意見"
    },
    tables: {
        plan_agri_src: [
            { level: 0, name: '甲、基金來源：', bud114: 25088694, ... },
            ...
        ],
        plan_agri_use: [...],
        headcount: [...],
        personnel_cost: [...],
        control: [...]
    }
}
```

### 3. 匯出與匯入

**匯出功能**：
- 格式：JSON
- 檔名：`{基金名稱}_{年度}_細項預算.json`
- 包含：所有表格數據、審核意見、元數據

**匯入功能**：
- 支援 JSON patch 模式（按欄位別名匹配更新）
- 欄位別名對應（例：「115年度預算」→「bud114」）

---

## ⚙️ 技術架構

### 核心欄位定義

```javascript
const COL = {
    dec112: { key: 'dec112', label: '113年度決算', type: 'number' },
    dec113: { key: 'dec113', label: '114年度決算', type: 'number' },
    bud114: { key: 'bud114', label: '115年度預算', type: 'number' },    // ✨ 本年度
    apr114: { key: 'apr114', label: '115年4月底在職', type: 'number' },
    level:  { key: 'level',  label: '層級', type: 'level' },
    name:   { key: 'name',   label: '項目/科目名稱', type: 'text' },
    desc:   { key: 'desc',   label: '摘要/說明', type: 'textarea' },
    orig:   { key: 'orig',   label: '原編數', type: 'number' },
    dep_diff: { key: 'dep_diff', label: '主管增減(-)', type: 'number' },
    dep_app: { key: 'dep_app', label: '主管核列', type: 'number', auto: 'dep_app' },
    gov_diff: { key: 'gov_diff', label: '院擬增減(-)', type: 'number' },
    gov_app: { key: 'gov_app', label: '院擬核列', type: 'number', auto: 'gov_app' }
};
```

### 自動計算邏輯

**recalcRow** — 單列計算
```javascript
dep_app = orig + dep_diff
gov_app = dep_app + gov_diff
```

**recalcPlanTable** — 業務計畫層級加總
```javascript
if (子列有值) {
    父列 = SUM(子列)  // 自動加總，父列鎖定
} else {
    父列 = 保留預設值  // 無子列時，保留 bud114
}
```

### 資料單位

| 表 | 單位 | 備註 |
|---|---|---|
| plan_*_src/use | 新臺幣千元 | 對應「新臺幣千元」標籤 |
| headcount | 人 | 員額人數 |
| personnel_cost | 新臺幣千元 | 用人費用 |
| control | 新臺幣千元 | 管制項目 |

---

## 🚀 使用流程

### 1. 開啟應用

```bash
# 直接在瀏覽器開啟
file:///Users/hsiehminchieh/Dev/Work/local_gov_budget/special-fund.html
```

### 2. 編輯預算

1. 選擇基金與截面（甲/乙）
2. 輸入各項目的預算數字（bud114 欄）
3. **自動保存**至瀏覽器 localStorage

### 3. 審核與核列

- 輸入「原編數」(orig)
- 主管審核後輸入「主管增減」(dep_diff)
- → 自動計算「主管核列」(dep_app = orig + dep_diff)
- 行政院輸入「院擬增減」(gov_diff)
- → 自動計算「院擬核列」(gov_app = dep_app + gov_diff)

### 4. 匯出

點擊「匯出 JSON」按鈕，下載當前編輯的預算數據檔案

### 5. 匯入

選擇 JSON 檔案匯入，系統會：
- 比對欄位別名（中文欄位名稱）
- 只更新存在的欄位（不覆蓋無關數據）
- 保留現有審核意見

---

## 🔧 開發與更新

### 修改預設值

**位置**：`special-fund.js` L93 - L455

**步驟**：
1. 在對應的 `plan_agri` / `plan_fish` 等陣列中找到要修改的列
2. 更改 `bud114` 的值（單位：千元）
3. 保存檔案，瀏覽器重新整理

**範例** — 更新農發基金第一計畫預算：
```javascript
// 原
{ level: 3, name: '一、提升農業經營及發展計畫', bud114: 18683875 },
// 更新為
{ level: 3, name: '一、提升農業經營及發展計畫', bud114: 20000000 },
```

### 新增表格行

```javascript
// 在現有陣列中插入新行
{ level: 3, name: '新增計畫名稱', bud114: 新預算值 }
```

### 新增表格

1. 在 `tableConfigs` 中新增表格配置
2. 在 `SAMPLES` 中新增對應資料
3. 在 HTML 中新增表格容器（ID: `sf-tbody-{table_id}`）

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

## 💡 常見操作

### 重置為預設值

點擊「重置為 Word 預設項目」按鈕 → 確認刪除 → 重新載入 SAMPLES 資料

### 清除所有數據

```javascript
// 在瀏覽器主控台執行
localStorage.clear();
location.reload();
```

### 備份目前編輯

匯出 JSON 檔案，保存到安全位置

### 比對預算變化

- 匯出當前版本的 JSON
- 匯出上一版本的 JSON
- 用 diff 工具比對兩個檔案

---

## 📝 欄位別名對應

用於 JSON 匯入時的欄位識別：

| 中文欄位名稱 | 內部 Key |
|---|---|
| 113年度決算 / 113決算 / 113年度決算數 | dec112 |
| 114年度決算 / 114決算 / 114年度決算數 | dec113 |
| **115年度預算 / 115預算 / 115年度預算數** | **bud114** |
| 115年4月底 / 115年4月底在職 / 115年4月底在職人員 | apr114 |
| 原編 / 原編數 / 原編金額 / 原編員額 | orig |
| 主管增減 / 主管請增減 / 主管增減數 | dep_diff |
| 院擬增減 / 院增減 / 院請增減 / 行政院初審擬增減 / 行政院初審增減 | gov_diff |

---

## ⚠️ 已知限制與注意事項

1. **LocalStorage 限制**：資料儲存在瀏覽器，容量約 5-10MB，跨瀏覽器/裝置不同步
2. **層級加總**：業務計畫表的自動加總規則無法手動覆蓋，若需特殊計算請透過子列輸入
3. **並行編輯**：不支援多人同時編輯同一檔案（建議使用版本控制或 cloud storage）
4. **舊瀏覽器**：需要 ES6 以上支援的現代瀏覽器（Chrome、Safari、Firefox、Edge）

---

## 📞 更新日誌

### 2026-05-21

✅ **初版完成**
- 實現業務計畫、員額、用人費用、管制項目四類表單
- 集成官方預算資料（FNGBRB5300 xlsx）
- 自動加總與審核流程
- JSON 匯出/匯入功能
- 本地儲存與恢復機制

**預設值更新**：
- 從官方 xlsx 提取 15 個業務計畫項目的 115 年度預算數
- 13 個管制項目預算值
- 9 個用人費用科目預算值
- 基金級別總額（6 個基金 × 2 個截面）

---

## 📄 授權與備註

本專案為農業部內部使用工具。  
預設資料來源：農業部農業特別收入基金 115 年度預算官方檔案  
開發日期：2026 年 5 月
