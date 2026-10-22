# 固定表單範本

此資料夾的每個 `名稱.json` 都是一份固定表單，以下列網址開啟（表頭、公式、預設項目都固定，填表人不能換範本）：

```
https://cormort.github.io/local_gov_budget/form-builder.html?form=名稱
```

## 新增或修改表單

1. 開啟 `form-builder.html`，按「📊 上傳 Excel 範本」產生表單，確認欄位與公式正確（可再增刪預設項目）。
2. 按「📌 匯出範本」，得到 `名稱.json`（名稱＝Excel 檔名）。
3. 將檔案放到本資料夾，commit 並推上 `main`；GitHub Pages 部署完成後網址即生效。

修改同名檔案即更新表單，已開啟舊版的人重新整理即可。範例：`業務計畫範例.json` → `form-builder.html?form=業務計畫範例`。
