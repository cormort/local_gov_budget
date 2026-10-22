# 自行架設：從申請 GitHub 帳號開始

照這份步驟，你會有一個屬於自己的查填網站：

```
https://你的帳號.github.io/local_gov_budget/
```

之後由你自己上傳 Excel 範本、發布表單連結、彙整各機關回傳的 JSON。全程免費、不需伺服器，大約 15 分鐘。

> **誰需要做這些步驟？** 只有「發動調查的承辦人」。填報的各機關不需要 GitHub 帳號，開連結填寫即可。

---

## 步驟 1　申請 GitHub 帳號

1. 開啟 <https://github.com/signup>
2. 依序輸入（也可以按上方「Continue with Google」用 Google 帳號註冊）：
   - **Email**：建議用公務信箱，之後通知與找回密碼都寄到這裡
   - **Password**：至少 15 字元，或至少 8 字元且含數字與小寫字母
   - **Username（帳號）**：**會直接出現在網址裡**，例如帳號 `abc-budget` → 網站是 `abc-budget.github.io`。建議用機關或單位的英文縮寫，只能用英文、數字與 `-`
   - **Your Country/Region**：選 **Taiwan**（必填）
3. 完成人機驗證，到信箱收驗證碼並輸入
4. 方案選 **Free** 即可

> GitHub 可能會要求開啟**雙重驗證（2FA）**，依畫面指示用手機驗證器 App 或簡訊設定即可。

## 步驟 2　複製這個專案到你的帳號（Fork）

1. 登入後開啟 <https://github.com/cormort/local_gov_budget>
2. 按右上角 **Fork**
3. 在「Create a new fork」頁面：
   - **Owner** 選你的帳號
   - **Repository name** 保持 `local_gov_budget`（改名的話，網址也會跟著變）
   - 勾選 **Copy the `main` branch only**
4. 按 **Create fork**

完成後你會在 `https://github.com/你的帳號/local_gov_budget` 看到一份一模一樣的專案。

## 步驟 3　開啟 GitHub Pages（讓專案變成網站）

1. 在你的專案頁面按上方 **Settings**
2. 左側選單點 **Pages**
3. 「Build and deployment」→ **Source** 選 **Deploy from a branch**
4. **Branch** 選 `main`，資料夾選 `/ (root)`，按 **Save**
5. 等 1～2 分鐘，重新整理 Settings → Pages，上方會出現
   「Your site is live at `https://你的帳號.github.io/local_gov_budget/`」

打開那個網址，看到「自訂表單查填」頁面就成功了。

> **如果等了 5 分鐘還沒出現**：到專案的 **Actions** 分頁，若看到「Workflows aren't being run on this forked repository」，按 **I understand my workflows, go ahead and enable them**，再回 Settings → Pages 按一次 Save。

## 步驟 4　（建議）清掉範例表單

`forms/` 資料夾裡有原作者的範例（`業務計畫範例.json`、`測試表.xlsx`）。不需要的話：

1. 進入 `forms/` 資料夾，點開要刪的檔案
2. 按右上角 **⋯** → **Delete file** → **Commit changes**

`forms/README.md` 請保留。

## 步驟 5　發布你的第一份表單

1. 準備 Excel 範本（規則見下方「範本規則」），檔名就是表單名稱，例如 `經費表.xlsx`
2. 開啟 `https://你的帳號.github.io/local_gov_budget/form-builder.html`，按「📊 上傳 Excel 範本」，**先確認欄位、公式欄（🔒）都正確**
3. 回到 GitHub 你的專案 → 點進 `forms` 資料夾 → **Add file** → **Upload files**
4. 把 Excel 拖進去 → 按 **Commit changes**
5. 等約 1 分鐘，把這個連結寄給填報人：

   ```
   https://你的帳號.github.io/local_gov_budget/form-builder.html?form=經費表
   ```

> 也可以在 form-builder 按「📌 匯出範本」→「前往 GitHub 發布 ↗」，上傳匯出的 `.json`；畫面上的「複製連結」會自動帶出**你自己的**網址。

## 步驟 6　收件與彙整

1. 填報人填好後按「💾 匯出 JSON」，用 email 等方式回傳給你
2. 開啟 `https://你的帳號.github.io/local_gov_budget/form-aggregate.html`
3. 把所有 JSON 一次拖進去 → 檢查繳交清單 → 按「📥 匯出 Excel」

## 日常操作

| 想做的事 | 怎麼做 |
|---|---|
| 修改表單 | 改好 Excel，用**同一個檔名**再上傳到 `forms/`（會覆蓋舊版），連結不變 |
| 新增另一份表單 | 用不同檔名上傳到 `forms/`，連結的 `?form=` 換成新檔名 |
| 讓同事也能發布表單 | Settings → **Collaborators** → **Add people**，輸入對方的 GitHub 帳號；對方接受邀請後即可上傳 |
| 取得原作者的新功能 | 專案首頁按 **Sync fork** → **Update branch** |

## 範本規則（Excel）

- 只讀**第一個工作表**；**第一欄是項目名稱**，第一欄空白的列不算項目
- 表頭可以用**合併儲存格**做多層（例：「預算數」橫跨「原編／增減／核列」）；只有一格有字的列（表名、「單位：千元」）會自動略過
- **公式**：同一列的 `+ − × ÷`、括號與 `SUM(B5:D5)` 會變成自動計算欄；跨列、跨工作表、`IF`、`ROUND` 等函數會提示並改為手填
- 有文字的欄是文字欄，空白欄預設為數值欄
- 不支援下拉選單、資料驗證與儲存格格式

## 注意事項

- **表單範本是公開的**：免費帳號的 GitHub Pages 只能用公開專案，`forms/` 裡的 Excel 或 JSON 任何人都看得到。**不要把個資或機敏資料放進範本**。填報人回傳的 JSON 不會經過網站，不受影響。
- **網站本身不收資料**：填報結果只在填表人的電腦和回傳的 JSON 裡，請自行保存。
- **網站更新後**，瀏覽器可能暫存舊版約 10 分鐘，按 `Cmd`／`Ctrl` + `Shift` + `R` 強制重新整理。
- **沒有 GitHub 帳號的同事**也能發動調查：在 form-builder 上傳 Excel → 「📌 匯出範本」，把 `.json` 寄給填報人，對方用「📂 匯入 JSON」開啟同一份表單（只是沒有固定連結）。
