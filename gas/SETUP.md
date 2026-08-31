# 部署步驟（GAS Web App + Google Sheets 後端）

## 1. 建立試算表與 Apps Script 專案
1. 到 Google Drive 建立一個新的 Google 試算表，例如命名為「115學年度社團志願選填」。
2. 開啟該試算表 →「擴充功能」→「Apps Script」，會開啟一個綁定此試算表的 Apps Script 專案。
3. 把預設的 `Code.gs` 內容清空，貼上本資料夾的 `Code.gs`。
4. 在專案裡新增檔案 `Setup.gs`，貼上本資料夾的 `Setup.gs`。
5. 再新增一個檔案 `SeedJunior.gs`，貼上本資料夾的 `SeedJunior.gs`。

## 2. 建立分頁結構
1. 在 Apps Script 編輯器上方的函式下拉選單選擇 `initSheets`，按執行（▶）。
2. 第一次執行會跳出授權畫面，選你的 Google 帳號 →「進階」→「前往（專案名稱）」→ 允許。
3. 執行完成後回到試算表，應該會看到 `Students / Clubs / Config / Notices / Responses` 五個分頁，且都已經有標題列。

## 3. 匯入國中部資料（現成資料，一鍵匯入）
1. 函式下拉選單改選 `seedJunior`，執行。
2. 完成後檢查：`Students` 應有 431 筆、`Clubs` 應有 18 筆、`Config`/`Notices` 各多一列 `junior` 的資料。
3. 高中部資料目前尚未提供，先執行 `seedSenior`（只會寫入一筆「尚未開放」的設定與公告，等之後有真正名冊/社團資料時，可以比照 `seedJunior` 的寫法自己加一個 `seedSenior` 版本，或直接在 `Students`/`Clubs`/`Config`/`Notices` 分頁手動貼資料，level 欄位填 `senior` 即可）。

## 4. 部署成 Web App
1. Apps Script 編輯器右上角「部署」→「新增部署作業」。
2. 類型選「網頁應用程式」。
3. 執行身分：**我**。
4. 具有存取權的使用者：**任何人**（學生端不需要登入 Google 帳號就能用）。
5. 按「部署」，複製產生的網址（結尾是 `/exec`）。
6. **之後只要改了 `Code.gs`／`Setup.gs` 的程式碼，記得「新增部署作業」發布新版本，網址通常維持不變，但一定要重新部署程式碼才會生效**（在編輯器裡按執行不會影響已部署的 Web App）。

## 5. 接上前端
1. 打開 `index.html`，把最上面的
   ```js
   const WEBAPP_URL = "REPLACE_WITH_YOUR_DEPLOYED_WEB_APP_URL";
   ```
   換成第 4 步拿到的 `/exec` 網址。
2. 用瀏覽器打開 `index.html`（或部署到 GitHub Pages 之類的靜態空間），選「國中部」，確認：
   - 公告文字、社團清單有正常帶出來（代表 `bootstrap` 有打通）。
   - 走完一次選填流程按「送出」，跳出「送出成功」畫面（代表 `submit` 有打通）。
   - 用同一個學號選「查詢我的選填紀錄」，能看到剛剛送出的志願（代表 `query` 有打通）。
   - 用同一學號再送出一次不同的志願，查詢應該看到最新一次的結果（覆蓋更新）。

## 6. 之後要調整社團容額、公告、開放時間
不用再改程式碼，直接到試算表：
- 社團名額/簡介/費用 → 改 `Clubs` 分頁。
- 開放/截止時間、公告文字、排除說明 → 改 `Config` / `Notices` 分頁。

## 疑難排解
- **fetch 失敗 / CORS 錯誤**：確認部署設定「具有存取權的使用者」是「任何人」而不是「本機構的使用者」；POST 送出時前端已用 `text/plain` 當 Content-Type 避開瀏覽器的 CORS 預檢請求，不要改成 `application/json`。
- **改了程式碼卻沒作用**：多半是忘記「新增部署作業」發布新版本，只存檔或只在編輯器裡執行是不會更新已部署的 Web App 的。
- **doGet/doPost 回傳空白或 500**：到 Apps Script 編輯器「執行項目」（左側時鐘圖示）看錯誤紀錄，或直接在編輯器裡呼叫 `bootstrap_('junior')` 印出結果排查。
