# 採購工作室

供應商專屬採購單與進度追蹤網站。選擇供應商，上傳客戶需求後直接套用原 Word 格式；九家供應商的原稿版面維持不變。興普、優技、采鑫追蹤表按各自原 Excel 格式匯出。

訂單號碼格式為 MH26A-10-035：年月依 Asia/Taipei 更新，尾碼自行填寫。介面採深藍及灰紫色系。

## 雲端同步

網站由 GitHub Pages 提供，Firebase Authentication 在背景自動建立匿名身分，沒有登入表單。Cloud Firestore 保存工作區；不使用 Analytics 或 Firebase Storage。

首次成功連線會從目前瀏覽器自動移轉供應商、模板、草稿、需求、採購紀錄與追蹤資料。Word 草稿編輯停止約一秒後自動儲存；其他表單按原本的儲存按鈕。狀態列明確區分本機保留、同步中、完成與版本衝突。

頁面首次連線後會產生帶有 #workspace=... 的專屬連結。使用「複製專屬連結」換到其他電腦即可讀取相同工作區。連結使用瀏覽器產生的 256-bit 隨機值，是讀寫資料的憑證，持有者皆可存取；不得放在公開儲存庫、截圖或公開網站中。Firestore 規則拒絕列出工作區，僅允許已取得完整隨機路徑的匿名身分讀寫。連結放在 fragment，並使用 no-referrer。

檔案及工作區 JSON 壓縮後分成每段最多 500,000 字元的文件，校驗 SHA-256 後才套用。先上傳新版本的片段，再以交易比對版號後更新工作區指標，避免半套資料及舊裝置覆寫新版本。成功後清除前版片段；中斷造成的未引用片段可由專案管理員清理。

離線修改先寫入以工作區分隔的 IndexedDB 快取，恢復網路後重試；若其他裝置已有更新，保留本機版本並提示衝突。使用「載入雲端版本」前會保存本機復原備份，可直接下載。完整備份與還原仍可使用。

原先的 IndexedDB workspace 資料保留，不因首次移轉刪除。不同網站來源的瀏覽器資料無法直接跨來源讀取；請優先在已有原資料的線上網站瀏覽器開啟。

## Firebase 初次設定

1. 啟用 Authentication 的匿名登入供應商。
2. 建立 Firestore Standard 的 (default) 資料庫。
3. 發布根目錄 firestore.rules；不要開放資料庫全面讀寫。
4. public-site/firebase-config.js 僅放公開的 Firebase Web 設定，不使用服務帳戶金鑰。
5. 若使用 Firebase CLI，於此目錄執行 firebase deploy --only firestore --project purchase-studio-95d84，以一併套用片段內容欄位的索引排除。

## GitHub Pages

推送 main 後，.github/workflows/pages.yml 只部署 public-site。GitHub 不保存採購內容。不要把真實原稿、進度表、備份或含資料的 samples.json 提交到 GitHub。公開的 samples.json 保持 files 空陣列。

## 驗證

執行 node tests/cloud-sync.test.mjs，涵蓋移轉、校驗、跨裝置更新、離線重試、版本衝突、編輯保護與未知連結隔離。

需求辨識測試：node tests/intake.test.mjs。Excel 工單依欄位名稱、成品區塊和勾選項目辨識；不使用固定儲存格位置、不以排版 PCS 當作訂購數量，也不把客戶料號代入銘星料號。缺少或矛盾的欄位保留目前值，並在上傳結果中顯示。重複上傳相同文件會更新原需求紀錄，不重複增加附件；保留草稿識別碼與原 Word 表格版型。
