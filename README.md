# 採購工作室

供應商專屬採購單網站。選擇供應商、填寫三碼單號，上傳 Word／Excel／PDF／圖片需求後直接套單並下載 Word。

單號格式為 `MH26A-10-035`：年月依 Asia/Taipei 自動更新，三碼尾碼自行輸入。介面使用深綠配色，沒有額外審核步驟。

## GitHub Pages

網站檔案位於 `public-site/`。GitHub 儲存庫的 Settings → Pages → Source 選擇 GitHub Actions。推送 main 後，`.github/workflows/pages.yml` 會部署網站；不需 npm 安裝或建置。

## 初次使用

公開網站不內建任何真實採購原稿、公司地址、報價或聯絡人。請在「供應商與模板」新增自己的 Word 模板，或於「開單資料與備份」還原原本瀏覽器匯出的備份。

所有模板、需求、草稿與紀錄都保存在使用中的瀏覽器。GitHub Pages 只提供程式檔案，不保存或同步你的採購資料；清除網站資料前請先備份。從本機網址移至 GitHub 網址時，也需先匯出再還原備份。

圖片 OCR 首次使用需連線下載辨識引擎與語言資料；文件內容不會上傳到 GitHub。讀不到的欄位留白，不自動猜測。

更新程式時修改 `public-site/` 的對應檔案。不要將真實原稿、含採購資料的 samples.json 或完整備份加入儲存庫。原始 Word 版型用於 Word 匯出；網頁預覽和 PDF 列印版採內容排版。
