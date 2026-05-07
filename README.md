精臣 NIIMBOT 官方商城 🏷️



這是一個輕量化的無伺服器 (Serverless) 購物車系統，專為「精臣標籤機與耗材」的線上選購所設計。

前端採用 React (CDN) 與 Tailwind CSS 開發，後端結合 Google Apps Script 與 Google 試算表，讓商家能輕鬆管理商品與接收訂單，並且完全免費託管。



🚀 系統特色



無伺服器架構 (Serverless)：不需租用主機，靜態網頁可直接免費部署於 GitHub Pages。



Google 試算表管理：無須複雜資料庫，商品上架與訂單接收皆在 Google Sheets 完成。



全方位響應式設計 (RWD)：完美適配手機、平板與桌上型電腦，給顧客最佳的購物體驗。



即時連動購物車：商品規格切換時，名稱、庫存與價格 100% 即時跳轉。



毛玻璃視覺設計：現代化 UI 設計，帶有流暢的過場動畫與互動回饋。



📂 專案檔案結構



index.html: 網站前端主程式（包含 UI 介面、React 邏輯與購物車系統）。



config.js: 網站環境設定檔（用來設定後端 API 網址、Logo 及首頁橫幅大圖）。



backend.gs: Google Apps Script 後端程式碼（僅供備份，需實際部署於 Google 試算表的 Apps Script 中）。



README.md: 專案說明文件。



⚙️ 如何修改與維護？



所有的機密網址與視覺圖片，都已經抽離到 config.js 中。未來若要舉辦活動換圖，或是重新部署後端，只需修改 config.js 即可，無需更動 index.html。



打開 config.js，您可以輕鬆修改以下數值：



GOOGLE\_SCRIPT\_URL: 您的 Google Apps Script 接收訂單網址。



STORE\_NAME: 您的商店名稱。



LOGO\_URL: 左上角品牌 Logo 網址。



BANNER\_PRINTER \& BANNER\_STICKER: 首頁雙橫幅的圖片網址。



📝 關於後端 (backend.gs)



若需要重新設定後端：



建立一個新的 Google 試算表。



點擊 擴充功能 > Apps Script。



將本專案中的 backend.gs 程式碼複製並貼上，覆蓋掉原有的程式碼。



儲存並選擇 setupDropdowns 執行一次，以自動建立商品清單與下拉選單。



發佈為「網頁應用程式」，將權限設為「所有人」，並將獲得的 URL 填入 config.js 中。



🛠️ 開發技術棧



前端: HTML5, React 18, Tailwind CSS, Lucide Icons



後端 / 資料庫: Google Apps Script (GAS), Google Sheets



Developed with modern UI/UX patterns for NIIMBOT e-commerce.

