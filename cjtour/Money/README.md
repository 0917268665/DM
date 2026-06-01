# 💼 旅行社獨立儀表板系統 (Cloudflare Pages + Google Sheets API)

這是一套為旅行社開發的企業級無伺服器 (Serverless) 帳務與出團管理系統。
* **獨立前端**：採用 HTML5/CSS3 響應式雙欄儀表板，設計精美（Glassmorphism 玻璃擬態風格），可直接託管在 **Cloudflare Pages** 獲得您專屬的公網網址，方便分享給業務或會計人員。
* **後端資料庫**：基於 **Google 試算表 (Google Sheets)** 與 **Google Apps Script (GAS) Web App API**。
* **財務引擎**：整合了領隊姓名、領隊代墊、業務佣金、退稅等完整欄位，並支援一鍵生成 PDF 報表下載與網頁端互動出團月曆。

---

## 📂 專案檔案結構
1. **[Code.js](file:///Users/chenyanbai/Documents/GitHub/DM/cjtour/Money/Code.js)**：GAS 後端 API 程式，提供 `doPost(e)` HTTP API 接口與報表生成邏輯。
2. **[index.html](file:///Users/chenyanbai/Documents/GitHub/DM/cjtour/Money/Index.html)**：前端網頁（可單獨部署於 Cloudflare 或是直接在本機雙擊開啟），支援 API 網址緩存設定。
3. **[README.md](file:///Users/chenyanbai/Documents/GitHub/DM/cjtour/Money/README.md)**：本部署手冊。

---

## 🛠️ 部署指南

本系統的部署分為 **「Google 試算表 API」** 與 **「Cloudflare Pages 網頁」** 兩部分：

### 第一部分：部署 Google Apps Script API 後端
1. **建立並命名試算表**：在 Google Sheets 中建立空白試算表。
2. **開啟編輯器**：點選上方選單 **「擴充功能」** ➔ **「Apps Script」**。
3. **複製代碼**：清空預設的 `Code.gs`，將本專案的 **[Code.js](file:///Users/chenyanbai/Documents/GitHub/DM/cjtour/Money/Code.js)** 程式碼全部貼上，點選儲存 💾。
4. **初始化工作表**：
   - 回到試算表網頁，重新整理。
   - 點選上方選單 **「💼 旅行社帳務系統」** ➔ **「初始化系統工作表」**（第一次執行需按提示點擊「允許授權」）。
   - 初始化成功後，試算表會自動建置所有工作表與範例資料。
5. **部署為 Web 應用程式 (Web App)**：
   - 在 Apps Script 編輯器右上角，點選 **「部署 (Deploy)」** ➔ **「新增部署 (New deployment)」**。
   - 點選左上角齒輪圖示，選擇 **「Web 應用程式 (Web app)」**。
   - **專案說明**：輸入「旅行社帳務 API」。
   - **執行身分**：選擇 **「您的帳戶 (Me)」**。
   - **誰有存取權**：選擇 **「任何人 (Anyone)」** *(⚠️ 請務必選擇 Anyone，否則 Cloudflare 網頁將無法存取)*。
   - 點選 **「部署」**。
   - 部署完成後，請**複製畫面上產生的「Web 應用程式網址」 (Web App URL)**，這就是您的後端 API 網址（例如：`https://script.google.com/macros/s/XXXXXX/exec`）。

---

### 第二部分：部署前端網頁至 Cloudflare Pages (完全免費)
1. **登入 Cloudflare**：進入 [Cloudflare 控制台](https://dash.cloudflare.com/)（若無帳號請免費註冊一個）。
2. **進入 Pages 設定**：點選左側選單的 **「Workers 和 Pages (Workers & Pages)」** ➔ 點選 **「Pages」** ➔ **「上傳資產 (Upload assets)」**。
3. **建立專案**：
   - 輸入您的專案名稱（例如 `cjtour-dashboard`）。
   - 點選 **「建立專案 (Create project)」**。
4. **上傳網頁檔案**：
   - 在本機新建一個資料夾，命名為 `public`（或任何您喜歡的名字）。
   - 將本專案中的 **[index.html](file:///Users/chenyanbai/Documents/GitHub/DM/cjtour/Money/Index.html)** 檔案重新命名為小寫的 **`index.html`**，並放到該資料夾中。
   - 將該資料夾直接拖曳上傳至 Cloudflare 網頁中。
   - 上傳成功後，點選 **「部署網站 (Deploy site)」**。
5. **取得專屬網址**：部署完成後，Cloudflare 會為您產生一個永久且安全的 HTTPS 網址（例如：`https://cjtour-dashboard.pages.dev`）。

---

## 🚀 連線與日常使用

1. **開啟您的 Cloudflare 網頁**（或在您電腦中直接按兩下打開 `index.html` 網頁）。
2. **設定 API 網址**：
   - 在網頁最頂部的 **「API URL：」** 輸入框中，貼上您在 **第一部分第 5 步** 複製的 Apps Script Web App 網址。
   - 貼上後按下 Enter 或點擊輸入框外，系統會自動在瀏覽器儲存此設定（`localStorage`），未來開啟網頁不需重新輸入。
3. **功能操作**：
   * **出團日曆**：連線成功後，日曆會自動顯示 2026 年 6 月的出團資訊。點選任何有「🚌 1 團」標記的日期，日曆下方會立即展開出團明細。
   * **快捷填寫訂單**：填入業務、領隊、客戶與金額等欄位，點擊「寫入試算表並計算利潤」。新增成功後，系統會自動重新整理日曆。
   * **下載 PDF 報表**：
     - 在「業務業績報表 (PDF)」選擇業務與時間，點擊產生。
     - 在「公司財務年報 (PDF)」點擊產生。
     - 系統會透過 Apps Script 自動在您的雲端硬碟建立 `旅行社帳務報表` 資料夾儲存 PDF，並**自動在新分頁彈出該 PDF** 供您下載或列印。
