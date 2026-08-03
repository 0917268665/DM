# 祥麟紙器 考勤打卡管理系統

一個高顏值、雙角色驗證、打卡地點/GPS網路限制、可自動五合一雙向同步 Google 試算表的企業級考勤打卡管理系統。

---

## 🚀 部署至 GitHub Pages 指南 (方案 B)

只需將本專案 Push 推送至您的 GitHub Repository，**系統將透過 GitHub Actions 自動免費發佈**，全公司員工皆可用手機/電腦開啟存取！

### 步驟 1：在 GitHub 建立 Repository
1. 開啟 [GitHub.com](https://github.com/new) 並建立一個新的公開 (Public) 或私有 (Private) Repository (例如 `HR-Attendance`)。

### 步驟 2：推送到 GitHub
在專案根目錄終端機執行：

```bash
git init
git add .
git commit -m "Deploy SmartHR Attendance System"
git branch -M main
git remote add origin https://github.com/<您的GitHub帳號>/HR-Attendance.git
git push -u origin main
```

### 步驟 3：開啟 GitHub Pages
1. 在 GitHub Repository 頁面點選【Settings】➔ 左側【Pages】。
2. **Build and deployment** 下方的 **Source** 選擇 **`GitHub Actions`**。
3. 稍等約 1 分鐘，即可獲得公開發佈網址：
   `https://<您的帳號>.github.io/HR-Attendance/`

---

## 📍 打卡地點與 GPS 圍欄限制功能

HR 主管可登入後台開啟地點打卡限制：

1. 主管登入 (`admin` / `admin888`) ➔ 點選 **【📊 主管後台】**。
2. 找到 **【📍 打卡地點與 WiFi 限制設定】**：
   - 開啟 **打卡限制開關**。
   - 設定公司中心的 **GPS 緯度 (Lat)**、**經度 (Lng)** 與 **允許半徑 (公尺)** (例如 `200` 公尺)。
3. 開啟後，員工打卡時瀏覽器會自動驗證 GPS 位置，**若超過公司範圍或未開啟定位授權，系統將自動阻擋打卡**並提示距離！

---

## 📊 Google 試算表五合一自動同步分頁

每次打卡與審核時，系統會自動在您的 Google 試算表中建立並更新 5 個工作表：
1. **`考勤紀錄`**：逐筆打卡流水明細
2. **`全公司每月出勤`**：自動維護各月份出勤天數、遲到次數與累計工時
3. **`全公司年度出勤`**：自動維護全公司年度出勤統計
4. **`個人單月出勤明細`**：**以員工個人與月份為單位的單月出勤明細帳冊 (紫底黑字標頭)**
5. **`員工名冊與假額度`**：雙向同步員工資料與特休天數

---

## 🔑 預設登入帳號密碼說明

* **員工登入**：`emp01` / 密碼：`123456` (張小明)
* **主管登入**：`admin` / 密碼：`admin888` (李梅主管 - 具備設定權限與全公司報表)
