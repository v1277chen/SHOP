# 鐘錶眼鏡客戶管理系統 (Watch & Glasses Customer Management System)

這是一個基於 **React** (前端) 與 **Google Apps Script** (後端) 開發的客戶管理系統。
專為鐘錶眼鏡行設計，提供客戶資料管理、驗光紀錄查詢、以及系統權限管理功能。

## ✨ 主要功能

*   **客戶管理 (CMF)**: 新增、修改、刪除客戶基本資料。
*   **驗光紀錄 (PRXMF)**: 管理詳細的左右眼驗光數據、隱形眼鏡規格與備註。
*   **伺服器端分頁與搜尋**: 支援大量資料的高效搜尋與分頁顯示 (Server-side Pagination & Search)。
*   **權限管理**: 整合 Firebase Auth 與 Google 登入，支援管理員 (Admin) 與一般使用者 (User) 角色。
*   **響應式設計**: 優化的 UI/UX，支援桌面與行動裝置操作。

## 🚀 部署說明 (Deployment)

本專案已設定 **GitHub Actions** 自動化部署流程。

1.  **啟用 GitHub Pages**:
    - 進入 GitHub Repository 的 **Settings** -> **Pages**。
    - 在 **Build and deployment** 下的 **Source** 選擇 `Deploy from a branch`。
    - 在 **Branch** 選擇 `gh-pages` 分支 (此分支由 Action 自動建立) 並選擇 `/ (root)` 資料夾。
    - 點擊 **Save**。

2.  **自動更新**:
    - 每次推送 (Push) 代碼至 `main` 分支時，GitHub Actions 會自動將最新版本的 `index.html` 部署至 `gh-pages` 分支。
    - 部署完成後，即可透過 GitHub Pages 提供的網址訪問系統。

## 🛠️ 開發說明 (Development)

### 檔案結構
*   `index.html`: 前端主程式 (React + Tailwind CSS)，包含所有 UI 與邏輯。
*   `code.gs`: 後端程式 (Google Apps Script)，負責處理 Google Sheets 資料讀寫與搜尋邏輯。

### 後端設定 (Google Apps Script)
1.  建立一個新的 Google Sheet，並設定 `CMF` (客戶資料) 與 `PRXMF` (驗光資料) 兩個工作表。
2.  開啟 **擴充功能** -> **Apps Script**，貼上 `code.gs` 的內容。
3.  修改 `SPREADSHEET_ID` 為您的試算表 ID。
4.  執行 **部署** -> **新增部署** -> 類型選擇 **網頁應用程式**。
    - 執行身分: **我 (Me)**
    - 誰可以存取: **所有人 (Anyone)** (注意：這是為了讓前端能存取 API，權限控管由前端 Firebase Auth 加強)
5.  複製生成的 **網頁應用程式網址**。

### 前端設定 (index.html)
1.  開啟 `index.html`。
2.  找到 `const GAS_API_URL`，將其替換為即將部署的 GAS 網頁應用程式網址。
3.  設定 `firebaseConfig`，填入您的 Firebase 專案設定。
4.  (可選) 調整 `DEFAULT_PAGE_SIZE` 設定每頁顯示筆數。

## 📝 搜尋與分頁邏輯
*   系統採用 **伺服器端搜尋 (Server-side Search)**。
*   當輸入搜尋條件並點擊「搜尋」時，前端會發送條件至 GAS 後端。
*   後端負責過濾資料並進行分頁 (Pagination)，僅回傳當前頁面的資料，大幅提升大數據量下的效能。

## 📜 授權
MIT License
