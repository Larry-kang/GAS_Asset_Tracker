# SAP 主權資產協議 (Sovereign Asset Protocol) v24.6

> **系統代號**: SAP  
> **核心目標**: 實現比特幣本位 (Bitcoin Standard) 的全自動化資產追蹤與戰略再平衡系統。  
> **版本狀態**: Stable (Open Source Release)

---

## 系統簡介 (Abstract)

SAP 是一套建構於 Google Apps Script (GAS) 之上的無伺服器 (Serverless) 金融指管系統。它不只是一個記帳工具，而是一個整合了「數據採集」、「風險監控」與「戰略執行」的自動化引擎。系統透過 WebSocket 與 REST API 即時整合全球前四大加密貨幣交易所與台灣證交所數據，為使用者提供上帝視角的資產儀表板。

## 核心功能 (Core Features)

### 1. Unified Asset Ledger (全域資產總帳)
*   **單一事實來源**: 導入 `Lib_SyncManager`，將 Binance, OKX, Pionex, BitoPro 數據標準化為統一格式 (`Exchange`, `Type`, `Currency`, `Amount`, `Status`)。
*   **原子化更新**: 採 Exchange-Level Atomic Replacement 策略，確保數據一致性。
*   **台股串接**: 自動抓取 TWSE 融資維持率與庫存市值。
*   **匯率換算**: 透過 Google Finance 與自建 API，即時計算 TWD/USD 總資產淨值 (NAV)。

### 2. 戰略決策引擎 (Strategic Engine)
*   **風險監控**: 即時計算多平台質押維持率 (Maintenance Ratio)，低於警戒線 (如 2.1) 自動發送紅燈警報。
*   **馬丁格爾策略**: 根據 BTC 回調幅度 (-30% ~ -70%)，自動計算建議抄底金額。
*   **資產再平衡**: 監控三大資產層 (Layer 1 儲備 / Layer 2 信用 / Layer 3 彈藥) 偏移率，提供再平衡建議。

### 3. devOps 與自動化
*   **CI/CD Pipeline**: 整合 GitHub Actions，實現 `git push` 即自動部署至 GAS。
*   **健康檢查**: 內建 `Util_HealthCheck` 模組，每日巡檢 API 連線與憑證狀態。
*   **日誌服務**: 自建 `Core_LogService`，將關鍵操作與錯誤結構化寫入獨立日誌分頁，便於追蹤。

---

## CI/CD 自動化部署流程

本專案採用 **GitHub Actions** 結合 **Google Clasp** 進行持續整合與部署。

### `setup_cicd` 是什麼？
這是一個專為 Windows 環境設計的 PowerShell 腳本 (`setup_cicd.ps1`)。
*   **功能**: 自動讀取您本地的 Google 登入憑證 (`~/.clasprc.json`)，將其轉換為 Base64 編碼，並透過 GitHub CLI (`gh`) 安全地上傳至您 GitHub Repo 的 Secrets (`CLASPRC_JSON_BASE64`)。
*   **目的**: 讓 GitHub Actions 雲端伺服器擁有「扮演您」的權限，能夠執行 `clasp push` 將程式碼寫入您的 Google Sheet 專案。

### 如何啟用？
1.  確認已安裝 `clasp` 並登入 (`clasp login`)。
2.  在專案根目錄執行 PowerShell 指令：
    ```powershell
    .\setup_cicd.ps1
    ```
3.  看到綠色成功訊息後，未來只需執行 `git push`，程式碼便會自動同步至 Google Apps Script。

---

## 專案結構 (Architecture)

```
/
├── Core_*.js           # [核心層] 戰略引擎、日誌服務、主控路徑
├── Lib_*.js            # [函式庫] 通用介面 (SyncManager, API Wrappers)
├── Event_*.js          # [事件層] 觸發器 (OnOpen, Webhook)
├── Sync_*.js           # [同步層] 各大交易所 API 實作 (Binance, OKX...)
├── Util_*.js           # [工具層] 通用工具 (HealthCheck, TWSE...)
├── .github/workflows/  # [CI/CD] GitHub Actions 設定檔
└── setup_cicd.ps1      # [部署] 憑證自動化設定腳本
```

---

## 快速開始 (Quick Start)

1.  **環境準備**: 安裝 Node.js, Clasp, Git。
2.  **拉取專案**: `git clone <REPO_URL>`
3.  **登入授權**: `clasp login`
4.  **設定 CI**: 執行 `.\setup_cicd.ps1`
5.  **開始開發**: 修改代碼後 `git push` 即可。

---

## Author

**Larry Kang**

* **Role:** Senior Backend Engineer | FinTech Specialist
* **Focus:** Distributed Systems, Payment Architecture, .NET Performance Tuning.
* **Contact:** [LinkedIn Profile](www.linkedin.com/in/larry-kang)
