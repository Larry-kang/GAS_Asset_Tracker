# SAP (Sovereign Asset Protocol) v24.6
> **全自動資產追蹤與戰略再平衡系統 - 比特幣本位版**

### 核心目標
本專案旨在實現「比特幣本位 (Bitcoin Standard)」的個人/實體資產自動化管理。透過 Google Apps Script (GAS) 整合多交易所、銀行與冷錢包數據，結合動態戰略決策引擎，確保資產維持在最佳配置狀態。
> **版本狀態**: Stable (Open Source Release)

---

## 系統簡介 (Abstract)

SAP 是一套建構於 Google Apps Script (GAS) 之上的無伺服器 (Serverless) 金融指管系統。它不只是一個記帳工具，而是一個整合了「數據採集」、「風險監控」與「戰略執行」的自動化引擎。系統透過 WebSocket 與 REST API 即時整合全球前四大加密貨幣交易所與台灣證交所數據，為使用者提供上帝視角的資產儀表板。

### 1. 核心層 (Core)
*   `Core_MainMaster.js`: 自動化主流程調度。
*   `Core_StrategicEngine.js`: 戰略決策引擎，處理 BTC 強度策略與再平衡。
*   `Core_LogService.js`: 結構化日誌服務，並提供自動清理機制。

### 2. 同步層 (Sync)
*   `Sync_Binance.js`, `Sync_Okx.js`, `Sync_BitoPro.js`: 各交易所餘額自動同步。
*   `Lib_SyncManager.js`: 統合數據處理與工作表寫入的核心庫。

### 3. 工具與配置層 (Util & Config)
*   `Config.js`: 系統常量與全局閾值。
*   `Util_Settings.js`: [v24.6 NEW] 中心化配置管理，封裝 PropertiesService 調用。
*   `Util_Credentials.js`: API 憑證權限管理中心。
*   `Util_HealthCheck.js`: 系統完整度與連線診斷。
*   `Util_ImportJSON.js`: 多功能 JSON 數據抓取工具。
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

### 🔗 Webhook API (Tunneling)

本系統支援接收來自地端橋接器 (PowerShell) 的 POST 請求，以實現內網穿透與狀態回報。

**Endpoint**: `[您的 Google Web App URL]`

**不支援 GET 請求**，僅接受 POST。所有請求必須包含 `action` 欄位。

#### 1. 更新穿透網址 (Update Tunnel)
用於電腦端將最新的 ngrok/cloudflared 公開網址回報給 GAS。
```json
{
  "action": "update_tunnel_url",
  "url": "https://xxxx-xxxx.ngrok.io",
  "password": "[PROXY_PASSWORD]"
}
```

#### 2. 觸發餘額更新 (Trigger Sync)
強制 GAS 執行一次餘額同步 (預設同步 Binance)。
```json
{
  "action": "trigger_balance_update",
  "password": "[PROXY_PASSWORD]"
}
```

#### 3. 客戶端錯誤回報 (Log Error)
將電腦端發生的錯誤記錄到 GAS 日誌。
```json
{
  "action": "log_client_error",
  "message": "PowerShell Script Crashing...",
  "password": "[PROXY_PASSWORD]"
}
```

---

## 動態資產配置比例設定 (Strategic Allocation)

自 **v24.6.1** 起，系統支援透過試算表動態調整 Layer 1, 2, 3 的配置目標。

### 設定步驟：
1. 開啟您的 Google Sheet。
2. 切換至 `Key Market Indicators` 工作表。
3. 在 `A 欄 (指標鍵名)` 與 `B 欄 (數值)` 加入以下對應項目：
   - `Alloc_L1_Target`: 設定 Layer 1 (數位儲備) 的目標佔比（例如：`0.65` 代表 65%）。
   - `Alloc_L2_Target`: 設定 Layer 2 (信用基底) 的目標佔比（例如：`0.25` 代表 25%）。
   - `Alloc_L3_Target`: 設定 Layer 3 (戰術流動性) 的目標佔比（例如：`0.10` 代表 10%）。

### 注意事項：
- **優先順序**: 試算表中的設定會覆蓋代碼中的 `defaultTarget`。若該鍵值不存在或為空，系統將回退至預設值 (L1: 60%, L2: 30%, L3: 10%)。
- **自動聯動**: 修改比例後，下次執行「戰略狀態報告」或「每日檢查」時，資產佔比計算與資金流向建議將自動套用新比例。

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

