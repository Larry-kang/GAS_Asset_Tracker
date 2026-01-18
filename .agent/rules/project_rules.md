---
trigger: always_on
description: GAS_Asset_Tracker (SAP) 專案開發規範
---

# GAS_Asset_Tracker 專案規則

> **規則繼承**: 自動繼承全域 `gas-ops` 技能規範。

---

## 1. 技術棧 (Tech Stack)

| 項目 | 規範 |
|------|------|
| 主要語言 | Google Apps Script (ES6+) |
| 版本來源 | `Config.VERSION`（唯一真相來源） |
| 部署方式 | GitHub Actions + clasp |

---

## 2. 檔案載入順序 (Load Order)

GAS 檔案以**名稱字母順序**載入。本專案依賴以下載入順序：

```
1. Config.js              # 全域常數（必須最先載入）
2. Core_LogService.js     # 日誌服務
3. Util_Cache.js          # 快取系統
4. Util_Settings.js       # 參照 ScriptCache
5. Util_Credentials.js    # 參照 Settings
6. Lib_SyncManager.js     # 參照 LogService
7. Sync_*.js              # 同步模組
8. Core_*.js (其他)       # 戰略引擎
9. Event_*.js             # 事件處理
```

> [!IMPORTANT]
> 若新增檔案，確保命名符合載入順序依賴。

---

## 3. API 金鑰命名規範 (Credential Naming)

所有 API 金鑰存於 `Script Properties`，命名格式：

| 金鑰 | 格式 | 範例 |
|------|------|------|
| API Key | `{EXCHANGE}_API_KEY` | `BINANCE_API_KEY` |
| Secret | `{EXCHANGE}_API_SECRET` | `OKX_API_SECRET` |
| Passphrase | `{EXCHANGE}_API_PASSPHRASE` | `OKX_API_PASSPHRASE` |
| 系統設定 | `UPPER_SNAKE_CASE` | `ADMIN_EMAIL`, `DISCORD_WEBHOOK_URL` |

---

## 4. 程式碼風格 (Code Style)

### 4.1 函式命名

| 類型 | 慣例 | 範例 |
|------|------|------|
| 公開函式 | camelCase | `runDailyInvestmentCheck()` |
| 私有函式 | camelCase + `_` 後綴 | `syncAllAssets_()` |
| 常數物件 | PascalCase | `Config`, `LogService`, `SyncManager` |

### 4.2 模組結構

每個 `Sync_*.js` 須遵循：

```javascript
function get{Exchange}Balance() {
  SyncManager.run("Sync_{Exchange}", () => {
    // 1. 取得憑證
    // 2. 驗證憑證
    // 3. 資料收集
    // 4. 呼叫 SyncManager.updateUnifiedLedger()
  });
}
```

---

## 5. 測試流程 (Testing)

### 5.1 驗證腳本

執行 `Test_Verification.js` 中的 `verifyOptimizations()` 驗證模組載入。

### 5.2 手動驗證

1. 開啟 Google Sheet
2. 選單 → 「系統維運管理」→「核心引導設定」
3. 確認觸發器建立成功

---

## 6. 部署流程 (Deployment)

```bash
# 標準流程 (推薦)
git add . && git commit -m "feat: 描述" && git push

# GitHub Actions 自動執行 clasp push
```

> [!CAUTION]
> 禁止直接在 GAS 網頁編輯器修改程式碼，除非緊急修復後立即 `clasp pull`。
