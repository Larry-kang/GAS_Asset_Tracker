# SAP (Sovereign Asset Protocol)
> 全自動資產追蹤與戰略再平衡系統

SAP 是一套建構於 Google Apps Script (GAS) 的資產管理與監控系統。它整合多交易所與市場資料來源，將價格更新、資產同步、風險監控、策略分析、通知與每日快照串成一條自動化流程。

## 系統重點

- 自動同步多交易所資產，並整合到統一資產台帳。
- 追蹤質押、借貸與理財部位，維持完整的資產與負債視角。
- 根據 `Balance Sheet` 與市場指標產出戰略監控與再平衡建議。
- 透過 Google Sheets UI、排程與通知機制提供日常操作入口。

## 主要模組

- `Core_MainMaster.js`: 主流程調度與每日例行工作。
- `Core_StrategicEngine.js`: 戰略決策、風險監控與報告邏輯。
- `Lib_SyncManager.js`: 統一資產台帳寫入與同步整合。
- `Sync_*.js`: 各交易所同步實作，目前包含 Binance、OKX、BitoPro、Bitget。
- `Util_*.js`: 設定、憑證、快取、健康檢查與通知等共用工具。
- `Event_*.js`: 試算表選單與 Web App 事件入口。

## 專案結構

```text
/
|-- Core_*.js
|-- Event_*.js
|-- Lib_*.js
|-- Sync_*.js
|-- Util_*.js
|-- Config.js
|-- appsscript.json
`-- .github/workflows/
```

## 公開版說明

此 repo 僅保留產品程式與必要的部署工作流。環境專屬的部署設定、橋接器設定、本地維運腳本與內部文件不包含在公開版中。

## Author

**Larry Kang**

