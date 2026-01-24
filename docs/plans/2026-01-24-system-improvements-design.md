# 系統優化設計文件 (System Improvements)

**日期**: 2026-01-24
**狀態**: 草案 (Draft)

## 1. 問題陳述 (Problem Statement)
目前系統有兩個主要改進點：
1.  **日誌閱讀性**：`System_Logs` 目前為舊資料在最上方 (ASC)，閱讀新資料需捲動到底部。且資料保留策略 (2000筆) 對於 Google Sheet 效能與閱讀性來說稍嫌寬鬆。
2.  **合約資產缺失**：使用者反應撈不到 Binance 合約餘額。經確認該 API Key 未開啟 "Futures" 權限，且程式目前無明確提示此權限錯誤。

## 2. 解決方案 (Solutions)

### A. System_Logs 優化
採用 **LIFO (Last In First Out) 堆疊模式**。

*   **排序**：`DESC` (最新的在 Row 2)。
*   **寫入方式**：由 `appendRow` 改為 `insertRowBefore(2)`。
*   **容量限制**：
    *   **Soft Limit**: 500 筆。
    *   **Cleanup Trigger**: 當總列數 > 550 時，一次刪除底部 50 筆。
    *   **理由**: 批量刪除比逐筆刪除更節省 GAS Quota (讀寫次數)。

### B. Binance Permission Discovery
不需修改資產抓取邏輯 (USD1 議題暫置)，而是加強 **權限偵測**。

*   **機制**：在 `getBinanceBalance` 流程中，對 Futures API 回傳的錯誤碼進行攔截。
*   **判定**：若收到 `401` 或 `403` 且 Context 為 Futures，則 Log 輸出明確指示：「請至幣安後台開啟 Futures 權限」。

## 3. 實作細節

### 3.1 Core_LogService.js
```javascript
// Pseudo Code
log(level, msg) {
  sheet.insertRowBefore(2);
  sheet.getRange("A2:D2").setValues([[date, level, context, msg]]);
  
  if (sheet.getLastRow() > 550) {
    sheet.deleteRows(501, sheet.getLastRow() - 500); 
  }
}
```

### 3.2 Sync_Binance.js
```javascript
// Pseudo Code
if (res.code === "-403" && isFuturesEndpoint) {
   SyncManager.log("WARNING", "Permission Denied: Enable Futures API Key", "Sync_Binance");
}
```

## 4. 驗收標準
- [ ] `System_Logs` 最新的一筆永遠在第 2 行。
- [ ] 當 Log 超過 550 筆時，自動縮減回 500 筆。
- [ ] 在未開權限狀況下，System Log 出現明確的「權限不足」警告，而非靜默失敗。
