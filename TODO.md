# TODO - P2 級別優化項目

## 低優先級改進清單

以下為可選的程式碼品質改進項目，可視專案需求於未來執行。

---

### ? P2-1: Legacy 支援檔案標記
**檔案**：Util_LegacySupport.js  
**狀態**：活躍使用中（被 Sync_Assets.js 引用）  
**建議**：確認是否計劃廢除，若否則移除 "Legacy" 命名以避免誤解

**優先度**：低  
**預估工作量**：0.5 小時

---

### ? P2-2: 測試檔案註解不足
**檔案**：Test_SAP_Simulation.js  
**建議**：加上執行方式與預期輸出的說明

**範例**：
```javascript
/**
 * HOW TO RUN:
 * 1. Open Apps Script editor
 * 2. Select debugSAPLogic function
 * 3. Click Run button
 * 
 * EXPECTED OUTPUT:
 * Console logs showing scenario test results
 */
```

**優先度**：低  
**預估工作量**：0.5 小時

---

### ? P2-3: API 憑證管理改進
**位置**：所有 Sync_*.js  
**建議**：統一從 PropertiesService 讀取，避免重複代碼

**當前狀況**：
每個 Sync 檔案都有類似的憑證讀取邏輯：
```javascript
const apiKey = PropertiesService.getScriptProperties().getProperty('BINANCE_API_KEY');
const apiSecret = PropertiesService.getScriptProperties().getProperty('BINANCE_API_SECRET');
```

**建議改進**：
建立統一的 `Util_Credentials.js`：
```javascript
function getExchangeCredentials(exchange) {
  const props = PropertiesService.getScriptProperties();
  return {
    apiKey: props.getProperty(`${exchange}_API_KEY`),
    apiSecret: props.getProperty(`${exchange}_API_SECRET`),
    // ...
  };
}
```

**優先度**：中低  
**預估工作量**：2 小時

---

### ? P2-4: 建立版本號常量
**建議**：在 Core_StrategicEngine.js 或新建 Config.js 定義版本號

**範例**：
```javascript
const SAP_VERSION = "24.6";
const SYSTEM_NAME = "SAP - Sovereign Asset Protocol";
```

取代散落在各處的硬編碼 "v24.5" 字串

**優先度**：低  
**預估工作量**：0.5 小時

---

### ? P2-5: Webhook 端點文件化
**檔案**：Event_Webhook.js  
**建議**：在 README 加入 Webhook API 規格說明

**應包含內容**：
- Webhook URL 格式
- 支援的事件類型
- Request/Response 範例
- 錯誤處理說明

**優先度**：中低  
**預估工作量**：1 小時

---

### ? P2-6: IMPORTJSON 標記為工具函數
**檔案**：Util_ImportJSON.js  
**建議**：加上 `@customfunction` JSDoc 標籤

**範例**：
```javascript
/**
 * Imports JSON data from a URL into spreadsheet.
 * @param {string} url The URL to fetch JSON from
 * @param {string} query JSONPath query
 * @return {Array} 2D array of values
 * @customfunction
 */
function IMPORTJSON(url, query) { ... }
```

**優先度**：低  
**預估工作量**：0.5 小時

---

## 總計

**項目數量**：6 個  
**總預估工作量**：5 小時  
**建議執行時機**：視專案需求，可在未來迭代中逐步完成

---

## 優先順序建議

若要執行 P2 項目，建議順序：
1. **P2-3**: API 憑證管理（最有實際價值）
2. **P2-5**: Webhook 文件化（提升專案完整度）
3. **P2-4**: 版本號常量（易於維護）
4. **P2-6**: IMPORTJSON 標記（改善工具完整性）
5. **P2-2**: 測試檔案文件（開發體驗）
6. **P2-1**: Legacy 命名檢視（最後處理）
