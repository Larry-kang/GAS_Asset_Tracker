// =======================================================
// --- Pionex 餘額查詢系統 (v2.2 - Standardized) ---
// --- Refactored to use Lib_SyncManager
// =======================================================

function getPionexBalance() {
  const MODULE_NAME = "Sync_Pionex";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('PIONEX_API_KEY');
    const apiSecret = props.getProperty('PIONEX_API_SECRET');
    const baseUrl = 'https://api.pionex.com';
    const endpoint = '/api/v1/account/balances';

    if (!apiKey || !apiSecret) {
      SyncManager.log("ERROR", "錯誤：未設定 PIONEX_API_KEY / SECRET", MODULE_NAME);
      return;
    }

    SyncManager.log("INFO", "正在獲取 Pionex 數據...", MODULE_NAME);

    // 呼叫 API
    const json = fetchPionexApi_(baseUrl, endpoint, {}, apiKey, apiSecret);

    if (json.result === true && json.data && json.data.balances) {
      const balances = json.data.balances;
      SyncManager.log("INFO", `API 回傳了 ${balances.length} 個幣種資料`, MODULE_NAME);

      const assetsMap = new Map();
      let debugRawData = [];

      balances.forEach(b => {
        const ccy = b.coin;
        const free = parseFloat(b.free) || 0;
        const frozen = parseFloat(b.frozen) || 0;
        const total = free + frozen;


        if (total > 0) {
          assetsMap.set(ccy, total);
          // 收集原始數據用於 Log (替代原本寫在 Sheet E 欄的做法，保持 UI 乾淨)
          debugRawData.push(`${ccy}: ${JSON.stringify(b)}`);
        }
      });

      // 寫入詳細 Log (給需要 Debug Raw Data 的人看)
      if (debugRawData.length > 0) {
        console.log("--- Pionex Raw Data Dump ---");
        console.log(debugRawData.join("\n"));
      }

      // 標準化寫入 Sheet
      SyncManager.writeToSheet(ss, 'Pionex Balance', ['Currency', 'Total', 'Last Updated'], assetsMap);
      SyncManager.log("INFO", "Pionex Sync Complete", MODULE_NAME);

    } else {
      SyncManager.log("ERROR", `API Error: ${JSON.stringify(json)}`, MODULE_NAME);
    }
  });
}

// =======================================================
// --- 輔助函式 (Helper Functions) ---
// =======================================================

function fetchPionexApi_(baseUrl, endpoint, params, apiKey, apiSecret) {
  const method = 'GET';
  const timestamp = new Date().getTime();

  const signatureParams = { ...params };
  signatureParams.timestamp = timestamp;

  const sortedKeys = Object.keys(signatureParams).sort();
  const queryString = sortedKeys.map(key => {
    return key + '=' + signatureParams[key];
  }).join('&');

  const fullPath = endpoint + '?' + queryString;
  const url = baseUrl + fullPath;

  const body = '';
  const stringToSign = method + fullPath + body;
  const signature = getPionexSignature_(stringToSign, apiSecret);

  const headers = {
    'PIONEX-KEY': apiKey,
    'PIONEX-SIGNATURE': signature,
    'Content-Type': 'application/json'
  };

  const options = {
    'method': method,
    'headers': headers,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    if (responseText && responseText.trim().startsWith('{')) {
      return JSON.parse(responseText);
    } else {
      return { result: false, message: "Non-JSON response: " + responseText };
    }
  } catch (e) {
    return { result: false, message: "Fetch error: " + e.message };
  }
}

function getPionexSignature_(stringToSign, secret) {
  const signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    stringToSign,
    secret
  );
  return signatureBytes.map(function (byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}