// =======================================================
// --- BitoPro 交易所餘額查詢系統 (v2.0 - Standardized) ---
// --- Refactored to use Lib_SyncManager
// =======================================================

function getBitoProBalance() {
  const MODULE_NAME = "Sync_BitoPro";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('BITOPRO_API_KEY');
    const apiSecret = props.getProperty('BITOPRO_API_SECRET');
    const baseUrl = 'https://api.bitopro.com/v3';

    if (!apiKey || !apiSecret) {
      SyncManager.log("ERROR", "錯誤：BitoPro API 金鑰或 Secret 未設定。", MODULE_NAME);
      return;
    }

    SyncManager.log("INFO", "正在獲取 BitoPro 帳戶餘額...", MODULE_NAME);

    // --- 1. 獲取帳戶餘額 ---
    const endpoint = '/accounts/balance';
    const json = fetchBitoProApi_(baseUrl, endpoint, 'GET', {}, apiKey, apiSecret);

    if (json && json.data) {
      const assetsMap = new Map();
      let count = 0;

      json.data.forEach(b => {
        // 官方欄位: amount (總額), available (可用), stake (圈存)
        // 標準化只取 amount (總額)
        const total = parseFloat(b.amount) || 0;

        if (total > 0) {
          assetsMap.set(b.currency.toUpperCase(), total);
          count++;
        }
      });

      SyncManager.log("INFO", `BitoPro API 回傳 ${count} 個有效資產。`, MODULE_NAME);

      // --- 2. 寫入標準化工作表 ---
      SyncManager.writeToSheet(ss, 'BitoPro Balance', ['Currency', 'Total', 'Last Updated'], assetsMap);
      SyncManager.log("INFO", "BitoPro Sync Complete", MODULE_NAME);

    } else {
      SyncManager.log("ERROR", `API FAILED: ${json.error || 'Unknown Error'}`, MODULE_NAME);
    }
  });
}

/**
 * [Helper] BitoPro API 呼叫專用函式
 * 修正時間誤差問題 (Time Skew Fix)
 */
function fetchBitoProApi_(baseUrl, endpoint, method, params, apiKey, apiSecret) {
  // ⭐ 修正點：將時間往回推 2000 毫秒 (2秒) 以容許網路延遲與時間誤差
  const nonce = Date.now() - 2000;

  const payloadObj = { ...params, nonce: nonce };
  const payloadJson = JSON.stringify(payloadObj);
  const payloadBase64 = Utilities.base64Encode(payloadJson);

  const signature = getBitoProSignature_(payloadBase64, apiSecret);

  const headers = {
    'X-BITOPRO-APIKEY': apiKey,
    'X-BITOPRO-PAYLOAD': payloadBase64,
    'X-BITOPRO-SIGNATURE': signature,
    'Content-Type': 'application/json'
  };

  const options = {
    'method': method,
    'headers': headers,
    'muteHttpExceptions': true
  };

  if (method === 'POST') {
    options.payload = payloadJson;
  }

  const url = baseUrl + endpoint;

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      console.warn(`[BitoProApi] Error Code: ${responseCode}, Body: ${responseText}`);
    }

    return JSON.parse(responseText);

  } catch (e) {
    console.error(`fetchBitoProApi_ FAILED: ${e.toString()}`);
    return { error: e.message };
  }
}

/**
 * [Crypto] 產生 BitoPro 簽章
 * 使用 HMAC-SHA384 並轉為 Hex 字串
 */
function getBitoProSignature_(payloadBase64, apiSecret) {
  const signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_384,
    payloadBase64,
    apiSecret
  );

  return signatureBytes.map(function (byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}