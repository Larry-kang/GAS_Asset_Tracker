// =======================================================
// --- Pionex 餘額查詢系統 (v2.1 - 除錯版) ---
// --- 專門用於抓取「消失的屯幣寶」
// --- 增加詳細日誌，不過濾任何幣種
// =======================================================

function getPionexBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Pionex Balance');
  if (!sheet) {
    sheet = ss.insertSheet('Pionex Balance');
    sheet.appendRow(['幣種', '總額', 'Free', 'Frozen', 'Raw Data']);
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('PIONEX_API_KEY');
  const apiSecret = props.getProperty('PIONEX_API_SECRET');

  if (!apiKey || !apiSecret) {
    ss.toast("錯誤：未設定 API Key/Secret");
    return;
  }

  const baseUrl = 'https://api.pionex.com';
  const endpoint = '/api/v1/account/balances';

  try {
    ss.toast('正在獲取 Pionex 完整數據 (Debug)...');

    // 呼叫 API
    const json = fetchPionexApi_(baseUrl, endpoint, {}, apiKey, apiSecret);

    Logger.log("--- API 原始回應 (前 1000 字) ---");
    Logger.log(JSON.stringify(json).substring(0, 1000));

    if (json.result === true && json.data && json.data.balances) {

      const sheetData = [];
      const balances = json.data.balances;

      Logger.log(`API 回傳了 ${balances.length} 個幣種資料`);

      balances.forEach(function (b) {
        const ccy = b.coin;
        const free = parseFloat(b.free) || 0;
        const frozen = parseFloat(b.frozen) || 0;
        const total = free + frozen;

        // 修改：完全不過濾，只要 total > 0 就顯示，並記錄原始數據
        if (total > 0) {
          sheetData.push([
            ccy,
            total,
            free,
            frozen,
            JSON.stringify(b) // 把該幣種的原始 JSON 也寫進去，看看有沒有特殊欄位
          ]);
        }
      });

      // 排序
      sheetData.sort((a, b) => b[1] - a[1]);

      // 寫入
      sheet.getRange('A2:E').clearContent();
      if (sheetData.length > 0) {
        sheet.getRange(2, 1, sheetData.length, 5).setValues(sheetData);
        sheet.getRange("F2").setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"));
      }

      ss.toast(`更新完成！請檢查 'Pionex Balance' 的 E 欄 (Raw Data)`);

    } else {
      Logger.log("API Error: " + JSON.stringify(json));
      ss.toast("API 回傳錯誤，請看日誌");
    }

  } catch (e) {
    Logger.log('Error: ' + e.toString());
    ss.toast('執行錯誤: ' + e.message);
  }
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