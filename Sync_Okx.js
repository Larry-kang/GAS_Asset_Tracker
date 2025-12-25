/**
 * OKX 餘額同步系統 (v2.2 - 除錯版)
 * 用於邏輯驗證的精簡版本。
 */
function getOkxBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('OKX Balance');
  if (!sheet) {
    ss.toast("錯誤: 找不到工作表 'OKX Balance'。");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OKX_API_KEY');
  const apiSecret = props.getProperty('OKX_API_SECRET');
  const apiPassphrase = props.getProperty('OKX_API_PASSPHRASE');

  if (!apiKey || !apiSecret || !apiPassphrase) {
    ss.toast("錯誤: 未在腳本屬性中設定 OKX API 憑證。");
    return;
  }

  const baseUrl = 'https://www.okx.com';
  const combinedTotals = new Map();
  var logMessages = [];

  try {
    // --- 1. 帳戶餘額 (Account Balance) ---
    ss.toast('正在獲取 OKX 帳戶餘額 (1/2)...');
    const accJson = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
    if (accJson.code === "0" && accJson.data && accJson.data[0] && accJson.data[0].details) {
      logMessages.push("Account: 成功");
      accJson.data[0].details.forEach(function (b) {
        const ccy = b.ccy;
        const totalAmt = (parseFloat(b.availBal) || 0) + (parseFloat(b.frozenBal) || 0);
        if (totalAmt > 0) {
          const currentTotal = combinedTotals.get(ccy) || 0;
          combinedTotals.set(ccy, currentTotal + totalAmt);
        }
      });
    } else { logMessages.push("Account 失敗: " + (accJson.msg || '未知錯誤')); }

    // --- 2. 簡單理財餘額 (Simple Earn) ---
    ss.toast('正在獲取 OKX 理財餘額 (2/2)...');
    const savingsJson = fetchOkxApi_(baseUrl, '/api/v5/finance/savings/balance', {}, apiKey, apiSecret, apiPassphrase);
    if (savingsJson.code === "0" && savingsJson.data) {
      logMessages.push("Earn: 成功");
      savingsJson.data.forEach(function (b) {
        const ccy = b.ccy;
        const totalAmt = parseFloat(b.amt) || 0;
        if (totalAmt > 0) {
          const currentTotal = combinedTotals.get(ccy) || 0;
          combinedTotals.set(ccy, currentTotal + totalAmt);
        }
      });
    } else {
      logMessages.push("Earn 失敗: " + (savingsJson.msg || '未知錯誤'));
    }

    // --- 3. 寫入工作表 ---
    var sheetData = [];
    combinedTotals.forEach(function (total, ccy) {
      if (total > 1e-8) {
        sheetData.push([ccy, total]);
      }
    });
    sheetData.sort((a, b) => b[1] - a[1]);

    sheet.getRange('A2:C').clearContent();
    if (sheetData.length > 0) {
      sheet.getRange(2, 1, sheetData.length, 2).setValues(sheetData);
      sheet.getRange(2, 3).setValue(new Date());
    }

    ss.toast('OKX 餘額同步完成！');

  } catch (e) {
    Logger.log('getOkxBalance 錯誤: ' + e.toString());
    ss.toast('腳本執行錯誤: ' + e.message);
  }
}

/**
 * OKX API 輔助函式
 */
function fetchOkxApi_(baseUrl, endpoint, params, apiKey, apiSecret, apiPassphrase) {
  const method = 'GET';
  const timestamp = new Date().toISOString();

  let queryString = '';
  if (params && Object.keys(params).length > 0) {
    queryString = '?' + Object.keys(params).map(function (key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    }).join('&');
  }

  const url = baseUrl + endpoint + queryString;
  const stringToSign = timestamp + method + endpoint;
  const signature = getOkxSignature_(stringToSign, apiSecret);

  const headers = {
    'OK-ACCESS-KEY': apiKey,
    'OK-ACCESS-SIGN': signature,
    'OK-ACCESS-TIMESTAMP': timestamp,
    'OK-ACCESS-PASSPHRASE': apiPassphrase,
    'Content-Type': 'application/json'
  };
  const options = { 'method': method, 'headers': headers, 'muteHttpExceptions': true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
    return JSON.parse(text);
  } catch (e) {
    return { code: "-1", msg: "連線錯誤: " + e.message };
  }
}

function getOkxSignature_(stringToSign, secret) {
  var signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    stringToSign,
    secret
  );
  return Utilities.base64Encode(signatureBytes);
}