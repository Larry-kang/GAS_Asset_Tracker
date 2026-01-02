// =======================================================
// --- BitoPro 交易所餘額查詢系統 (v1.0) ---
// --- [極簡化] "Minimalist" Edition
// --- 1. 查詢 /v3/accounts/balance (帳戶餘額)
// --- 2. 採用與 OKX 腳本一致的錯誤處理與日誌風格
// --- 3. 使用 PropertiesService 管理金鑰
// =======================================================

/**
 * 主函式：獲取並寫入 BitoPro 帳戶餘額
 */
function getBitoProBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'BitoPro Balance';
  let sheet = ss.getSheetByName(sheetName);

  // 若無分頁則自動建立
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['幣種', '總額', '可用', '凍結/質押']); // 初始化標題
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('BITOPRO_API_KEY');
  const apiSecret = props.getProperty('BITOPRO_API_SECRET');

  if (!apiKey || !apiSecret) {
    ss.toast("錯誤：BitoPro API 金鑰或 Secret 未在指令碼屬性中設定。");
    return;
  }

  const baseUrl = 'https://api.bitopro.com/v3';
  var logMessages = [];
  var sheetData = [];

  try {
    ss.toast('正在獲取 BitoPro 帳戶餘額...');

    // --- 1. 獲取帳戶餘額 ---
    const endpoint = '/accounts/balance';
    const json = fetchBitoProApi_(baseUrl, endpoint, 'GET', {}, apiKey, apiSecret);

    if (json && json.data) {
      logMessages.push("1. 帳戶餘額 API OK");

      // 過濾與整理資料
      json.data.forEach(function (b) {
        // ⭐ 修正：官方欄位是 amount (總額), available (可用), stake (圈存/凍結)
        const total = parseFloat(b.amount) || 0;
        const available = parseFloat(b.available) || 0;
        const stake = parseFloat(b.stake) || 0;

        // 只記錄有餘額的幣種
        if (total > 0) {
          sheetData.push([
            b.currency.toUpperCase(),
            total,
            available,
            stake
          ]);
        }
      });

    } else {
      logMessages.push("1. 帳戶餘額 FAILED: " + (json.error || 'Unknown Error'));
    }

    // --- 2. 寫入工作表 ---
    Logger.log("===== BitoPro API 請求日誌 =====");
    Logger.log(logMessages.join("\n"));
    Logger.log("==================================");

    // 排序
    sheetData.sort((a, b) => b[1] - a[1]);

    // 1. 強制寫入標題
    const headers = ['幣種', '總額', '可用', '凍結/質押', '更新時間'];
    sheet.getRange(1, 1, 1, 5).setValues([headers]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');

    // 清除舊資料 (保留標題列)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

    if (sheetData.length > 0) {
      sheet.getRange(2, 1, sheetData.length, 4).setValues(sheetData);
      sheet.getRange(2, 5).setValue(timestamp);
    } else {
      sheet.getRange(2, 1, 1, 5).setValues([['No Assets', 0, 0, 0, timestamp]]);
    }

    Logger.log('成功更新 ' + sheetData.length + ' 種 BitoPro 資產餘額。');
    ss.toast('BitoPro 餘額更新成功！');

  } catch (e) {
    Logger.log('getBitoProBalance 發生嚴重錯誤: ' + e.toString());
    ss.toast('BitoPro 腳本錯誤: ' + e.message);
  }
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

  Logger.log(`----- Requesting BitoPro API: ${endpoint} -----`);

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    // 除錯用：如果還是失敗，印出完整的 Body
    if (responseCode !== 200) {
      Logger.log(`Error Code: ${responseCode}, Body: ${responseText}`);
    }

    return JSON.parse(responseText);

  } catch (e) {
    Logger.log(`fetchBitoProApi_ FAILED: ${e.toString()}`);
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

  // Convert Byte Array to Hex String
  return signatureBytes.map(function (byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}