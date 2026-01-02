// =======================================================
// --- Binance 業務邏輯層 (Service Layer) v2.1 ---
// --- 專注於執行查詢與寫入試算表，模組化設計並整合日誌系統
// =======================================================

/**
 * [核心功能] 獲取 Binance 餘額 (Orchestrator)
 */
function getBinanceBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. 讀取設定
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('BINANCE_API_KEY');
    const apiSecret = props.getProperty('BINANCE_API_SECRET');
    const baseUrl = props.getProperty('TUNNEL_URL');
    const proxyPassword = props.getProperty('PROXY_PASSWORD');

    // 防呆與連線檢查
    if (!apiKey || !apiSecret) {
      const msg = "❌ 錯誤：請先設定 BINANCE_API_KEY 與 SECRET";
      ss.toast(msg);
      LogService.error(msg, "Sync_Binance");
      return;
    }

    if (!baseUrl || !proxyPassword) {
      LogService.warn("跳過：尚未接收到 Tunnel 網址 (電腦可能未開機)，保留舊資料。", "Sync_Binance");
      return;
    }

    const logMessages = [];

    // --- 步驟 A: 獲取現貨餘額 (Spot) ---
    const spotResult = fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    logMessages.push(`Spot: ${spotResult.status}`);
    if (!spotResult.success) LogService.warn(`Spot Fetch Error: ${spotResult.status}`, "Sync_Binance");

    // --- 步驟 B: 獲取理財餘額 (Earn) ---
    const earnResult = fetchEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword);
    logMessages.push(`Earn: ${earnResult.status}`);
    if (!earnResult.success && earnResult.status.startsWith('Failed')) LogService.warn(`Earn Fetch Error: ${earnResult.status}`, "Sync_Binance");

    // --- 步驟 C: 獲取借貸訂單 (Loans) ---
    const loanResult = fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword);
    logMessages.push(`Loans: ${loanResult.status}`);
    if (!loanResult.success && loanResult.status.startsWith('Failed')) LogService.warn(`Loan Fetch Error: ${loanResult.status}`, "Sync_Binance");

    // Console log for debugging
    console.log("===== Binance Update Log =====");
    console.log(logMessages.join("\n"));

    // --- 步驟 D: 整合並寫入資產表 (Balance Sheet) ---
    // 無論成功與否都嘗試寫入 (若失敗則顯示錯誤)
    updateBalanceSheet_(ss, spotResult, earnResult);

    // --- 步驟 E: 寫入借貸表 (Loan Sheet) ---
    // 無論成功與否都嘗試寫入 (若失敗則顯示錯誤)
    updateLoanSheet_(ss, loanResult);

    // --- 最終日誌 ---
    if (spotResult.success || earnResult.success || loanResult.success) {
      LogService.info(`Sync Complete. Spot:${spotResult.status}, Earn:${earnResult.status}, Loans:${loanResult.status}`, "Sync_Binance");
      ss.toast(`Binance 同步部分完成`);
    } else {
      LogService.warn(`Sync All Failed. Spot:${spotResult.status}, Earn:${earnResult.status}, Loans:${loanResult.status}`, "Sync_Binance");
      ss.toast(`Binance 同步全部失敗`);
    }

  } catch (e) {
    LogService.error(`Critical Crash: ${e.message}`, "Sync_Binance");
    ss.toast(`Binance Sync Error: ${e.message}`);
    console.error(e);
  }
}

/**
 * [私有] A. 獲取現貨餘額 (Spot)
 */
function fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/api/v3/account', {}, apiKey, apiSecret, proxyPassword);

  if (res.code !== "0") {
    console.error(`[Spot] API Failed: ${res.msg}`);
    return { success: false, status: "Failed: " + res.msg };
  }

  const balances = new Map();
  // Spot returns object with 'balances' array
  if (res.data && res.data.balances) {
    let count = 0;
    res.data.balances.forEach(b => {
      const free = parseFloat(b.free);
      const locked = parseFloat(b.locked);
      if ((free + locked) > 0) {
        balances.set(b.asset, free + locked);
        count++;
      }
    });
    console.log(`[Spot] Found ${count} non-zero assets. Raw Balances length: ${res.data.balances.length}`);
    if (count > 0) console.log(`[Spot] Top 3 Sample: ${JSON.stringify([...balances.entries()].slice(0, 3))}`);
  } else {
    console.warn(`[Spot] 'balances' field missing in response: ${JSON.stringify(res.data).substring(0, 100)}`);
  }
  return { success: true, status: "OK", data: balances };
}

/**
 * [私有] B. 獲取理財餘額 (Simple Earn Flexible)
 */
function fetchEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword) {
  // Simple Earn Flexible Positions
  const res = fetchBinanceApi_(baseUrl, '/sapi/v1/simple-earn/flexible/position', { limit: 100 }, apiKey, apiSecret, proxyPassword);

  if (res.code !== "0") {
    console.warn(`[Earn] API Failed (might be empty/no perm): ${res.msg}`);
    // 某些帳戶可能無 Simple Earn 權限或未開通，視為空
    return { success: false, status: "Failed: " + res.msg, data: new Map() };
  }

  const balances = new Map();
  // SAPI usually returns array directly or { rows: [] }
  const rows = Array.isArray(res.data) ? res.data : (res.data.rows || []);

  console.log(`[Earn] API returned ${rows.length} positions.`);

  rows.forEach(row => {
    const asset = row.asset;
    const amount = parseFloat(row.totalAmount);
    if (amount > 0) {
      const current = balances.get(asset) || 0;
      balances.set(asset, current + amount);
    }
  });

  if (balances.size > 0) console.log(`[Earn] Processed ${balances.size} assets. Sample: ${JSON.stringify([...balances.entries()].slice(0, 3))}`);

  return { success: true, status: "OK", data: balances };
}

/**
 * [私有] C. 獲取借貸訂單 (Flexible Loan)
 */
function fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword) {
  // Flexible Loan Ongoing Orders
  const res = fetchBinanceApi_(baseUrl, '/sapi/v2/loan/flexible/ongoing/orders', { limit: 100 }, apiKey, apiSecret, proxyPassword);

  if (res.code !== "0") {
    console.error(`[Loans] API Failed: ${res.msg}`);
    return { success: false, status: "Failed: " + res.msg, data: [] };
  }

  const rows = Array.isArray(res.data) ? res.data : (res.data.rows || []);
  console.log(`[Loans] Found ${rows.length} active loan orders.`);

  const orders = rows.map(r => ({
    loanCoin: r.loanCoin,
    totalDebt: parseFloat(r.totalDebt),
    collateralCoin: r.collateralCoin,
    collateralAmount: parseFloat(r.collateralAmount),
    currentLTV: parseFloat(r.currentLTV)
  }));

  return { success: true, status: "OK", data: orders };
}

/**
 * [工具] 更新資產總表 (Spot + Earn)
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {Object} spotResult
 * @param {Object} earnResult
 */
function updateBalanceSheet_(ss, spotResult, earnResult) {
  let sheet = ss.getSheetByName('Binance Balance');
  if (!sheet) sheet = ss.insertSheet('Binance Balance');

  // 1. 寫入固定標題 (Enforce Headers)
  sheet.getRange('A1:C1').setValues([['Currency', 'Total', 'Last Updated']]);
  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f3f3');

  // 若兩者皆失敗，顯示錯誤
  if (!spotResult.success && !earnResult.success) {
    sheet.getRange('A2:C').clearContent();
    sheet.getRange('A2:A3').setValues([
      [`Error: Spot - ${spotResult.status}`],
      [`Error: Earn - ${earnResult.status}`]
    ]);
    sheet.getRange('A2:A3').setFontColor('red');
    sheet.getRange('C2').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"));
    return;
  }

  // 正常合併邏輯
  const spotData = spotResult.success ? spotResult.data : new Map();
  const earnData = earnResult.success ? earnResult.data : new Map();

  console.log(`[UpdateSheet] Spot Success: ${spotResult.success}, Data Type: ${typeof spotData}, IsMap: ${spotData instanceof Map}`);
  console.log(`[UpdateSheet] Earn Success: ${earnResult.success}, Data Type: ${typeof earnData}, IsMap: ${earnData instanceof Map}`);

  const combinedTotals = new Map();

  // Helper to merge data into combinedTotals
  const mergeData = (data, sourceName) => {
    if (!data) return;

    let count = 0;
    // Case A: Map (Iterable)
    if (data instanceof Map) {
      data.forEach((val, key) => {
        const current = combinedTotals.get(key) || 0;
        combinedTotals.set(key, current + val);
        count++;
      });
      console.log(`[UpdateSheet] Merged ${count} items from ${sourceName} (Map).`);
      return;
    }

    // Case B: Plain Object (Not Iterable)
    if (typeof data === 'object') {
      Object.keys(data).forEach(key => {
        const val = data[key];
        const current = combinedTotals.get(key) || 0;
        combinedTotals.set(key, current + val);
        count++;
      });
      console.log(`[UpdateSheet] Merged ${count} items from ${sourceName} (Object).`);
      return;
    }

    // Case C: Logging unexpected type
    console.warn(`[Sync_Binance] Unexpected data type for merge: ${typeof data}`);
  };

  mergeData(spotData, "Spot");
  mergeData(earnData, "Earn");

  const sheetData = [];
  combinedTotals.forEach((val, asset) => {
    if (val > 0.00000001) sheetData.push([asset, val]);
  });
  sheetData.sort((a, b) => b[1] - a[1]);

  console.log(`[UpdateSheet] Final Sheet Data Rows: ${sheetData.length}`);

  // 清除舊資料
  sheet.getRange('A2:C').clearContent();
  sheet.getRange('A2:C').setFontColor('black');

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

  if (sheetData.length > 0) {
    sheet.getRange(2, 1, sheetData.length, 2).setValues(sheetData);
    sheet.getRange(2, 3).setValue(timestamp);
    ss.toast(`TwBinance 資產更新成功！(${sheetData.length} 幣種)`);
  } else {
    // 明確顯示無資產 (No Assets Found)
    sheet.getRange(2, 1, 1, 3).setValues([['No Assets Found (Check API/Log)', 0, timestamp]]);
    console.warn("[UpdateSheet] No non-zero assets found to write.");
    ss.toast("Binance 同步完成但無資產 (0 幣種)");
  }
}

/**
 * [工具] 更新借貸表 (Loans)
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {Object} loanResult
 */
function updateLoanSheet_(ss, loanResult) {
  const SHEET_NAME = 'Binance Loans';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 初始化分頁
  if (!sheet) {
    if (loanResult.success && loanResult.data.length > 0) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
    } else if (!loanResult.success) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Action', '', '', 'Updated']]);
    } else {
      return;
    }
  }

  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();

  if (loanResult.success && loanResult.data.length > 0) {
    const rows = loanResult.data.map(o => [
      o.loanCoin, o.totalDebt, o.collateralCoin, o.collateralAmount, o.currentLTV, new Date()
    ]);

    // 復原表頭 (若之前是錯誤顯示)
    if (sheet.getRange(1, 1).getValue() === 'Status') {
      sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
    }

    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    sheet.getRange('A2:F').setFontColor('black');
    ss.toast(`Binance 借貸更新：${rows.length} 筆訂單`);

  } else if (!loanResult.success) {
    // 錯誤顯示
    const errorMsg = loanResult.status;
    let action = "Check Logs";
    if (errorMsg.includes('403')) action = "Check Proxy/API Key";

    sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Action', '', '', 'Updated']]);
    sheet.getRange(2, 1, 1, 6).setValues([['Error', errorMsg, action, '', '', new Date()]]);
    sheet.getRange(2, 1, 1, 3).setFontColor('red');
  }
}

/**
 * [私有工具] 連線核心 (不變)
 */
function fetchBinanceApi_(baseUrl, endpoint, params, apiKey, apiSecret, proxyPassword) {
  const timestamp = new Date().getTime();
  params.timestamp = timestamp;
  params.recvWindow = 10000;

  const queryString = Object.keys(params).map(key => `${key}=${encodeURIComponent(params[key])}`).join('&');
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, queryString, apiSecret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  const url = `${baseUrl}${endpoint}?${queryString}&signature=${signature}`;

  const options = {
    'method': 'GET',
    'headers': {
      'X-MBX-APIKEY': apiKey,
      'Content-Type': 'application/json',
      'x-proxy-auth': proxyPassword
    },
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 403) return { code: "-403", msg: "Proxy Auth Failed" };

    let json;
    try { json = JSON.parse(text); } catch (e) { return { code: "-2", msg: "Invalid JSON" }; }

    if (json.msg || (json.code && json.code !== 0)) {
      return { code: json.code || "-999", msg: json.msg, data: json };
    }
    return { code: "0", data: json };

  } catch (e) {
    return { code: "-1", msg: "Connection Error: " + e.message };
  }
}