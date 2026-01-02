// =======================================================
// --- Binance 業務邏輯層 (Service Layer) v2.2 ---
// --- Refactored to use Lib_SyncManager
// =======================================================

/**
 * [核心功能] 獲取 Binance 餘額 (Orchestrator)
 */
function getBinanceBalance() {
  const MODULE_NAME = "Sync_Binance";

  // 使用 SyncManager 包裝執行邏輯 (自動錯誤處理與日誌)
  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. 讀取設定
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('BINANCE_API_KEY');
    const apiSecret = props.getProperty('BINANCE_API_SECRET');
    const baseUrl = props.getProperty('TUNNEL_URL');
    const proxyPassword = props.getProperty('PROXY_PASSWORD');

    // 防呆與連線檢查
    if (!apiKey || !apiSecret) {
      const msg = "❌ 錯誤：請先設定 BINANCE_API_KEY 與 SECRET";
      SyncManager.log("ERROR", msg, MODULE_NAME);
      return;
    }

    if (!baseUrl || !proxyPassword) {
      SyncManager.log("WARNING", "跳過：尚未接收到 Tunnel 網址 (電腦可能未開機)，保留舊資料。", MODULE_NAME);
      return;
    }

    // --- 步驟 A: 獲取現貨餘額 (Spot) ---
    const spotResult = fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (!spotResult.success) SyncManager.log("WARNING", `Spot Fetch Error: ${spotResult.status}`, MODULE_NAME);

    // --- 步驟 B: 獲取理財餘額 (Earn) ---
    const earnResult = fetchEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (!earnResult.success && earnResult.status.startsWith('Failed')) SyncManager.log("WARNING", `Earn Fetch Error: ${earnResult.status}`, MODULE_NAME);

    // --- 步驟 C: 獲取借貸訂單 (Loans) ---
    const loanResult = fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (!loanResult.success && loanResult.status.startsWith('Failed')) SyncManager.log("WARNING", `Loan Fetch Error: ${loanResult.status}`, MODULE_NAME);

    // --- 步驟 D: 整合並寫入資產表 (Balance Sheet) ---
    // 使用 SyncManager.mergeAssets 自動合併 Map 或 Object
    const combinedData = SyncManager.mergeAssets(
      spotResult.success ? spotResult.data : null,
      earnResult.success ? earnResult.data : null
    );

    // 標準化寫入 Sheet (包含 Header, Sort, Time, Empty State)
    SyncManager.writeToSheet(ss, 'Binance Balance', ['Currency', 'Total', 'Last Updated'], combinedData);

    // --- 步驟 E: 寫入借貸表 (Loan Sheet) ---
    // Loan Sheet 比較特殊 (多欄位)，暫時維持獨立函式，但使用 SyncManager 的 Log
    updateLoanSheet_(ss, loanResult);

    SyncManager.log("INFO", "Binance Sync Workflow Completed", MODULE_NAME);
  });
}

/**
 * [私有] A. 獲取現貨餘額 (Spot)
 */
function fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/api/v3/account', {}, apiKey, apiSecret, proxyPassword);

  if (res.code !== "0") {
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
    // 使用 SyncManager.log 但降級為 DEBUG 性質 (顯示在 console)
    console.log(`[Spot] Found ${count} non-zero assets.`);
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
    // 某些帳戶可能無 Simple Earn 權限或未開通，視為空
    return { success: false, status: "Failed: " + res.msg, data: new Map() };
  }

  const balances = new Map();
  const rows = Array.isArray(res.data) ? res.data : (res.data.rows || []);

  rows.forEach(row => {
    const asset = row.asset;
    const amount = parseFloat(row.totalAmount);
    if (amount > 0) {
      const current = balances.get(asset) || 0;
      balances.set(asset, current + amount);
    }
  });

  return { success: true, status: "OK", data: balances };
}

/**
 * [私有] C. 獲取借貸訂單 (Flexible Loan)
 */
function fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword) {
  // Flexible Loan Ongoing Orders
  const res = fetchBinanceApi_(baseUrl, '/sapi/v2/loan/flexible/ongoing/orders', { limit: 100 }, apiKey, apiSecret, proxyPassword);

  if (res.code !== "0") {
    return { success: false, status: "Failed: " + res.msg, data: [] };
  }

  const rows = Array.isArray(res.data) ? res.data : (res.data.rows || []);

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
 * [工具] 更新借貸表 (Loans)
 * 保留此獨立函式因為其資料結構特殊 (多欄位)，不適合用通用的 Key-Value Writer
 * 但使用 SyncManager.log
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
      sheet = ss.insertSheet(SHEET_NAME); // Error State Header
      sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Action', '', '', 'Updated']]);
    } else {
      return; // No Loans, No Sheet needed
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
    SyncManager.log("INFO", `Binance Loans Updated: ${rows.length} orders`, "Sync_Binance");

  } else if (!loanResult.success) {
    // 錯誤顯示
    const errorMsg = loanResult.status;
    let action = "Check Logs";
    if (errorMsg.includes('403')) action = "Check Proxy/API Key";

    sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Action', '', '', 'Updated']]);
    sheet.getRange(2, 1, 1, 6).setValues([['Error', errorMsg, action, '', '', new Date()]]);
    SyncManager.log("WARNING", `Binance Loan Error Displayed: ${errorMsg}`, "Sync_Binance");
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