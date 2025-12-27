// =======================================================
// --- Binance 業務邏輯層 (Service Layer) v2.1 ---
// --- 專注於執行查詢與寫入試算表，模組化 설계並整合日誌系統
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
    let balanceUpdated = false;
    if (spotResult.success || earnResult.success) {
      updateBalanceSheet_(ss, spotResult.data, earnResult.data);
      balanceUpdated = true;
    } else {
      console.log("⚠️ Spot 與 Earn 皆未更新，跳過寫入餘額表");
    }

    // --- 步驟 E: 寫入借貸表 (Loan Sheet) ---
    let loansUpdated = false;
    if (loanResult.success) {
      updateLoanSheet_(ss, loanResult.data);
      loansUpdated = true;
    }

    // --- 最終日誌 ---
    if (balanceUpdated || loansUpdated) {
      LogService.info(`Sync Complete. Spot:${spotResult.status}, Earn:${earnResult.status}, Loans:${loanResult.status}`, "Sync_Binance");
    } else {
      LogService.warn(`Sync Incomplete. No data written to sheets.`, "Sync_Binance");
    }

  } catch (e) {
    LogService.error(`Critical Crash: ${e.message}`, "Sync_Binance");
    ss.toast(`Binance Sync Error: ${e.message}`);
    console.error(e);
  }
}

/**
 * [模組 A] 獲取 Spot 餘額
 * @returns {Object} { success: boolean, status: string, data: Map<string, number> }
 */
function fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/api/v3/account', {}, apiKey, apiSecret, proxyPassword);
  const balances = new Map();

  if (res.code === "0" && res.data && res.data.balances) {
    res.data.balances.forEach(b => {
      if (b.asset.startsWith('LD')) return; // 忽略 LD 幣 (Earn 舊格式)
      const total = parseFloat(b.free) + parseFloat(b.locked);
      if (total > 0) {
        balances.set(b.asset, total);
      }
    });
    return { success: true, status: "OK", data: balances };
  } else {
    return { success: false, status: `Failed (${res.msg})`, data: balances };
  }
}

/**
 * [模組 B] 獲取 Earn 餘額
 * @returns {Object} { success: boolean, status: string, data: Map<string, number> }
 */
function fetchEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/sapi/v1/simple-earn/flexible/position', { size: 100 }, apiKey, apiSecret, proxyPassword);
  const positions = new Map();

  if (res.code === "0" && res.data && Array.isArray(res.data.rows)) {
    res.data.rows.forEach(p => {
      const total = parseFloat(p.totalAmount);
      if (total > 0) {
        const currentTotal = positions.get(p.asset) || 0;
        positions.set(p.asset, currentTotal + total);
      }
    });
    return { success: true, status: "OK", data: positions };
  } else {
    // Earn 為空或失敗視為正常流程的一部分 (可能沒買)
    const msg = res.msg || 'No Data';
    return { success: res.code === "0", status: `Info (${msg})`, data: positions };
  }
}

/**
 * [模組 C] 獲取 Loan 訂單
 * @returns {Object} { success: boolean, status: string, data: Array<Object> }
 */
function fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword) {
  // 使用 v2 接口查詢進行中的活期借幣訂單
  const res = fetchBinanceApi_(baseUrl, '/sapi/v2/loan/flexible/ongoing/orders', {}, apiKey, apiSecret, proxyPassword);

  if (res.code === "0" && res.data && Array.isArray(res.data.rows)) {
    // 轉換資料格式
    const orders = res.data.rows.map(order => ({
      loanCoin: order.loanCoin,
      totalDebt: parseFloat(order.totalDebt),
      collateralCoin: order.collateralCoin,
      collateralAmount: parseFloat(order.collateralAmount),
      currentLTV: parseFloat(order.currentLTV)
    }));
    return { success: true, status: "OK", data: orders };
  } else {
    const msg = res.msg || 'No Loan';
    return { success: res.code === "0", status: `Info (${msg})`, data: [] };
  }
}

/**
 * [工具] 更新資產總表 (Spot + Earn)
 */
function updateBalanceSheet_(ss, spotMap, earnMap) {
  let sheet = ss.getSheetByName('Binance Balance');
  if (!sheet) sheet = ss.insertSheet('Binance Balance');

  // 合併 Spot 與 Earn
  const combinedTotals = new Map(spotMap);
  earnMap.forEach((val, asset) => {
    const current = combinedTotals.get(asset) || 0;
    combinedTotals.set(asset, current + val);
  });

  const sheetData = [];
  combinedTotals.forEach((val, asset) => {
    if (val > 0.00000001) sheetData.push([asset, val]);
  });
  sheetData.sort((a, b) => b[1] - a[1]);

  if (sheetData.length > 0) {
    sheet.getRange('A2:C').clearContent();
    sheet.getRange(2, 1, sheetData.length, 2).setValues(sheetData);
    sheet.getRange(2, 3).setValue(new Date());
    ss.toast(`Binance 資產更新成功！(${sheetData.length} 幣種)`);
  }
}

/**
 * [工具] 更新借貸表 (Loans)
 */
function updateLoanSheet_(ss, loanOrders) {
  const SHEET_NAME = 'Binance Loans';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 如果沒有該分頁且有資料，則建立；若沒資料且沒分頁，則不動作
  if (!sheet && loanOrders.length > 0) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 設定表頭
    sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
  } else if (!sheet) {
    return;
  }

  // 清除舊資料
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
  }

  if (loanOrders.length > 0) {
    const rows = loanOrders.map(o => [
      o.loanCoin,
      o.totalDebt,
      o.collateralCoin,
      o.collateralAmount,
      o.currentLTV,
      new Date()
    ]);
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    ss.toast(`Binance 借貸更新：${rows.length} 筆訂單`);
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