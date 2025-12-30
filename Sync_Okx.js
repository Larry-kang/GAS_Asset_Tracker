/**
 * OKX 餘額同步系統 (v24.5 - Modular Edition)
 * 專注於執行查詢與寫入試算表，模組化設計並整合借貸監控
 */

/**
 * [總指揮] 獲取 OKX 餘額與債務 (Orchestrator)
 */
function getOkxBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. 讀取設定
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('OKX_API_KEY');
    const apiSecret = props.getProperty('OKX_API_SECRET');
    const apiPassphrase = props.getProperty('OKX_API_PASSPHRASE');
    const baseUrl = 'https://www.okx.com';

    // 防呆檢查
    if (!apiKey || !apiSecret || !apiPassphrase) {
      const msg = "❌ 錯誤：請先在腳本屬性中設定 OKX_API_KEY, OKX_API_SECRET, OKX_API_PASSPHRASE";
      ss.toast(msg);
      Logger.log(msg);
      return;
    }

    const logMessages = [];

    // --- 步驟 A: 獲取帳戶餘額 (Spot/Funding) ---
    ss.toast('正在獲取 OKX 帳戶餘額 (1/3)...');
    const accountResult = fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Account: ${accountResult.status}`);
    if (!accountResult.success) Logger.log(`Account Fetch Error: ${accountResult.status}`);

    // --- 步驟 B: 獲取理財餘額 (Simple Earn) ---
    ss.toast('正在獲取 OKX 理財餘額 (2/3)...');
    const earnResult = fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Earn: ${earnResult.status}`);
    if (!earnResult.success) Logger.log(`Earn Fetch Error: ${earnResult.status}`);

    // --- 步驟 C: 獲取借貸訂單 (Flexible Loans) ---
    ss.toast('正在獲取 OKX 借貸資訊 (3/3)...');
    const loanResult = fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Loans: ${loanResult.status}`);
    if (!loanResult.success) Logger.log(`Loan Fetch Error: ${loanResult.status}`);

    // Console log for debugging
    console.log("===== OKX Update Log =====");
    console.log(logMessages.join("\n"));

    // --- 步驟 D: 整合並寫入資產表 (Balance Sheet) ---
    let balanceUpdated = false;
    if (accountResult.success || earnResult.success) {
      updateBalanceSheet_(ss, accountResult.data, earnResult.data);
      balanceUpdated = true;
    } else {
      console.log("⚠️ Account 與 Earn 皆未更新，跳過寫入 'OKX Balance' 工作表");
    }

    // --- 步驟 E: 寫入借貸表 (Loan Sheet) ---
    let loansUpdated = false;
    if (loanResult.success) {
      updateLoanSheet_(ss, loanResult.data);
      loansUpdated = true;
    }

    // --- 最終通知 ---
    const finalMsg = `OKX 同步完成！\n餘額: ${balanceUpdated ? '更新' : '跳過'}, 借貸: ${loansUpdated ? '更新' : '跳過'}`;
    ss.toast(finalMsg);
    Logger.log(finalMsg);

  } catch (e) {
    Logger.log(`Critical Crash in Sync_Okx: ${e.message}`);
    ss.toast(`OKX Sync Error: ${e.message}`);
    console.error(e);
  }
}

/**
 * [模組 A] 獲取 Account 餘額 (含資金與交易帳戶)
 * API: /api/v5/account/balance
 */
function fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
  const balances = new Map();

  if (res.code === "0" && res.data && res.data[0] && res.data[0].details) {
    res.data[0].details.forEach(b => {
      // OKX 回傳 availBal (可用) + frozenBal (凍結/掛單)
      // 注意：cashBal 是現金餘額，eq 是權益(美金計價)，這裡是抓數量
      // 我們使用 cashBal (現金餘額) 或 availBal + frozenBal。
      // 通常 cashBal 包含了所有該幣種的數量 (除了借貸負債可能不在此顯現)
      // 這裡採用 availBal + frozenBal 較為保險
      const total = (parseFloat(b.availBal) || 0) + (parseFloat(b.frozenBal) || 0);
      if (total > 0) {
        balances.set(b.ccy, total);
      }
    });
    return { success: true, status: "OK", data: balances };
  } else {
    return { success: false, status: `Failed (${res.msg || 'No Data'})`, data: balances };
  }
}

/**
 * [模組 B] 獲取 Simple Earn 餘額
 * API: /api/v5/finance/savings/balance
 */
function fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/finance/savings/balance', {}, apiKey, apiSecret, apiPassphrase);
  const positions = new Map();

  if (res.code === "0" && res.data && Array.isArray(res.data)) {
    res.data.forEach(b => {
      const total = parseFloat(b.amt) || 0;
      if (total > 0) {
        const ccy = b.ccy;
        const currentTotal = positions.get(ccy) || 0;
        positions.set(ccy, currentTotal + total);
      }
    });
    return { success: true, status: "OK", data: positions };
  } else {
    // 沒資料視為成功 (可能沒買理財)
    return { success: res.code === "0", status: `Info (${res.msg || 'No Earn Data'})`, data: positions };
  }
}

/**
 * [模組 C] 獲取 Flexible Loan 訂單
 * API: /api/v5/finance/flexible-loan/orders (正在進行的訂單)
 */
function fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  // 查詢狀態為 'alive' (有效) 的訂單
  const res = fetchOkxApi_(baseUrl, '/api/v5/finance/flexible-loan/orders', { state: 'alive' }, apiKey, apiSecret, apiPassphrase);

  if (res.code === "0" && res.data && Array.isArray(res.data)) {
    // 轉換資料格式
    const orders = res.data.map(order => ({
      loanCoin: order.loanCcy,
      totalDebt: parseFloat(order.loanAmt || order.borrowAmt || 0) + parseFloat(order.interest || 0), // 借款 + 利息
      collateralCoin: order.collateralCcy,
      collateralAmount: parseFloat(order.collateralAmt || 0),
      currentLTV: parseFloat(order.curLtv || 0) // LTV
    }));
    return { success: true, status: "OK", data: orders };
  } else {
    // 404 或空資料
    const msg = res.msg || 'No Loan Data';
    return { success: res.code === "0", status: `Info (${msg})`, data: [] };
  }
}

/**
 * [寫入模組] 更新資產總表 (Account + Earn) -> 'OKX Balance'
 */
function updateBalanceSheet_(ss, spotMap, earnMap) {
  const SHEET_NAME = 'OKX Balance';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

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
    Logger.log(`[OKX Balance] Updated ${sheetData.length} rows.`);
  }
}

/**
 * [寫入模組] 更新借貸表 (Loans) -> 'OKX Loans'
 */
function updateLoanSheet_(ss, loanOrders) {
  const SHEET_NAME = 'OKX Loans';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 如果有資料但沒分頁 -> 建立
  // 如果有分頁 -> 清空並更新
  // 如果沒資料且沒分頁 -> 不動作
  if (!sheet && loanOrders.length > 0) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
  } else if (!sheet) {
    return;
  }

  // 清除舊數據 (保留標題)
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
    Logger.log(`[OKX Loans] Updated ${rows.length} orders.`);
  }
}

/**
 * [私有工具] OKX API 連線核心
 */
function fetchOkxApi_(baseUrl, endpoint, params, apiKey, apiSecret, apiPassphrase) {
  const method = 'GET';
  const timestamp = new Date().toISOString(); // ISO format for OKX

  let queryString = '';
  if (params && Object.keys(params).length > 0) {
    queryString = '?' + Object.keys(params).map(function (key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    }).join('&');
  }

  const url = baseUrl + endpoint + queryString;
  // OKX 簽名規則: timestamp + method + requestPath (+ body if POST)
  const stringToSign = timestamp + method + endpoint + queryString.replace('?', '?'); // Query string logic varies, OKX V5 usually includes query in path for sign
  // 修正簽名路徑: 如果有 query params，必須包含在 endpoint 內進行簽名
  const signPath = endpoint + queryString;
  const stringToSignCorrect = timestamp + method + signPath;

  const signature = getOkxSignature_(stringToSignCorrect, apiSecret);

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
    const json = JSON.parse(text);
    return json;
  } catch (e) {
    return { code: "-1", msg: "Connection Error: " + e.message };
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