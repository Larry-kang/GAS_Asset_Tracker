/**
 * OKX 餘額同步系統 (v24.5 - Modular Edition)
 * 修正：恢復使用 Flexible Loan API，並提示使用者開啟權限
 */

/**
 * [總指揮] 獲取 OKX 餘額與債務 (Orchestrator)
 */
function getOkxBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('OKX_API_KEY');
    const apiSecret = props.getProperty('OKX_API_SECRET');
    const apiPassphrase = props.getProperty('OKX_API_PASSPHRASE');
    const baseUrl = 'https://www.okx.com';

    if (!apiKey || !apiSecret || !apiPassphrase) {
      ss.toast("❌ 錯誤：請先設定 OKX_API_KEY");
      return;
    }

    const logMessages = [];

    // --- 步驟 A: 獲取帳戶餘額 (Spot/Funding) ---
    ss.toast('正在獲取 OKX 帳戶餘額 (1/3)...');
    const accountResult = fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Account: ${accountResult.status}`);

    // --- 步驟 B: 獲取理財餘額 (Simple Earn) ---
    ss.toast('正在獲取 OKX 理財餘額 (2/3)...');
    const earnResult = fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Earn: ${earnResult.status}`);

    // --- 步驟 C: 獲取借貸訂單 (Flexible Loans) ---
    // [修正] 恢復使用專用 API，因為 Account.liab 無法顯示完整質押詳情
    ss.toast('正在獲取 OKX 借貸資訊 (3/3)...');
    const loanResult = fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Loans: ${loanResult.status}`); // 若 Forbidden 會顯示於此

    console.log("===== OKX Update Log =====");
    console.log(logMessages.join("\n"));

    // --- 步驟 D: 整合並寫入資產表 ---
    if (accountResult.success || earnResult.success) {
      updateBalanceSheet_(ss, accountResult.data, earnResult.data);
    }

    // --- 步驟 E: 寫入借貸表 ---
    if (loanResult.success) {
      updateLoanSheet_(ss, loanResult.data);
    } else if (loanResult.status.includes('Forbidden')) {
      // 若權限不足，嘗試清空或寫入提示
      Logger.log("⚠️ 借貸抓取失敗 (Forbidden)：請檢查 API Key 是否開啟 'Read' 權限");
    }

    const finalMsg = `OKX 同步完成！`;
    ss.toast(finalMsg);

  } catch (e) {
    Logger.log(`Critical Crash: ${e.message}`);
    ss.toast(`Error: ${e.message}`);
  }
}

// ... fetchOkxAccount_ (還原為純餘額抓取，移除 liab 混合邏輯) ...
function fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
  const balances = new Map();

  if (res.code === "0" && res.data && res.data[0] && res.data[0].details) {
    res.data[0].details.forEach(b => {
      const total = (parseFloat(b.availBal) || 0) + (parseFloat(b.frozenBal) || 0);
      if (total > 0) balances.set(b.ccy, total);
    });
    return { success: true, status: "OK", data: balances };
  } else {
    return { success: false, status: `Failed (${res.msg || 'No Data'})`, data: balances };
  }
}

// ... fetchOkxEarn_ (保持不變) ...
function fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/finance/savings/balance', {}, apiKey, apiSecret, apiPassphrase);
  const positions = new Map();
  if (res.code === "0" && res.data) {
    res.data.forEach(b => {
      if (parseFloat(b.amt) > 0) positions.set(b.ccy, (positions.get(b.ccy) || 0) + parseFloat(b.amt));
    });
    return { success: true, status: "OK", data: positions };
  }
  return { success: res.code === "0", status: `Info (${res.msg})`, data: positions };
}

// ... fetchOkxLoans_ (恢復並修正 endpoint) ...
function fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  // 嘗試使用 borrow-info 與 orders (先試 orders，因為它有明細)
  // 若 Forbidden，則是用戶權限問題，無法透過改 endpoint 解決
  const endpoint = '/api/v5/finance/flexible-loan/orders';
  const res = fetchOkxApi_(baseUrl, endpoint, { state: 'alive' }, apiKey, apiSecret, apiPassphrase);

  if (res.code === "0" && res.data) {
    const orders = res.data.map(order => ({
      loanCoin: order.loanCcy,
      totalDebt: parseFloat(order.loanAmt || 0) + parseFloat(order.interest || 0),
      collateralCoin: order.collateralCcy,
      collateralAmount: parseFloat(order.collateralAmt || 0),
      currentLTV: parseFloat(order.curLtv || 0)
    }));
    return { success: true, status: "OK", data: orders };
  } else {
    // [Debug] Log specific error code
    const errorMsg = `Code: ${res.code}, Msg: ${res.msg}`;
    console.log(`[FetchLoans] Error: ${errorMsg}`);
    return { success: false, status: `Info (${errorMsg})`, data: [] };
  }
}

// ... updateBalanceSheet_ (保持不變) ...
function updateBalanceSheet_(ss, spotMap, earnMap) {
  const SHEET_NAME = 'OKX Balance';
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  const combined = new Map(spotMap);
  earnMap.forEach((v, k) => combined.set(k, (combined.get(k) || 0) + v));

  const rows = [];
  combined.forEach((v, k) => { if (v > 0) rows.push([k, v]); });
  rows.sort((a, b) => b[1] - a[1]);

  sheet.getRange('A2:C').clearContent();
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    sheet.getRange(2, 3).setValue(new Date());
  }
}

// ... updateLoanSheet_ (保持不變) ...
function updateLoanSheet_(ss, orders) {
  const SHEET_NAME = 'OKX Loans';
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet && orders.length > 0) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
  } else if (!sheet) return;

  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();

  if (orders.length > 0) {
    const rows = orders.map(o => [o.loanCoin, o.totalDebt, o.collateralCoin, o.collateralAmount, o.currentLTV, new Date()]);
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

// ... fetchOkxApi_ & getOkxSignature_ (保持不變) ...
function fetchOkxApi_(baseUrl, endpoint, params, apiKey, apiSecret, apiPassphrase) {
  // ... (Standard Implementation) ...
  const method = 'GET';
  const timestamp = new Date().toISOString();
  let qs = '';
  if (params && Object.keys(params).length) {
    qs = '?' + Object.keys(params).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
  }
  const url = baseUrl + endpoint + qs;
  const signStr = timestamp + method + endpoint + qs; // Note: QS logic
  const signature = getOkxSignature_(signStr, apiSecret);

  const options = {
    method: 'GET',
    headers: {
      'OK-ACCESS-KEY': apiKey,
      'OK-ACCESS-SIGN': signature,
      'OK-ACCESS-TIMESTAMP': timestamp,
      'OK-ACCESS-PASSPHRASE': apiPassphrase,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  try {
    return JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  } catch (e) { return { code: "-1", msg: e.message }; }
}

function getOkxSignature_(str, secret) {
  return Utilities.base64Encode(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, str, secret));
}