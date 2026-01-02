/**
 * OKX 餘額同步系統 (v24.5 - Modular & Standardized)
 * 重構：使用 Lib_SyncManager 標準化流程 (Log/Merge/Sheet)
 */

/**
 * [總指揮] 獲取 OKX 餘額與債務 (Orchestrator)
 */
function getOkxBalance() {
  const MODULE_NAME = "Sync_Okx";

  // 使用 SyncManager 包裝執行邏輯
  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('OKX_API_KEY');
    const apiSecret = props.getProperty('OKX_API_SECRET');
    const apiPassphrase = props.getProperty('OKX_API_PASSPHRASE');
    const baseUrl = 'https://www.okx.com';

    if (!apiKey || !apiSecret || !apiPassphrase) {
      SyncManager.log("ERROR", "❌ 錯誤：請先設定 OKX_API_KEY, SECRET, PASSPHRASE", MODULE_NAME);
      return;
    }

    // --- 步驟 A: 獲取帳戶餘額 (Spot/Funding) ---
    const accountResult = fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (!accountResult.success) SyncManager.log("WARNING", `Account Fetch Error: ${accountResult.status}`, MODULE_NAME);

    // --- 步驟 B: 獲取理財餘額 (Simple Earn) ---
    const earnResult = fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (!earnResult.success) SyncManager.log("WARNING", `Earn Fetch Error: ${earnResult.status}`, MODULE_NAME);

    // --- 步驟 C: 獲取借貸訂單 (Flexible Loans) ---
    const loanResult = fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (!loanResult.success && loanResult.status.includes('Forbidden')) {
      SyncManager.log("WARNING", `Loan Fetch Forbidden (API Key need 'Trade' perm).`, MODULE_NAME);
    }

    // --- 步驟 D: 整合並寫入資產表 ---
    // 使用 SyncManager.mergeAssets
    const combinedData = SyncManager.mergeAssets(
      accountResult.success ? accountResult.data : null,
      earnResult.success ? earnResult.data : null
    );

    // 標準化寫入 Sheet
    SyncManager.writeToSheet(ss, 'OKX Balance', ['Currency', 'Total', 'Last Updated'], combinedData);

    // --- 步驟 E: 寫入借貸表 ---
    updateLoanSheet_(ss, loanResult);

    SyncManager.log("INFO", "OKX Sync Complete", MODULE_NAME);
  });
}

// ... fetchOkxAccount_ (MODIFIED: Net Balance = Avail + Frozen - Liab) ...
function fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
  const balances = new Map();

  if (res.code === "0" && res.data && res.data[0] && res.data[0].details) {
    let count = 0;
    res.data[0].details.forEach(b => {
      // 淨餘額 = 可用 + 凍結 - 負債 (liab)
      // 若有負債，liab 會是正數，需減除
      const total = (parseFloat(b.availBal) || 0) + (parseFloat(b.frozenBal) || 0) - (parseFloat(b.liab) || 0);

      // 只要不等於 0 都要紀錄 (包含負值)
      if (Math.abs(total) > 1e-8) {
        balances.set(b.ccy, total);
        count++;
      }
    });
    console.log(`[OkxAccount] Found ${count} assets.`);
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
/**
 * [模組 C] 獲取 Flexible Loan 訂單
 */
function fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase) {
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
    const errorMsg = `Code: ${res.code}, Msg: ${res.msg}`;

    // 提供明確的 Debug 提示
    let statusHint = errorMsg;
    if (res.code == "50100" || (res.msg && res.msg.toLowerCase().includes("forbidden"))) {
      statusHint = "Forbidden (Please try enabling 'Trade' permission on API Key)";
    }

    return { success: false, status: statusHint, data: [] };
  }
}

// ... updateLoanSheet_ (保留獨立邏輯但使用 SyncManager.log) ...
function updateLoanSheet_(ss, loanResult) {
  const SHEET_NAME = 'OKX Loans';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 若無工作表且有資料或有錯誤，則建立
  if (!sheet) {
    if (loanResult.success && loanResult.data.length > 0) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
    } else if (!loanResult.success) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Suggested Action', '', '', 'Updated']]);
    } else {
      return; // 沒錯也沒資料，不動作
    }
  }

  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();

  if (loanResult.success && loanResult.data.length > 0) {
    const orders = loanResult.data;
    const rows = orders.map(o => [o.loanCoin, o.totalDebt, o.collateralCoin, o.collateralAmount, o.currentLTV, new Date()]);

    // 確保表頭回到正常模式
    if (sheet.getRange(1, 2).getValue() !== 'Debt') {
      sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
    }
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    SyncManager.log("INFO", `OKX Loans updated: ${rows.length} items`, "Sync_Okx");

  } else if (!loanResult.success) {
    // 顯示錯誤資訊
    const errorMsg = loanResult.status;
    const action = errorMsg.includes('Forbidden') ? "Enable 'Trade' Permission on API Key" : "Check Logs";

    sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Suggested Action', '', '', 'Updated']]);
    sheet.getRange(2, 1, 1, 6).setValues([['Error', errorMsg, action, '', '', new Date()]]);
    sheet.getRange(2, 1, 1, 3).setFontColor('red');

    SyncManager.log("WARNING", `OKX Loan Error Displayed: ${errorMsg}`, "Sync_Okx");
  }
}

// ... fetchOkxApi_ (保持不變) ...
function fetchOkxApi_(baseUrl, endpoint, params, apiKey, apiSecret, apiPassphrase) {
  const method = 'GET';
  const timestamp = new Date().toISOString(); // ISO 8601 for OKX
  let qs = '';
  if (params && Object.keys(params).length) {
    qs = '?' + Object.keys(params).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
  }
  const url = baseUrl + endpoint + qs;

  // 修正簽名邏輯: 必須嚴格按照 (timestamp + method + requestPath + body)
  // requestPath 必須包含 query string (?a=b)
  const signStr = timestamp + method + endpoint + qs;

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