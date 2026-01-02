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
    // 無論成功與否都呼叫，以便在失敗時寫入錯誤提示
    updateLoanSheet_(ss, loanResult);

    const finalMsg = `OKX 同步完成！`;
    ss.toast(finalMsg);

  } catch (e) {
    Logger.log(`Critical Crash: ${e.message}`);
    ss.toast(`Error: ${e.message}`);
  }
}

// ... fetchOkxAccount_ (MODIFIED: Net Balance = Avail + Frozen - Liab) ...
function fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
  const balances = new Map();

  if (res.code === "0" && res.data && res.data[0] && res.data[0].details) {
    res.data[0].details.forEach(b => {
      // 淨餘額 = 可用 + 凍結 - 負債 (liab)
      // 若有負債，liab 會是正數，需減除
      const total = (parseFloat(b.availBal) || 0) + (parseFloat(b.frozenBal) || 0) - (parseFloat(b.liab) || 0);

      // 只要不等於 0 都要紀錄 (包含負值)
      if (Math.abs(total) > 1e-8) {
        balances.set(b.ccy, total);
      }
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
/**
 * [模組 C] 獲取 Flexible Loan 訂單
 * 注意：此 endpoint (orders) 雖然文件標示只需 Read 權限，但部分帳戶可能因風控需開啟 Trade 權限。
 * 若回傳 Forbidden，這就是唯一解法。
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
    console.log(`[FetchLoans] Error: ${errorMsg}`);

    // 提供明確的 Debug 提示
    let statusHint = errorMsg;
    if (res.code == "50100" || (res.msg && res.msg.toLowerCase().includes("forbidden"))) {
      statusHint = "Forbidden (Please try enabling 'Trade' permission on API Key)";
    }

    return { success: false, status: statusHint, data: [] };
  }
}

/**
 * [工具] 更新資產總表
 */
function updateBalanceSheet_(ss, spotData, earnData) {
  let sheet = ss.getSheetByName('OKX Balance');
  if (!sheet) sheet = ss.insertSheet('OKX Balance');

  // 1. 強制寫入標題
  sheet.getRange('A1:C1').setValues([['Currency', 'Total', 'Last Updated']]);
  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#e6f7ff');

  const combined = new Map();

  const mergeData = (data) => {
    if (!data) return;
    if (data instanceof Map) {
      data.forEach((val, key) => {
        const current = combined.get(key) || 0;
        combined.set(key, current + val);
      });
    } else if (typeof data === 'object') {
      Object.keys(data).forEach(key => {
        const val = data[key];
        const current = combined.get(key) || 0;
        combined.set(key, current + val);
      });
    }
  };

  mergeData(spotData);
  mergeData(earnData);

  const rows = [];
  combined.forEach((v, k) => { if (v > 0) rows.push([k, v]); });
  rows.sort((a, b) => b[1] - a[1]); // Descending Sort

  sheet.getRange('A2:C').clearContent();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    sheet.getRange(2, 3).setValue(timestamp);
  } else {
    // 顯示無資產
    sheet.getRange(2, 1, 1, 3).setValues([['No Assets', 0, timestamp]]);
  }
}

// ... updateLoanSheet_ (保持不變) ...
// ... updateLoanSheet_ (MODIFIED: 支援錯誤顯示) ...
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
    // 確保表頭回到正常模式 (如果之前變成錯誤顯示模式)
    if (sheet.getRange(1, 2).getValue() !== 'Debt') {
      sheet.getRange(1, 1, 1, 6).setValues([['Loan Coin', 'Debt', 'Collateral Coin', 'Collateral Amt', 'LTV', 'Updated']]);
    }
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  } else if (!loanResult.success) {
    // 顯示錯誤資訊
    const errorMsg = loanResult.status;
    const action = errorMsg.includes('Forbidden') ? "Enable 'Trade' Permission on API Key" : "Check Logs";
    // 更改表頭以適應錯誤訊息
    sheet.getRange(1, 1, 1, 6).setValues([['Status', 'Details', 'Suggested Action', '', '', 'Updated']]);
    sheet.getRange(2, 1, 1, 6).setValues([['Error', errorMsg, action, '', '', new Date()]]);
    // 紅色警示
    sheet.getRange(2, 1, 1, 3).setFontColor('red');
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