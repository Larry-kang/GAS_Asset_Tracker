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

    // --- 步驟 A: 獲取帳戶餘額 (Spot/Funding) & 負債 (Liabilities) ---
    ss.toast('正在獲取 OKX 帳戶與負債 (1/2)...');
    const accountResult = fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Account: ${accountResult.status}`);
    if (!accountResult.success) Logger.log(`Account Fetch Error: ${accountResult.status}`);

    // --- 步驟 B: 獲取理財餘額 (Simple Earn) ---
    ss.toast('正在獲取 OKX 理財餘額 (2/2)...');
    const earnResult = fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase);
    logMessages.push(`Earn: ${earnResult.status}`);
    if (!earnResult.success) Logger.log(`Earn Fetch Error: ${earnResult.status}`);

    // --- (已停用) 步驟 C: 獲取借貸訂單 (Flexible Loans) ---
    // 因 User 回報權限 Forbidden (403)，改由從 Account Balance 直接讀取 "liab" (負債)
    // const loanResult = fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase);

    // Console log for debugging
    console.log("===== OKX Update Log =====");
    console.log(logMessages.join("\n"));

    // --- 步驟 D: 整合並寫入資產表 (Balance Sheet) ---
    let balanceUpdated = false;
    // accountResult.data is now { balances: Map, debts: Map }
    const spotMap = accountResult.data ? accountResult.data.balances : new Map();
    const debtMap = accountResult.data ? accountResult.data.debts : new Map();
    const earnMap = earnResult.data || new Map();

    if (accountResult.success || earnResult.success) {
      updateBalanceSheet_(ss, spotMap, earnMap);
      balanceUpdated = true;
    } else {
      console.log("⚠️ Account 與 Earn 皆未更新，跳過寫入 'OKX Balance' 工作表");
    }

    // --- 步驟 E: 寫入借貸表 (Loan Sheet) ---
    // 使用從 Account Balance 抓到的負債資料
    let loansUpdated = false;
    if (debtMap.size > 0) {
      // 轉換格式為 updateLoanSheet_ 接受的陣列
      // 由於 Account Balance 只提供負債總額，無法區分對應的質押品，故 Collateral 顯示為 'Unified/Unknown'
      const simulatedLoanOrders = [];
      debtMap.forEach((amount, ccy) => {
        simulatedLoanOrders.push({
          loanCoin: ccy,
          totalDebt: amount,
          collateralCoin: 'Unified',
          collateralAmount: 0,
          currentLTV: 0 // 無法計算單一 LTV
        });
      });
      updateLoanSheet_(ss, simulatedLoanOrders);
      loansUpdated = true;
    } else {
      // 如果沒有負債，也要更新空表 (或清除)
      // 若原表有資料但現在沒負債，這裡選擇清除內容
      // updateLoanSheet_(ss, []); // 視需求決定是否清空，目前保持若無資料則不動
      if (accountResult.success) {
        // 若成功讀取且無負債，嘗試清空舊資料 (判定為已還款)
        // 但為了安全，暫不清空，僅 Log
        Logger.log("No liabilities detected.");
      }
    }

    // --- 最終通知 ---
    const finalMsg = `OKX 同步完成！\n餘額: ${balanceUpdated ? '更新' : '跳過'}, 借貸: ${loansUpdated ? '更新' : '無需更新'}`;
    ss.toast(finalMsg);
    Logger.log(finalMsg);

  } catch (e) {
    Logger.log(`Critical Crash in Sync_Okx: ${e.message}`);
    ss.toast(`OKX Sync Error: ${e.message}`);
    console.error(e);
  }
}

/**
 * [模組 A] 獲取 Account 餘額 (含資金與交易帳戶) + 負債偵測
 * API: /api/v5/account/balance
 * return: { success, status, data: { balances: Map, debts: Map } }
 */
function fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
  const balances = new Map();
  const debts = new Map();

  if (res.code === "0" && res.data && res.data[0] && res.data[0].details) {
    res.data[0].details.forEach(b => {
      // 1. 資產運算
      const total = (parseFloat(b.availBal) || 0) + (parseFloat(b.frozenBal) || 0);
      if (total > 0) {
        balances.set(b.ccy, total);
      }

      // 2. 負債運算 (Extraction from liability field)
      // liab: 幣種負債
      const liability = parseFloat(b.liab) || 0;
      if (liability > 0) {
        debts.set(b.ccy, liability);
      }
    });
    return { success: true, status: "OK", data: { balances: balances, debts: debts } };
  } else {
    return { success: false, status: `Failed (${res.msg || 'No Data'})`, data: { balances: balances, debts: debts } };
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
 * [模組 C] (已停用) 獲取 Flexible Loan 訂單
 * 僅保留代碼供參考，目前因權限問題不呼叫
 */
function fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  return { success: false, status: "Skipped (Forbidden)", data: [] };
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
 * [私有工具] OKX API 連線核心 (維持不變)
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