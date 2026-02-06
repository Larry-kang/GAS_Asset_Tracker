// =======================================================
// --- OKX Service Layer (v3.1 - Debugging) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getOkxBalance() {
  const MODULE_NAME = "Sync_Okx";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const creds = Credentials.get('OKX');
    const { apiKey, apiSecret, apiPassphrase } = creds;
    const baseUrl = 'https://www.okx.com';

    if (!apiKey || !apiSecret || !apiPassphrase) {
      SyncManager.log("ERROR", "Missing OKX Keys", MODULE_NAME);
      return;
    }

    const assetList = [];

    // A. Account (Spot + Funding + Margin Debt)
    const accRes = fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (accRes.success && accRes.data) {
      accRes.data.forEach(item => {
        if (item.avail > 0) assetList.push({ ccy: item.ccy, amt: item.avail, type: 'Spot', status: 'Available' });
        if (item.frozen > 0) assetList.push({ ccy: item.ccy, amt: item.frozen, type: 'Spot', status: 'Frozen' });
        if (item.liab > 0) assetList.push({ ccy: item.ccy, amt: -Math.abs(item.liab), type: 'Loan', status: 'Debt', meta: 'Margin Liability' });
      });
    } else {
      SyncManager.log("ERROR", `Failed to fetch Account: ${accRes.status || 'Unknown Error'}`, MODULE_NAME);
    }

    // B. Simple Earn
    const earnRes = fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (earnRes.success && earnRes.data) {
      earnRes.data.forEach(item => {
        assetList.push({ ccy: item.ccy, amt: item.amt, type: 'Earn', status: 'Staked' });
      });
    } else {
      // Earn might be empty or permissions issue?
      if (earnRes.status) SyncManager.log("WARNING", `Fetch Earn Failed: ${earnRes.status}`, MODULE_NAME);
    }

    // C. Flexible Loan Orders
    const loanRes = fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (loanRes.success && loanRes.data) {
      loanRes.data.forEach(item => {
        if (item.type === 'Debt') {
          assetList.push({
            ccy: item.loanCoin,
            amt: -Math.abs(item.totalDebt),
            type: 'Loan',
            status: 'Debt',
            meta: `Flex Loan (LTV: ${(item.currentLTV * 100).toFixed(2)}%)`
          });
        }
        else if (item.type === 'Collateral') {
          assetList.push({
            ccy: item.collateralCoin,
            amt: item.collateralAmount,
            type: 'Loan',
            status: 'Collateral',
            meta: `Flex Collateral`
          });
        }
      });
    } else {
      // Loan failing is common if no loans exist or permissions missing
      if (loanRes.status) SyncManager.log("WARNING", `Fetch Loans Failed: ${loanRes.status}`, MODULE_NAME);
    }

    // D. Positions (Futures/Swap)
    const posRes = fetchOkxPositions_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (posRes.success && posRes.data) {
      posRes.data.forEach(item => {
        assetList.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Positions',
          status: 'Position',
          meta: `${item.instId} (${item.mgnMode})`
        });
      });
    } else if (posRes.status) {
      SyncManager.log("WARNING", `Fetch Positions Failed: ${posRes.status}`, MODULE_NAME);
    }

    // E. Funding Assets
    const fundRes = fetchOkxAssets_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (fundRes.success && fundRes.data) {
      fundRes.data.forEach(item => {
        assetList.push({ ccy: item.ccy, amt: item.amt, type: 'Funding', status: 'Available' });
      });
    } else if (fundRes.status) {
      SyncManager.log("WARNING", `Fetch Funding Failed: ${fundRes.status}`, MODULE_NAME);
    }

    // F. Structured Products (Shark Fin, etc.)
    const structRes = fetchOkxStructured_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (structRes.success && structRes.data) {
      structRes.data.forEach(item => {
        assetList.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Structured',
          status: 'Active',
          meta: `Structured ID: ${item.ordId}`
        });
      });
    } else if (structRes.status) {
      SyncManager.log("WARNING", `Fetch Structured Failed: ${structRes.status}`, MODULE_NAME);
    }

    // --- Update Ledger ---
    SyncManager.log("INFO", `Collected ${assetList.length} asset entries. Updating Ledger...`, MODULE_NAME);
    SyncManager.updateUnifiedLedger(ss, "OKX", assetList);
  });
}

// --- Helpers ---

function fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/balance', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "0" && res.data && res.data[0] && res.data[0].details) {
    const rawList = [];
    res.data[0].details.forEach(b => {
      rawList.push({
        ccy: b.ccy,
        avail: parseFloat(b.availBal) || 0,
        frozen: parseFloat(b.frozenBal) || 0,
        liab: parseFloat(b.liab) || 0
      });
    });
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/finance/savings/balance', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "0") {
    const rawList = [];
    if (res.data) {
      res.data.forEach(b => {
        if (parseFloat(b.amt) > 0) rawList.push({ ccy: b.ccy, amt: parseFloat(b.amt) });
      });
    }
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  // Use 'loan-info' instead of 'orders' for unified flexible loan position
  const res = fetchOkxApi_(baseUrl, '/api/v5/finance/flexible-loan/loan-info', {}, apiKey, apiSecret, apiPassphrase);

  if (res.code === "0") {
    const rawList = [];
    if (res.data && Array.isArray(res.data)) {
      res.data.forEach(entry => {
        const ltv = entry.curLTV || "0";

        // 1. Debt
        if (entry.loanData) {
          entry.loanData.forEach(l => {
            rawList.push({
              loanCoin: l.ccy,
              totalDebt: parseFloat(l.amt),
              type: 'Debt',
              currentLTV: parseFloat(ltv)
            });
          });
        }

        // 2. Collateral
        if (entry.collateralData) {
          entry.collateralData.forEach(c => {
            rawList.push({
              collateralCoin: c.ccy,
              collateralAmount: parseFloat(c.amt),
              type: 'Collateral',
              currentLTV: parseFloat(ltv)
            });
          });
        }
      });
    }
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxPositions_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/account/positions', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "0") {
    const rawList = [];
    if (res.data) {
      res.data.forEach(p => {
        const pos = parseFloat(p.pos);
        if (pos !== 0) {
          rawList.push({
            ccy: p.ccy || p.instId.split('-')[0],
            amt: pos,
            instId: p.instId,
            mgnMode: p.mgnMode
          });
        }
      });
    }
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxAssets_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/asset/balances', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "0") {
    const rawList = [];
    if (res.data) {
      res.data.forEach(b => {
        const bal = parseFloat(b.bal);
        if (bal > 0) {
          rawList.push({ ccy: b.ccy, amt: bal });
        }
      });
    }
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxStructured_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchOkxApi_(baseUrl, '/api/v5/finance/staking-defi/orders-active', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "0") {
    const rawList = [];
    if (res.data) {
      res.data.forEach(o => {
        const amt = parseFloat(o.amt);
        if (amt > 0) {
          rawList.push({ ccy: o.ccy, amt: amt, ordId: o.ordId });
        }
      });
    }
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxApi_(baseUrl, endpoint, params, apiKey, apiSecret, apiPassphrase) {
  const method = 'GET';
  const timestamp = new Date().toISOString();
  let qs = '';
  if (params && Object.keys(params).length) {
    qs = '?' + Object.keys(params).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
  }
  const url = baseUrl + endpoint + qs;
  const signStr = timestamp + method + endpoint + qs;
  const signature = Utilities.base64Encode(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, signStr, apiSecret));

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
    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();
    const json = JSON.parse(content);
    return json;
  } catch (e) { return { code: "-1", msg: e.message }; }
}
