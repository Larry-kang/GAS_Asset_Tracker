// =======================================================
// --- Binance Service Layer (v3.0 - Unified Ledger) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getBinanceBalance() {
  const MODULE_NAME = "Sync_Binance";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const creds = Credentials.get('BINANCE');
    const { apiKey, apiSecret, tunnelUrl: baseUrl, proxyPassword } = creds;

    if (!Credentials.isValid(creds)) {
      SyncManager.log("ERROR", "Missing BINANCE_API_KEY or SECRET", MODULE_NAME);
      return;
    }
    if (!baseUrl || !proxyPassword) {
      SyncManager.log("WARNING", "Skipping: No Tunnel URL", MODULE_NAME);
      return;
    }

    // --- Data Collection ---
    const assetList = [];

    // A. Spot
    const spotRes = fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (spotRes.success && spotRes.data) {
      spotRes.data.forEach(item => {
        // [Rule] Exclude 'LD' (Locked/Earn) assets from Spot to prevent duplication/clutter
        if (item.asset.startsWith('LD')) return;

        // Free
        if (item.free > 0) {
          assetList.push({ ccy: item.asset, amt: item.free, type: 'Spot', status: 'Available' });
        }
        // Locked
        if (item.locked > 0) {
          assetList.push({ ccy: item.asset, amt: item.locked, type: 'Spot', status: 'Frozen' });
        }
      });
    }

    // B. Earn
    const earnRes = fetchEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (earnRes.success && earnRes.data) {
      earnRes.data.forEach(item => {
        assetList.push({ ccy: item.asset, amt: item.amount, type: 'Earn', status: 'Staked' });
      });
    }

    // C. Loans
    const loanRes = fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (loanRes.success && loanRes.data) {
      loanRes.data.forEach(order => {
        // Debt (Negative)
        assetList.push({
          ccy: order.loanCoin,
          amt: -Math.abs(order.totalDebt),
          type: 'Loan',
          status: 'Debt',
          meta: `LTV: ${(order.currentLTV * 100).toFixed(2)}%`
        });
        // Collateral
        assetList.push({
          ccy: order.collateralCoin,
          amt: order.collateralAmount,
          type: 'Loan',
          status: 'Collateral',
          meta: `Ref: ${order.loanCoin}`
        });
      });
    }

    // --- Write to Unified Ledger ---
    SyncManager.log("INFO", `Collected ${assetList.length} asset entries. Updating Ledger...`, MODULE_NAME);
    SyncManager.updateUnifiedLedger(ss, "Binance", assetList);

    // Clean up old sheets if needed (manual cleanup recommended later)
  });
}

// --- Helpers (Modified to return raw lists instead of Map) ---

function fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/api/v3/account', {}, apiKey, apiSecret, proxyPassword);
  if (res.code !== "0") return { success: false };

  // Return Raw Array: [{asset, free, locked}]
  const rawList = [];
  if (res.data && res.data.balances) {
    res.data.balances.forEach(b => {
      const free = parseFloat(b.free);
      const locked = parseFloat(b.locked);
      if (free + locked > 0) {
        rawList.push({ asset: b.asset, free, locked });
      }
    });
  }
  return { success: true, data: rawList };
}

function fetchEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/sapi/v1/simple-earn/flexible/position', { limit: 100 }, apiKey, apiSecret, proxyPassword);
  if (res.code !== "0") return { success: false };

  const rawList = [];
  const rows = Array.isArray(res.data) ? res.data : (res.data.rows || []);
  rows.forEach(r => {
    if (parseFloat(r.totalAmount) > 0) {
      rawList.push({ asset: r.asset, amount: parseFloat(r.totalAmount) });
    }
  });
  return { success: true, data: rawList };
}

function fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/sapi/v2/loan/flexible/ongoing/orders', { limit: 100 }, apiKey, apiSecret, proxyPassword);
  if (res.code !== "0") return { success: false };

  const rawList = [];
  const rows = Array.isArray(res.data) ? res.data : (res.data.rows || []);
  rows.forEach(r => {
    rawList.push({
      loanCoin: r.loanCoin,
      totalDebt: parseFloat(r.totalDebt),
      collateralCoin: r.collateralCoin,
      collateralAmount: parseFloat(r.collateralAmount),
      currentLTV: parseFloat(r.currentLTV)
    });
  });
  return { success: true, data: rawList };
}

// ... fetchBinanceApi_ (Keep existing logic) ...
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
    'headers': { 'X-MBX-APIKEY': apiKey, 'Content-Type': 'application/json', 'x-proxy-auth': proxyPassword },
    'muteHttpExceptions': true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 403) return { code: "-403", msg: "Proxy Auth Failed" };
    return { code: "0", data: JSON.parse(response.getContentText()) };
  } catch (e) { return { code: "-1", msg: e.message }; }
}
