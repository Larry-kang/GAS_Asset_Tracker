// =======================================================
// --- Binance Service Layer (v3.0 - Unified Ledger) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getBinanceBalance() {
  const MODULE_NAME = "Sync_Binance";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = SyncManager.createResult("Binance");
    const creds = Credentials.get('BINANCE');
    const { apiKey, apiSecret, tunnelUrl: baseUrl, proxyPassword } = creds;

    if (!Credentials.isValid(creds)) {
      SyncManager.registerSourceCheck(result, {
        name: 'Credentials',
        required: true,
        success: false,
        message: 'Missing BINANCE_API_KEY or SECRET'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }
    if (!baseUrl || !proxyPassword) {
      SyncManager.registerSourceCheck(result, {
        name: 'Bridge',
        required: true,
        success: false,
        message: 'Missing Tunnel URL or Proxy Password'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }

    // A. Spot
    const spotRes = fetchSpotBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    let spotRows = 0;
    if (spotRes.success && spotRes.data) {
      spotRes.data.forEach(item => {
        // [Rule] Exclude 'LD' (Locked/Earn) assets from Spot to prevent duplication/clutter
        if (item.asset.startsWith('LD')) return;

        // Free
        if (item.free > 0) {
          result.assets.push({ ccy: item.asset, amt: item.free, type: 'Spot', status: 'Available' });
          spotRows++;
        }
        // Locked
        if (item.locked > 0) {
          result.assets.push({ ccy: item.asset, amt: item.locked, type: 'Spot', status: 'Frozen' });
          spotRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'Spot', required: true, success: true, rows: spotRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Spot',
        required: true,
        success: false,
        message: 'Spot fetch failed'
      });
    }

    // B. Funding (Wallet)
    const fundingRes = fetchFundingBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    let fundingRows = 0;
    if (fundingRes.success && fundingRes.data) {
      fundingRes.data.forEach(item => {
        if (item.free > 0) {
          result.assets.push({ ccy: item.asset, amt: item.free, type: 'Funding', status: 'Available' });
          fundingRows++;
        }
        if (item.locked > 0) {
          result.assets.push({ ccy: item.asset, amt: item.locked, type: 'Funding', status: 'Frozen' });
          fundingRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'Funding', required: true, success: true, rows: fundingRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Funding',
        required: true,
        success: false,
        message: 'Funding fetch failed'
      });
    }

    // C. USD-M Futures
    const usdmRes = fetchUsdmFuturesBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    let usdmRows = 0;
    if (usdmRes.success && usdmRes.data) {
      usdmRes.data.forEach(item => {
        if (item.balance > 0) {
          result.assets.push({ ccy: item.asset, amt: item.balance, type: 'USD-M Futures', status: 'Equity' });
          usdmRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'USD-M Futures', required: false, success: true, rows: usdmRows });
    } else if (usdmRes.code === "-403") {
      SyncManager.registerSourceCheck(result, {
        name: 'USD-M Futures',
        required: false,
        success: false,
        message: "Missing 'Enable Futures' API Permission. USD-M data skipped."
      });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'USD-M Futures',
        required: false,
        success: false,
        message: 'USD-M futures fetch failed'
      });
    }

    // D. COIN-M Futures
    const coinmRes = fetchCoinmFuturesBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    let coinmRows = 0;
    if (coinmRes.success && coinmRes.data) {
      coinmRes.data.forEach(item => {
        if (item.balance > 0) {
          result.assets.push({ ccy: item.asset, amt: item.balance, type: 'COIN-M Futures', status: 'Equity' });
          coinmRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'COIN-M Futures', required: false, success: true, rows: coinmRows });
    } else if (coinmRes.code === "-403") {
      SyncManager.registerSourceCheck(result, {
        name: 'COIN-M Futures',
        required: false,
        success: false,
        message: "Missing 'Enable Futures' API Permission. COIN-M data skipped."
      });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'COIN-M Futures',
        required: false,
        success: false,
        message: 'COIN-M futures fetch failed'
      });
    }

    // E. Earn - Flexible
    const flexibleEarnRes = fetchFlexibleEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword);
    let flexibleEarnRows = 0;
    if (flexibleEarnRes.success && flexibleEarnRes.data) {
      flexibleEarnRes.data.forEach(item => {
        result.assets.push({
          ccy: item.asset,
          amt: item.amount,
          type: 'Earn',
          status: 'Flexible',
          meta: buildBinanceFlexibleEarnMeta_(item)
        });
        flexibleEarnRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Earn Flexible', required: true, success: true, rows: flexibleEarnRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Earn Flexible',
        required: true,
        success: false,
        message: 'Flexible earn fetch failed'
      });
    }

    // F. Earn - Locked
    const lockedEarnRes = fetchLockedEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword);
    let lockedEarnRows = 0;
    if (lockedEarnRes.success && lockedEarnRes.data) {
      lockedEarnRes.data.forEach(item => {
        result.assets.push({
          ccy: item.asset,
          amt: item.amount,
          type: 'Earn',
          status: 'Locked',
          meta: buildBinanceLockedEarnMeta_(item)
        });
        lockedEarnRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Earn Locked', required: true, success: true, rows: lockedEarnRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Earn Locked',
        required: true,
        success: false,
        message: 'Locked earn fetch failed'
      });
    }

    // G. Earn Summary (optional verification)
    const earnSummaryRes = fetchEarnAccountSummary_(baseUrl, apiKey, apiSecret, proxyPassword);
    if (earnSummaryRes.success && earnSummaryRes.data) {
      SyncManager.registerSourceCheck(result, { name: 'Earn Summary', required: false, success: true, rows: 0 });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Earn Summary',
        required: false,
        success: false,
        message: 'Earn summary fetch failed'
      });
    }

    // H. Loans
    const loanRes = fetchLoanOrders_(baseUrl, apiKey, apiSecret, proxyPassword);
    let loanRows = 0;
    if (loanRes.success && loanRes.data) {
      loanRes.data.forEach(order => {
        // Debt (Negative)
        result.assets.push({
          ccy: order.loanCoin,
          amt: -Math.abs(order.totalDebt),
          type: 'Loan',
          status: 'Debt',
          meta: `LTV: ${(order.currentLTV * 100).toFixed(2)}%`
        });
        loanRows++;
        // Collateral
        result.assets.push({
          ccy: order.collateralCoin,
          amt: order.collateralAmount,
          type: 'Loan',
          status: 'Collateral',
          meta: `Ref: ${order.loanCoin}`
        });
        loanRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Loans', required: true, success: true, rows: loanRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Loans',
        required: true,
        success: false,
        message: 'Loan fetch failed'
      });
    }

    SyncManager.log("INFO", `Collected ${result.assets.length} asset entries from Binance.`, MODULE_NAME);
    return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);

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

function fetchFundingBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  // Funding API requires POST
  const res = fetchBinanceApi_(baseUrl, '/sapi/v1/asset/get-funding-asset', {}, apiKey, apiSecret, proxyPassword, 'POST');
  if (res.code !== "0") return { success: false };

  const rawList = [];
  if (Array.isArray(res.data)) {
    res.data.forEach(item => {
      const free = parseFloat(item.free) || 0;
      const locked = parseFloat(item.locked) || 0;
      const freeze = parseFloat(item.freeze) || 0;
      const totalLocked = locked + freeze;
      if (free + totalLocked > 0) {
        rawList.push({ asset: item.asset, free: free, locked: totalLocked });
      }
    });
  }
  return { success: true, data: rawList };
}

function fetchUsdmFuturesBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/fapi/v2/balance', {}, apiKey, apiSecret, proxyPassword);
  if (res.code !== "0") return { success: false, code: res.code };

  const rawList = [];
  if (Array.isArray(res.data)) {
    res.data.forEach(item => {
      const balance = parseFloat(item.balance);
      if (balance > 0) {
        rawList.push({ asset: item.asset, balance });
      }
    });
  }
  return { success: true, data: rawList };
}

function fetchCoinmFuturesBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/dapi/v1/balance', {}, apiKey, apiSecret, proxyPassword);
  if (res.code !== "0") return { success: false, code: res.code };

  const rawList = [];
  if (Array.isArray(res.data)) {
    res.data.forEach(item => {
      const balance = parseFloat(item.balance);
      if (balance > 0) {
        rawList.push({ asset: item.asset, balance });
      }
    });
  }
  return { success: true, data: rawList };
}

function fetchFlexibleEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const paged = fetchBinanceEarnPagedRows_(
    baseUrl,
    '/sapi/v1/simple-earn/flexible/position',
    {},
    apiKey,
    apiSecret,
    proxyPassword
  );
  if (!paged.success) return paged;

  const rawList = [];
  paged.data.forEach(r => {
    const amount = parseFloat(r.totalAmount);
    if (amount > 0) {
      rawList.push({
        asset: r.asset,
        amount: amount,
        productId: r.productId,
        latestAnnualPercentageRate: r.latestAnnualPercentageRate,
        tierAnnualPercentageRate: r.tierAnnualPercentageRate,
        yesterdayRealTimeRewards: r.yesterdayRealTimeRewards,
        cumulativeTotalRewards: r.cumulativeTotalRewards,
        autoSubscribe: r.autoSubscribe,
        canRedeem: r.canRedeem,
        collateralAmount: r.collateralAmount,
        airDropAsset: r.airDropAsset
      });
    }
  });
  return { success: true, data: rawList };
}

function fetchLockedEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const paged = fetchBinanceEarnPagedRows_(
    baseUrl,
    '/sapi/v1/simple-earn/locked/position',
    {},
    apiKey,
    apiSecret,
    proxyPassword
  );
  if (!paged.success) return paged;

  const rawList = [];
  paged.data.forEach(r => {
    const amount = parseFloat(r.amount);
    if (amount > 0) {
      rawList.push({
        asset: r.asset,
        amount: amount,
        projectId: r.projectId,
        positionId: r.positionId,
        duration: r.duration,
        APY: r.APY,
        rewardAsset: r.rewardAsset,
        rewardAmt: r.rewardAmt,
        status: r.status,
        redeemTo: r.redeemTo,
        autoSubscribe: r.autoSubscribe,
        canRedeemEarly: r.canRedeemEarly,
        canFastRedemption: r.canFastRedemption,
        purchaseTime: r.purchaseTime
      });
    }
  });
  return { success: true, data: rawList };
}

function fetchEarnAccountSummary_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBinanceApi_(baseUrl, '/sapi/v1/simple-earn/account', {}, apiKey, apiSecret, proxyPassword);
  if (res.code !== "0") return { success: false };
  return { success: true, data: res.data || {} };
}

function fetchBinanceEarnPagedRows_(baseUrl, endpoint, baseParams, apiKey, apiSecret, proxyPassword) {
  const pageSize = 100;
  const maxPages = 20;
  const allRows = [];

  for (let current = 1; current <= maxPages; current++) {
    const params = Object.assign({}, baseParams || {}, { current: current, size: pageSize });
    const res = fetchBinanceApi_(baseUrl, endpoint, params, apiKey, apiSecret, proxyPassword);
    if (res.code !== "0") return { success: false, code: res.code, msg: res.msg };

    const rows = Array.isArray(res.data) ? res.data : ((res.data && res.data.rows) || []);
    if (!rows.length) break;

    allRows.push.apply(allRows, rows);

    const total = Array.isArray(res.data) ? 0 : parseInt(res.data.total, 10) || 0;
    if (rows.length < pageSize) break;
    if (total > 0 && allRows.length >= total) break;
  }

  return { success: true, data: allRows };
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

function buildBinanceFlexibleEarnMeta_(item) {
  const parts = [];
  pushBinanceMetaPart_(parts, 'productId', item.productId);
  pushBinanceMetaPart_(parts, 'latestAPR', item.latestAnnualPercentageRate);
  pushBinanceMetaPart_(parts, 'tierAPR', normalizeBinanceTierApr_(item.tierAnnualPercentageRate));
  pushBinanceMetaPart_(parts, 'autoSubscribe', item.autoSubscribe);
  pushBinanceMetaPart_(parts, 'canRedeem', item.canRedeem);
  pushBinanceMetaPart_(parts, 'collateralAmount', item.collateralAmount);
  pushBinanceMetaPart_(parts, 'airDropAsset', item.airDropAsset);
  pushBinanceMetaPart_(parts, 'yesterdayRewards', item.yesterdayRealTimeRewards);
  pushBinanceMetaPart_(parts, 'totalRewards', item.cumulativeTotalRewards);
  return parts.join('; ');
}

function buildBinanceLockedEarnMeta_(item) {
  const parts = [];
  pushBinanceMetaPart_(parts, 'projectId', item.projectId);
  pushBinanceMetaPart_(parts, 'positionId', item.positionId);
  pushBinanceMetaPart_(parts, 'duration', item.duration);
  pushBinanceMetaPart_(parts, 'APY', item.APY);
  pushBinanceMetaPart_(parts, 'rewardAsset', item.rewardAsset);
  pushBinanceMetaPart_(parts, 'rewardAmt', item.rewardAmt);
  pushBinanceMetaPart_(parts, 'status', item.status);
  pushBinanceMetaPart_(parts, 'redeemTo', item.redeemTo);
  pushBinanceMetaPart_(parts, 'autoSubscribe', item.autoSubscribe);
  pushBinanceMetaPart_(parts, 'canRedeemEarly', item.canRedeemEarly);
  pushBinanceMetaPart_(parts, 'canFastRedemption', item.canFastRedemption);
  pushBinanceMetaPart_(parts, 'purchaseTime', item.purchaseTime);
  return parts.join('; ');
}

function normalizeBinanceTierApr_(value) {
  if (!value) return '';
  if (typeof value === 'string') return value;
  try {
    return JSON.stringify(value);
  } catch (e) {
    return '';
  }
}

function pushBinanceMetaPart_(parts, label, value) {
  if (value === null || value === undefined || value === '') return;
  parts.push(`${label}=${value}`);
}

// ... fetchBinanceApi_ (Keep existing logic) ...
function fetchBinanceApi_(baseUrl, endpoint, params, apiKey, apiSecret, proxyPassword, method = 'GET') {
  const timestamp = new Date().getTime();
  params.timestamp = timestamp;
  params.recvWindow = 10000;
  const queryString = Object.keys(params).map(key => `${key}=${encodeURIComponent(params[key])}`).join('&');
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, queryString, apiSecret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  const url = `${baseUrl}${endpoint}?${queryString}&signature=${signature}`;
  const options = {
    'method': method,
    'headers': { 'X-MBX-APIKEY': apiKey, 'Content-Type': 'application/json', 'x-proxy-auth': proxyPassword },
    'muteHttpExceptions': true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 403) return { code: "-403", msg: "Proxy or API Permission Auth Failed", endpoint: endpoint };
    if (responseCode >= 400) return { code: responseCode.toString(), msg: response.getContentText(), endpoint: endpoint };

    return { code: "0", data: JSON.parse(response.getContentText()) };
  } catch (e) { return { code: "-1", msg: e.message }; }
}
