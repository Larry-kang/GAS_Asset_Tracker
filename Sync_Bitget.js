// =======================================================
// --- Bitget Service Layer (v1.0 - Unified Ledger)
// --- Supports Spot, Savings, Crypto Loan, and optional Funding assets
// =======================================================

function getBitgetBalance() {
  const MODULE_NAME = "Sync_Bitget";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const creds = Credentials.get('BITGET');
    const { apiKey, apiSecret, apiPassphrase } = creds;
    const baseUrl = 'https://api.bitget.com';

    if (!apiKey || !apiSecret || !apiPassphrase) {
      SyncManager.log("ERROR", "Missing BITGET_API_KEY / BITGET_API_SECRET / BITGET_API_PASSPHRASE", MODULE_NAME);
      return;
    }

    const assetList = [];
    const requiredFailures = [];

    // A. Spot
    const spotRes = fetchBitgetSpotAssets_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (spotRes.success && spotRes.data) {
      spotRes.data.forEach(item => {
        const coin = String(item.coin || "").toUpperCase();
        const available = parseBitgetNumber_(item.available);
        const frozen = parseBitgetNumber_(item.frozen);
        const locked = parseBitgetNumber_(item.locked);
        const restricted = parseBitgetNumber_(item.limitAvailable);

        if (!coin) return;
        if (available > 0) assetList.push({ ccy: coin, amt: available, type: 'Spot', status: 'Available', meta: 'Spot Available' });
        if (frozen > 0) assetList.push({ ccy: coin, amt: frozen, type: 'Spot', status: 'Frozen', meta: 'Spot Frozen' });
        if (locked > 0) assetList.push({ ccy: coin, amt: locked, type: 'Spot', status: 'Locked', meta: 'Spot Locked' });
        if (restricted > 0) assetList.push({ ccy: coin, amt: restricted, type: 'Spot', status: 'Restricted', meta: 'Spot LimitAvailable' });
      });
    } else {
      requiredFailures.push(`Spot assets fetch failed: ${spotRes.status || 'Unknown error'}`);
    }

    // B. Savings - Flexible
    const flexibleRes = fetchBitgetSavingsAssets_(baseUrl, apiKey, apiSecret, apiPassphrase, 'flexible');
    if (flexibleRes.success && flexibleRes.data) {
      flexibleRes.data.forEach(item => {
        const coin = String(item.productCoin || "").toUpperCase();
        const holdAmount = parseBitgetNumber_(item.holdAmount);
        if (!coin || holdAmount <= 0) return;

        assetList.push({
          ccy: coin,
          amt: holdAmount,
          type: 'Earn',
          status: 'Flexible',
          meta: buildBitgetSavingsMeta_(item)
        });
      });
    } else {
      requiredFailures.push(`Flexible savings fetch failed: ${flexibleRes.status || 'Unknown error'}`);
    }

    // C. Savings - Fixed
    const fixedRes = fetchBitgetSavingsAssets_(baseUrl, apiKey, apiSecret, apiPassphrase, 'fixed');
    if (fixedRes.success && fixedRes.data) {
      fixedRes.data.forEach(item => {
        const coin = String(item.productCoin || "").toUpperCase();
        const holdAmount = parseBitgetNumber_(item.holdAmount);
        if (!coin || holdAmount <= 0) return;

        assetList.push({
          ccy: coin,
          amt: holdAmount,
          type: 'Earn',
          status: 'Fixed',
          meta: buildBitgetSavingsMeta_(item)
        });
      });
    } else {
      requiredFailures.push(`Fixed savings fetch failed: ${fixedRes.status || 'Unknown error'}`);
    }

    // D. Crypto Loan - Ongoing
    const loanRes = fetchBitgetBorrowOngoing_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (loanRes.success && loanRes.data) {
      loanRes.data.forEach(item => {
        const loanCoin = String(item.loanCoin || "").toUpperCase();
        const pledgeCoin = String(item.pledgeCoin || "").toUpperCase();
        const loanAmount = parseBitgetNumber_(item.loanAmount);
        const interestAmount = parseBitgetNumber_(item.interestAmount);
        const totalDebt = loanAmount + interestAmount;
        const pledgeAmount = parseBitgetNumber_(item.pledgeAmount);

        if (loanCoin && totalDebt > 0) {
          assetList.push({
            ccy: loanCoin,
            amt: -Math.abs(totalDebt),
            type: 'Loan',
            status: 'Debt',
            meta: buildBitgetLoanDebtMeta_(item, loanAmount, interestAmount)
          });
        }

        if (pledgeCoin && pledgeAmount > 0) {
          assetList.push({
            ccy: pledgeCoin,
            amt: pledgeAmount,
            type: 'Loan',
            status: 'Collateral',
            meta: buildBitgetLoanCollateralMeta_(item)
          });
        }
      });
    } else {
      requiredFailures.push(`Crypto loan fetch failed: ${loanRes.status || 'Unknown error'}`);
    }

    // E. Funding (Optional)
    const fundingRes = fetchBitgetFundingAssets_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (fundingRes.success && fundingRes.data) {
      fundingRes.data.forEach(item => {
        const coin = String(item.coin || "").toUpperCase();
        const available = parseBitgetNumber_(item.available);
        const frozen = parseBitgetNumber_(item.frozen);
        const usdtValue = parseBitgetNumber_(item.usdtValue);
        const meta = usdtValue > 0 ? `usdtValue=${usdtValue}` : '';

        if (!coin) return;
        if (available > 0) assetList.push({ ccy: coin, amt: available, type: 'Funding', status: 'Available', meta: meta });
        if (frozen > 0) assetList.push({ ccy: coin, amt: frozen, type: 'Funding', status: 'Frozen', meta: meta });
      });
    } else {
      SyncManager.log("WARNING", `Funding assets skipped: ${fundingRes.status || 'Unknown error'}`, MODULE_NAME);
    }

    if (requiredFailures.length > 0) {
      requiredFailures.forEach(msg => SyncManager.log("ERROR", msg, MODULE_NAME));
      SyncManager.log("ERROR", "Required Bitget data incomplete. Ledger update aborted to avoid partial overwrite.", MODULE_NAME);
      return;
    }

    SyncManager.log("INFO", `Collected ${assetList.length} asset entries. Updating Ledger...`, MODULE_NAME);
    SyncManager.updateUnifiedLedger(ss, "Bitget", assetList);
  });
}

function fetchBitgetSpotAssets_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchBitgetApi_(baseUrl, '/api/v2/spot/account/assets', { assetType: 'all' }, apiKey, apiSecret, apiPassphrase);
  if (res.code === "00000" && Array.isArray(res.data)) {
    return { success: true, data: res.data };
  }
  return { success: false, status: bitgetStatus_(res) };
}

function fetchBitgetSavingsAssets_(baseUrl, apiKey, apiSecret, apiPassphrase, periodType) {
  const allRows = [];
  let cursor = '';
  let page = 0;

  while (page < 10) {
    const params = { periodType: periodType, limit: '100' };
    if (cursor) params.idLessThan = cursor;

    const res = fetchBitgetApi_(baseUrl, '/api/v2/earn/savings/assets', params, apiKey, apiSecret, apiPassphrase);
    if (res.code !== "00000") {
      return { success: false, status: bitgetStatus_(res) };
    }

    const payload = res.data || {};
    const rows = Array.isArray(payload.resultList) ? payload.resultList : [];
    allRows.push.apply(allRows, rows);

    cursor = payload.endId || '';
    if (rows.length < 100 || !cursor) break;
    page++;
  }

  return { success: true, data: allRows };
}

function fetchBitgetFundingAssets_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const res = fetchBitgetApi_(baseUrl, '/api/v2/account/funding-assets', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "00000" && Array.isArray(res.data)) {
    return { success: true, data: res.data };
  }
  return { success: false, status: bitgetStatus_(res) };
}

function fetchBitgetBorrowOngoing_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  // Priority 1: UTA crypto loan endpoint
  let res = fetchBitgetApi_(baseUrl, '/api/v3/loan/borrow-ongoing', {}, apiKey, apiSecret, apiPassphrase);
  if (res.code === "00000" && Array.isArray(res.data)) {
    return { success: true, data: res.data };
  }

  // Fallback: Classic account crypto loan endpoint
  if (res.code === "40084") {
    res = fetchBitgetApi_(baseUrl, '/api/v2/earn/loan/ongoing-orders', {}, apiKey, apiSecret, apiPassphrase);
    if (res.code === "00000" && Array.isArray(res.data)) {
      return { success: true, data: res.data };
    }
  }
  return { success: false, status: bitgetStatus_(res) };
}

function fetchBitgetApi_(baseUrl, endpoint, params, apiKey, apiSecret, apiPassphrase, method = 'GET', body = null) {
  const upperMethod = method.toUpperCase();
  const timestamp = String(Date.now());
  const queryString = buildBitgetQueryString_(params || {});
  const requestPath = queryString ? `${endpoint}?${queryString}` : endpoint;
  const bodyString = body ? JSON.stringify(body) : '';
  const signPayload = timestamp + upperMethod + requestPath + bodyString;
  const signatureBytes = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, signPayload, apiSecret);
  const signature = Utilities.base64Encode(signatureBytes);
  const url = baseUrl + requestPath;

  const options = {
    method: upperMethod,
    headers: {
      'ACCESS-KEY': apiKey,
      'ACCESS-SIGN': signature,
      'ACCESS-TIMESTAMP': timestamp,
      'ACCESS-PASSPHRASE': apiPassphrase,
      'locale': 'en-US',
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (bodyString) {
    options.payload = bodyString;
  }

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const content = response.getContentText();
    const json = content ? JSON.parse(content) : {};

    if (!json.code) {
      json.code = responseCode.toString();
      json.msg = json.msg || json.message || content;
    }

    return json;
  } catch (e) {
    return { code: "-1", msg: e.message };
  }
}

function buildBitgetQueryString_(params) {
  const keys = Object.keys(params).filter(key => params[key] !== null && params[key] !== undefined && params[key] !== '');
  keys.sort();
  return keys.map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`).join('&');
}

function parseBitgetNumber_(value) {
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

function bitgetStatus_(res) {
  return `Code: ${res.code || 'UNKNOWN'}, Msg: ${res.msg || res.message || 'Unknown error'}`;
}

function buildBitgetSavingsMeta_(item) {
  const metaParts = [];
  const apy = getBitgetCurrentApy_(item.apy);
  const totalProfit = parseBitgetNumber_(item.totalProfit);

  if (item.interestCoin) metaParts.push(`interestCoin=${item.interestCoin}`);
  if (item.status) metaParts.push(`status=${item.status}`);
  if (item.productLevel) metaParts.push(`level=${item.productLevel}`);
  if (apy) metaParts.push(`apy=${apy}`);
  if (totalProfit !== 0) metaParts.push(`totalProfit=${totalProfit}`);
  if (item.periodType === 'fixed' && item.period) metaParts.push(`period=${item.period}`);
  if (item.holdDays) metaParts.push(`holdDays=${item.holdDays}`);

  return metaParts.join('; ');
}

function getBitgetCurrentApy_(apyList) {
  if (!Array.isArray(apyList) || apyList.length === 0) return '';
  const currentApy = apyList[0] && apyList[0].currentApy;
  return currentApy ? String(currentApy) : '';
}

function buildBitgetLoanDebtMeta_(item, loanAmount, interestAmount) {
  const metaParts = [];
  const pledgeRate = parseBitgetNumber_(item.pledgeRate);
  const supRate = parseBitgetNumber_(item.supRate);
  const forceRate = parseBitgetNumber_(item.forceRate);

  if (item.orderId) metaParts.push(`orderId=${item.orderId}`);
  if (loanAmount > 0) metaParts.push(`principal=${loanAmount}`);
  if (interestAmount > 0) metaParts.push(`interest=${interestAmount}`);
  if (item.pledgeCoin) metaParts.push(`pledgeCoin=${item.pledgeCoin}`);
  if (pledgeRate > 0) metaParts.push(`pledgeRate=${pledgeRate}`);
  if (supRate > 0) metaParts.push(`supRate=${supRate}`);
  if (forceRate > 0) metaParts.push(`forceRate=${forceRate}`);

  return metaParts.join('; ');
}

function buildBitgetLoanCollateralMeta_(item) {
  const metaParts = [];
  const pledgeRate = parseBitgetNumber_(item.pledgeRate);
  const supRate = parseBitgetNumber_(item.supRate);
  const forceRate = parseBitgetNumber_(item.forceRate);

  if (item.orderId) metaParts.push(`orderId=${item.orderId}`);
  if (item.loanCoin) metaParts.push(`loanCoin=${item.loanCoin}`);
  if (item.loanAmount) metaParts.push(`loanAmount=${item.loanAmount}`);
  if (item.interestAmount) metaParts.push(`interest=${item.interestAmount}`);
  if (pledgeRate > 0) metaParts.push(`pledgeRate=${pledgeRate}`);
  if (supRate > 0) metaParts.push(`supRate=${supRate}`);
  if (forceRate > 0) metaParts.push(`forceRate=${forceRate}`);

  return metaParts.join('; ');
}
