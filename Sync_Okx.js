// =======================================================
// --- OKX Service Layer (v3.1 - Debugging) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getOkxBalance() {
  const MODULE_NAME = "Sync_Okx";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = SyncManager.createResult("OKX");
    const creds = Credentials.get('OKX');
    const { apiKey, apiSecret, apiPassphrase } = creds;
    const baseUrl = 'https://www.okx.com';

    if (!apiKey || !apiSecret || !apiPassphrase) {
      SyncManager.registerSourceCheck(result, {
        name: 'Credentials',
        required: true,
        success: false,
        message: 'Missing OKX Keys'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }

    // A. Account (Spot + Funding + Margin Debt)
    const accRes = fetchOkxAccount_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let accountRows = 0;
    if (accRes.success && accRes.data) {
      accRes.data.forEach(item => {
        const accountMeta = buildOkxAccountMeta_(item);
        if (item.avail > 0) {
          result.assets.push({ ccy: item.ccy, amt: item.avail, type: 'Spot', status: 'Available', meta: accountMeta });
          accountRows++;
        }
        if (item.frozen > 0) {
          result.assets.push({ ccy: item.ccy, amt: item.frozen, type: 'Spot', status: 'Frozen', meta: accountMeta });
          accountRows++;
        }
        if (item.liab > 0) {
          result.assets.push({
            ccy: item.ccy,
            amt: -Math.abs(item.liab),
            type: 'Loan',
            status: 'Debt',
            meta: appendOkxMetaPrefix_('Margin Liability', accountMeta)
          });
          accountRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'Account', required: true, success: true, rows: accountRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Account',
        required: true,
        success: false,
        message: accRes.status || 'Unknown Error'
      });
    }

    // B. Simple Earn
    const earnRes = fetchOkxEarn_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let earnRows = 0;
    if (earnRes.success && earnRes.data) {
      earnRes.data.forEach(item => {
        result.assets.push({ ccy: item.ccy, amt: item.amt, type: 'Earn', status: 'Staked' });
        earnRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Earn', required: true, success: true, rows: earnRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Earn',
        required: true,
        success: false,
        message: earnRes.status || 'Unknown Error'
      });
    }

    // C. Flexible Loan Orders
    const loanRes = fetchOkxLoans_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let loanRows = 0;
    if (loanRes.success && loanRes.data) {
      loanRes.data.forEach(item => {
        if (item.type === 'Debt') {
          result.assets.push({
            ccy: item.loanCoin,
            amt: -Math.abs(item.totalDebt),
            type: 'Loan',
            status: 'Debt',
            meta: `Flex Loan (LTV: ${(item.currentLTV * 100).toFixed(2)}%)`
          });
          loanRows++;
        }
        else if (item.type === 'Collateral') {
          result.assets.push({
            ccy: item.collateralCoin,
            amt: item.collateralAmount,
            type: 'Loan',
            status: 'Collateral',
            meta: `Flex Collateral`
          });
          loanRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'Loans', required: true, success: true, rows: loanRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Loans',
        required: true,
        success: false,
        message: loanRes.status || 'Unknown Error'
      });
    }

    // D. Positions (Futures/Swap)
    const posRes = fetchOkxPositions_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let positionRows = 0;
    if (posRes.success && posRes.data) {
      posRes.data.forEach(item => {
        result.assets.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Positions',
          status: 'Position',
          meta: buildOkxPositionMeta_(item)
        });
        positionRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Positions', required: false, success: true, rows: positionRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Positions',
        required: false,
        success: false,
        message: posRes.status || 'Unknown Error'
      });
    }

    // E. Funding Assets
    const fundRes = fetchOkxAssets_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let fundingRows = 0;
    if (fundRes.success && fundRes.data) {
      fundRes.data.forEach(item => {
        result.assets.push({ ccy: item.ccy, amt: item.amt, type: 'Funding', status: 'Available' });
        fundingRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Funding', required: true, success: true, rows: fundingRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Funding',
        required: true,
        success: false,
        message: fundRes.status || 'Unknown Error'
      });
    }

    // F. Structured Products (Shark Fin, etc.)
    const structRes = fetchOkxStructured_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let structuredRows = 0;
    if (structRes.success && structRes.data) {
      structRes.data.forEach(item => {
        result.assets.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Structured',
          status: 'Active',
          meta: `Structured ID: ${item.ordId}`
        });
        structuredRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Structured', required: false, success: true, rows: structuredRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Structured',
        required: false,
        success: false,
        message: structRes.status || 'Unknown Error'
      });
    }

    SyncManager.log("INFO", `Collected ${result.assets.length} asset entries from OKX.`, MODULE_NAME);
    return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
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
        liab: parseFloat(b.liab) || 0,
        eqUsd: b.eqUsd,
        cashBal: b.cashBal,
        upl: b.upl,
        uplLiab: b.uplLiab,
        interest: b.interest,
        imr: b.imr,
        mmr: b.mmr,
        maxLoan: b.maxLoan,
        spotInUseAmt: b.spotInUseAmt || b.clSpotInUseAmt,
        borrowFroz: b.borrowFroz
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
            mgnMode: p.mgnMode,
            posSide: p.posSide,
            avgPx: p.avgPx,
            markPx: p.markPx,
            upl: p.upl,
            liqPx: p.liqPx,
            lever: p.lever,
            notionalUsd: p.notionalUsd,
            adl: p.adl,
            uTime: p.uTime
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

function buildOkxAccountMeta_(item) {
  const parts = [];

  pushOkxMetaPart_(parts, 'eqUsd', item.eqUsd);
  pushOkxMetaPart_(parts, 'cashBal', item.cashBal);
  pushOkxMetaPart_(parts, 'upl', item.upl);
  pushOkxMetaPart_(parts, 'uplLiab', item.uplLiab);
  pushOkxMetaPart_(parts, 'interest', item.interest);
  pushOkxMetaPart_(parts, 'imr', item.imr);
  pushOkxMetaPart_(parts, 'mmr', item.mmr);
  pushOkxMetaPart_(parts, 'maxLoan', item.maxLoan);
  pushOkxMetaPart_(parts, 'spotInUseAmt', item.spotInUseAmt);
  pushOkxMetaPart_(parts, 'borrowFroz', item.borrowFroz);

  return parts.join('; ');
}

function buildOkxPositionMeta_(item) {
  const parts = [];

  pushOkxMetaPart_(parts, 'instId', item.instId);
  pushOkxMetaPart_(parts, 'posSide', item.posSide);
  pushOkxMetaPart_(parts, 'mgnMode', item.mgnMode);
  pushOkxMetaPart_(parts, 'avgPx', item.avgPx);
  pushOkxMetaPart_(parts, 'markPx', item.markPx);
  pushOkxMetaPart_(parts, 'upl', item.upl);
  pushOkxMetaPart_(parts, 'liqPx', item.liqPx);
  pushOkxMetaPart_(parts, 'lever', item.lever);
  pushOkxMetaPart_(parts, 'notionalUsd', item.notionalUsd);
  pushOkxMetaPart_(parts, 'adl', item.adl);
  pushOkxMetaPart_(parts, 'uTime', item.uTime);

  return parts.join('; ');
}

function pushOkxMetaPart_(parts, key, value) {
  if (value === null || value === undefined) return;
  const text = String(value).trim();
  if (!text || text === '0' || text === '0.0' || text === '0.00' || text === '0.00000000') return;
  parts.push(`${key}=${text}`);
}

function appendOkxMetaPrefix_(prefix, meta) {
  if (!meta) return prefix;
  return `${prefix}; ${meta}`;
}
