// =======================================================
// --- OKX Service Layer (v3.1 - Debugging) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

var __okxRecurringDebugPayload = null;

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
        result.assets.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Earn',
          status: 'Staked',
          meta: buildOkxEarnMeta_(item)
        });
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

    // F. Staking / DeFi Active Orders
    const structRes = fetchOkxStructured_(baseUrl, apiKey, apiSecret, apiPassphrase);
    let structuredRows = 0;
    if (structRes.success && structRes.data) {
      structRes.data.forEach(item => {
        result.assets.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Staking/DeFi',
          status: 'Active',
          meta: buildOkxStakingMeta_(item)
        });
        structuredRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Staking/DeFi', required: false, success: true, rows: structuredRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Staking/DeFi',
        required: false,
        success: false,
        message: structRes.status || 'Unknown Error'
      });
    }

    // G. Recurring Buy (Debug only; no ledger writes yet)
    const recurringRes = fetchOkxRecurringBuyDebug_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (recurringRes.success) {
      SyncManager.registerSourceCheck(result, {
        name: 'Recurring Buy Debug',
        required: false,
        success: true,
        rows: recurringRes.rowCount || 0,
        message: recurringRes.message || ''
      });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Recurring Buy Debug',
        required: false,
        success: false,
        message: recurringRes.status || 'Unknown Error'
      });
    }

    SyncManager.log("INFO", `Collected ${result.assets.length} asset entries from OKX.`, MODULE_NAME);
    return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
  });
}

function debugOkxRecurringBuy() {
  const MODULE_NAME = "Sync_Okx";
  SyncManager.run(MODULE_NAME, () => {
    const creds = Credentials.get('OKX');
    const { apiKey, apiSecret, apiPassphrase } = creds;
    const baseUrl = 'https://www.okx.com';

    if (!apiKey || !apiSecret || !apiPassphrase) {
      SyncManager.log("ERROR", "[OKX Recurring Debug] Missing OKX Keys", MODULE_NAME);
      return false;
    }

    const recurringRes = fetchOkxRecurringBuyDebug_(baseUrl, apiKey, apiSecret, apiPassphrase);
    if (!recurringRes.success) {
      SyncManager.log("ERROR", `[OKX Recurring Debug] ${recurringRes.status || 'Unknown Error'}`, MODULE_NAME);
      return false;
    }

    SyncManager.log("INFO", `[OKX Recurring Debug] ${recurringRes.message || 'Completed'}`, MODULE_NAME);
    return true;
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
        if (parseFloat(b.amt) > 0) {
          rawList.push({
            ccy: b.ccy,
            amt: parseFloat(b.amt),
            rate: b.rate,
            loanAmt: b.loanAmt,
            pendingAmt: b.pendingAmt,
            earnings: b.earnings,
            redemptAmt: b.redemptAmt
          });
        }
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
          rawList.push({
            ccy: o.ccy,
            amt: amt,
            ordId: o.ordId,
            productId: o.productId,
            protocol: o.protocol,
            protocolType: o.protocolType,
            term: o.term,
            apy: o.apy,
            state: o.state,
            purchasedTime: o.purchasedTime,
            estSettlementTime: o.estSettlementTime
          });
        }
      });
    }
    return { success: true, data: rawList };
  }
  return { success: false, status: `Code: ${res.code}, Msg: ${res.msg}` };
}

function fetchOkxRecurringBuyDebug_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const pendingRes = fetchOkxApi_(baseUrl, '/api/v5/tradingBot/recurring/orders-algo-pending', {}, apiKey, apiSecret, apiPassphrase);
  if (pendingRes.code !== "0") {
    if (String(pendingRes.code) === '50120') {
      SyncManager.log(
        "WARNING",
        `[OKX Recurring Debug] Recurring endpoint permission denied. Falling back to read-only BTC spot fills. code=${pendingRes.code}`,
        "Sync_Okx"
      );
      return fetchOkxSpotBtcFillDebug_(baseUrl, apiKey, apiSecret, apiPassphrase);
    }
    return { success: false, status: `Pending Code: ${pendingRes.code}, Msg: ${pendingRes.msg}` };
  }

  const pendingList = Array.isArray(pendingRes.data) ? pendingRes.data : [];
  const firstPending = pendingList[0] || null;
  const firstAlgoId = extractOkxRecurringAlgoId_(firstPending);

  let detailsRes = null;
  let subOrdersRes = null;
  let derivedSummary = null;

  if (firstAlgoId) {
    detailsRes = fetchOkxApi_(
      baseUrl,
      '/api/v5/tradingBot/recurring/orders-algo-details',
      { algoId: firstAlgoId },
      apiKey,
      apiSecret,
      apiPassphrase
    );

    subOrdersRes = fetchOkxApi_(
      baseUrl,
      '/api/v5/tradingBot/recurring/sub-orders',
      { algoId: firstAlgoId },
      apiKey,
      apiSecret,
      apiPassphrase
    );

    derivedSummary = buildOkxRecurringDerivedSummary_(firstPending, detailsRes, subOrdersRes);
  }

  const historyRes = fetchOkxApi_(
    baseUrl,
    '/api/v5/tradingBot/recurring/orders-algo-history',
    { limit: 10 },
    apiKey,
    apiSecret,
    apiPassphrase
  );

  const preview = {
    pendingCount: pendingList.length,
    firstPendingKeys: firstPending ? Object.keys(firstPending) : [],
    firstPending: toOkxDebugPreview_(firstPending),
    firstAlgoId: firstAlgoId || '',
    detailsCode: detailsRes ? detailsRes.code : '',
    detailsKeys: detailsRes && Array.isArray(detailsRes.data) && detailsRes.data[0] ? Object.keys(detailsRes.data[0]) : [],
    detailsPreview: detailsRes && Array.isArray(detailsRes.data) ? toOkxDebugPreview_(detailsRes.data[0]) : null,
    subOrdersCode: subOrdersRes ? subOrdersRes.code : '',
    subOrdersCount: subOrdersRes && Array.isArray(subOrdersRes.data) ? subOrdersRes.data.length : 0,
    firstSubOrderKeys: subOrdersRes && Array.isArray(subOrdersRes.data) && subOrdersRes.data[0] ? Object.keys(subOrdersRes.data[0]) : [],
    firstSubOrder: subOrdersRes && Array.isArray(subOrdersRes.data) ? toOkxDebugPreview_(subOrdersRes.data[0]) : null,
    historyCode: historyRes ? historyRes.code : '',
    historyCount: historyRes && Array.isArray(historyRes.data) ? historyRes.data.length : 0,
    firstHistoryKeys: historyRes && Array.isArray(historyRes.data) && historyRes.data[0] ? Object.keys(historyRes.data[0]) : [],
    firstHistory: historyRes && Array.isArray(historyRes.data) ? toOkxDebugPreview_(historyRes.data[0]) : null,
    derivedSummary: derivedSummary
  };

  setLastOkxRecurringDebugPayload_(preview);
  SyncManager.log("INFO", `[OKX Recurring Debug] ${JSON.stringify(preview)}`, "Sync_Okx");

  return {
    success: true,
    rowCount: pendingList.length,
    message: `method=recurring; pending=${pendingList.length}; firstAlgoId=${firstAlgoId || 'n/a'}; subOrders=${preview.subOrdersCount}; history=${preview.historyCount}`
  };
}

function fetchOkxSpotBtcFillDebug_(baseUrl, apiKey, apiSecret, apiPassphrase) {
  const fillsHistoryRes = fetchOkxApi_(
    baseUrl,
    '/api/v5/trade/fills-history',
    { instType: 'SPOT', instId: 'BTC-USDT', limit: 100 },
    apiKey,
    apiSecret,
    apiPassphrase
  );

  if (fillsHistoryRes.code !== "0") {
    return { success: false, status: `FillsHistory Code: ${fillsHistoryRes.code}, Msg: ${fillsHistoryRes.msg}` };
  }

  const fillsRes = fetchOkxApi_(
    baseUrl,
    '/api/v5/trade/fills',
    { instType: 'SPOT', instId: 'BTC-USDT', limit: 100 },
    apiKey,
    apiSecret,
    apiPassphrase
  );

  const historyRows = Array.isArray(fillsHistoryRes.data) ? fillsHistoryRes.data : [];
  const recentRows = fillsRes.code === "0" && Array.isArray(fillsRes.data) ? fillsRes.data : [];
  const mergedRows = dedupeOkxRowsByBillId_([].concat(historyRows, recentRows));
  const buyRows = mergedRows.filter(isOkxBtcSpotBuyFill_);

  let totalBoughtBtc = 0;
  let totalInvestedUsdt = 0;

  buyRows.forEach(row => {
    const btcQty = parseOkxNumberMaybe_(pickFirstDefinedOkxValue_(row, null, [
      'fillSz',
      'accFillSz',
      'sz'
    ]));
    const fillPx = parseOkxNumberMaybe_(pickFirstDefinedOkxValue_(row, null, [
      'fillPx',
      'px',
      'avgPx'
    ]));
    const explicitUsdt = parseOkxNumberMaybe_(pickFirstDefinedOkxValue_(row, null, [
      'fillNotionalUsd',
      'notionalUsd',
      'fillUsdPx'
    ]));

    totalBoughtBtc += btcQty;
    totalInvestedUsdt += explicitUsdt > 0 ? explicitUsdt : (btcQty * fillPx);
  });

  const preview = {
    method: 'spot_fills_fallback',
    fillsHistoryCode: fillsHistoryRes.code,
    fillsHistoryCount: historyRows.length,
    fillsCode: fillsRes.code,
    fillsCount: recentRows.length,
    mergedCount: mergedRows.length,
    buyCount: buyRows.length,
    firstFillKeys: mergedRows[0] ? Object.keys(mergedRows[0]) : [],
    firstFill: toOkxDebugPreview_(mergedRows[0]),
    derivedSummary: {
      instId: 'BTC-USDT',
      totalBoughtBtc: totalBoughtBtc || null,
      totalInvestedUsdt: totalInvestedUsdt || null,
      derivedAvgPrice: totalBoughtBtc > 0 ? (totalInvestedUsdt / totalBoughtBtc) : null
    }
  };

  setLastOkxRecurringDebugPayload_(preview);
  SyncManager.log("INFO", `[OKX Recurring Debug Fallback] ${JSON.stringify(preview)}`, "Sync_Okx");

  return {
    success: true,
    rowCount: buyRows.length,
    message: `method=spot_fills_fallback; buyRows=${buyRows.length}; totalBoughtBtc=${formatOkxDebugNumber_(totalBoughtBtc)}; avgPx=${formatOkxDebugNumber_(totalBoughtBtc > 0 ? (totalInvestedUsdt / totalBoughtBtc) : 0)}`
  };
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

function extractOkxRecurringAlgoId_(item) {
  if (!item || typeof item !== 'object') return '';
  return String(item.algoId || item.algoClOrdId || item.id || '').trim();
}

function toOkxDebugPreview_(value) {
  if (!value || typeof value !== 'object') return value || null;
  const preview = {};
  Object.keys(value).slice(0, 20).forEach(key => {
    preview[key] = value[key];
  });
  return preview;
}

function buildOkxRecurringDerivedSummary_(pendingItem, detailsRes, subOrdersRes) {
  const details = detailsRes && Array.isArray(detailsRes.data) ? detailsRes.data[0] : null;
  const subOrders = subOrdersRes && Array.isArray(subOrdersRes.data) ? subOrdersRes.data : [];
  const summary = {
    nextInvestTime: pickFirstDefinedOkxValue_(pendingItem, details, ['nextInvestTime', 'nextTime']),
    recurringAmount: pickFirstDefinedOkxValue_(pendingItem, details, ['amount', 'recurringAmount', 'investmentAmt', 'investAmt']),
    quoteCcy: pickFirstDefinedOkxValue_(pendingItem, details, ['quoteCcy', 'quoteCcyName', 'investCcy']),
    baseCcy: pickFirstDefinedOkxValue_(pendingItem, details, ['baseCcy', 'ccy']),
    executeCount: subOrders.length
  };

  let totalInvested = 0;
  let totalBought = 0;

  subOrders.forEach(order => {
    totalInvested += parseOkxNumberMaybe_(pickFirstDefinedOkxValue_(order, null, [
      'investmentAmt',
      'investAmt',
      'amt',
      'sz',
      'orderAmt'
    ]));
    totalBought += parseOkxNumberMaybe_(pickFirstDefinedOkxValue_(order, null, [
      'filledSz',
      'fillSz',
      'accFillSz',
      'baseSz',
      'qty'
    ]));
  });

  if (totalInvested > 0) summary.totalInvested = totalInvested;
  if (totalBought > 0) {
    summary.totalBought = totalBought;
    summary.derivedAvgPrice = totalInvested > 0 ? totalInvested / totalBought : null;
  }

  return summary;
}

function pickFirstDefinedOkxValue_(primary, secondary, keys) {
  const sources = [primary, secondary];
  for (let i = 0; i < sources.length; i++) {
    const src = sources[i];
    if (!src || typeof src !== 'object') continue;
    for (let j = 0; j < keys.length; j++) {
      const value = src[keys[j]];
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return value;
      }
    }
  }
  return null;
}

function parseOkxNumberMaybe_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

function isOkxBtcSpotBuyFill_(row) {
  if (!row || typeof row !== 'object') return false;
  const instId = String(row.instId || '').toUpperCase();
  const side = String(row.side || '').toLowerCase();
  const execType = String(row.execType || '').toLowerCase();
  if (instId !== 'BTC-USDT') return false;
  if (side !== 'buy') return false;
  if (execType && execType !== 't' && execType !== 'm') return true;
  return true;
}

function dedupeOkxRowsByBillId_(rows) {
  const seen = {};
  const output = [];
  rows.forEach(row => {
    if (!row || typeof row !== 'object') return;
    const key = String(
      row.billId ||
      row.tradeId ||
      row.fillId ||
      `${row.instId || ''}:${row.ts || row.fillTime || row.uTime || ''}:${row.fillSz || row.sz || ''}:${row.fillPx || row.px || ''}`
    );
    if (!key || seen[key]) return;
    seen[key] = true;
    output.push(row);
  });
  return output;
}

function formatOkxDebugNumber_(value) {
  if (!value || !isFinite(value)) return '0';
  return Number(value).toFixed(8);
}

function setLastOkxRecurringDebugPayload_(payload) {
  __okxRecurringDebugPayload = payload || null;
}

function getLastOkxRecurringDebugPayload_() {
  return __okxRecurringDebugPayload;
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

function buildOkxStakingMeta_(item) {
  const parts = [];

  pushOkxMetaPart_(parts, 'ordId', item.ordId);
  pushOkxMetaPart_(parts, 'productId', item.productId);
  pushOkxMetaPart_(parts, 'protocol', item.protocol);
  pushOkxMetaPart_(parts, 'protocolType', item.protocolType);
  pushOkxMetaPart_(parts, 'term', item.term);
  pushOkxMetaPart_(parts, 'apy', item.apy);
  pushOkxMetaPart_(parts, 'state', item.state);
  pushOkxMetaPart_(parts, 'purchasedTime', item.purchasedTime);
  pushOkxMetaPart_(parts, 'estSettlementTime', item.estSettlementTime);

  return parts.join('; ');
}

function buildOkxEarnMeta_(item) {
  const parts = [];

  pushOkxMetaPart_(parts, 'rate', item.rate);
  pushOkxMetaPart_(parts, 'loanAmt', item.loanAmt);
  pushOkxMetaPart_(parts, 'pendingAmt', item.pendingAmt);
  pushOkxMetaPart_(parts, 'earnings', item.earnings);
  pushOkxMetaPart_(parts, 'redemptAmt', item.redemptAmt);

  return parts.join('; ');
}
