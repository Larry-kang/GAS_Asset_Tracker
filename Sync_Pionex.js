// =======================================================
// --- Pionex Service Layer (v4.0 - Unified Ledger)
// --- Deployable root copy for GAS / clasp push
// --- Modernized to latest documented Spot / Futures / Earn / Bot APIs
// =======================================================

function getPionexBalance() {
  const MODULE_NAME = "Sync_Pionex";

  return SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = SyncManager.createResult("Pionex");
    const creds = Credentials.get('PIONEX');
    const { apiKey, apiSecret } = creds;
    const baseUrl = 'https://api.pionex.com';

    if (!apiKey || !apiSecret) {
      SyncManager.registerSourceCheck(result, {
        name: 'Credentials',
        required: true,
        success: false,
        message: 'Missing PIONEX_API_KEY / PIONEX_API_SECRET'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }

    // A. Trade Spot
    const spotRes = fetchPionexSpotBalances_(baseUrl, apiKey, apiSecret);
    let spotRows = 0;
    if (spotRes.success && Array.isArray(spotRes.data)) {
      spotRes.data.forEach(item => {
        const coin = pionexUpper_(item.coin);
        const free = parsePionexNumber_(item.free);
        const frozen = parsePionexNumber_(item.frozen);

        if (!coin) return;
        if (free > 0) {
          result.assets.push({ ccy: coin, amt: free, type: 'Spot', status: 'Available' });
          spotRows++;
        }
        if (frozen > 0) {
          result.assets.push({ ccy: coin, amt: frozen, type: 'Spot', status: 'Frozen' });
          spotRows++;
        }
      });
      SyncManager.registerSourceCheck(result, { name: 'Trade Spot', required: true, success: true, rows: spotRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Trade Spot',
        required: true,
        success: false,
        message: pionexStatus_(spotRes.raw)
      });
    }

    // B. Futures Account Detail (optional enrichment)
    const futuresDetailRes = fetchPionexFuturesAccountDetail_(baseUrl, apiKey, apiSecret);
    let futuresDetailLookup = {};
    if (futuresDetailRes.success && futuresDetailRes.data) {
      futuresDetailLookup = buildPionexFuturesDetailLookup_(futuresDetailRes.data);
      SyncManager.registerSourceCheck(result, {
        name: 'Futures Detail',
        required: false,
        success: true,
        rows: (futuresDetailRes.data.balances || []).length + (futuresDetailRes.data.positions || []).length
      });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Futures Detail',
        required: false,
        success: false,
        message: pionexStatus_(futuresDetailRes.raw)
      });
    }

    // C. Futures Balances
    const futuresBalanceRes = fetchPionexFuturesBalances_(baseUrl, apiKey, apiSecret);
    let futuresBalanceRows = 0;
    if (futuresBalanceRes.success && futuresBalanceRes.data) {
      const crossBalances = Array.isArray(futuresBalanceRes.data.balances) ? futuresBalanceRes.data.balances : [];
      const isolatedList = Array.isArray(futuresBalanceRes.data.isolates) ? futuresBalanceRes.data.isolates : [];

      crossBalances.forEach(item => {
        const coin = pionexUpper_(item.coin);
        const free = parsePionexNumber_(item.free);
        const frozen = parsePionexNumber_(item.frozen);
        const debts = parsePionexNumber_(item.debts);
        const meta = buildPionexFuturesCrossMeta_(coin, futuresDetailLookup);

        if (!coin) return;
        if (free > 0) {
          result.assets.push({ ccy: coin, amt: free, type: 'Futures', status: 'Available', meta: meta });
          futuresBalanceRows++;
        }
        if (frozen > 0) {
          result.assets.push({ ccy: coin, amt: frozen, type: 'Futures', status: 'Frozen', meta: meta });
          futuresBalanceRows++;
        }
        if (debts > 0) {
          result.assets.push({ ccy: coin, amt: -Math.abs(debts), type: 'Loan', status: 'Debt', meta: 'Futures Cross Debt' });
          futuresBalanceRows++;
        }
      });

      isolatedList.forEach(entry => {
        const symbol = String(entry.symbol || '');
        const isolatedMode = String(entry.isolatedMode || '');
        const balances = Array.isArray(entry.balances) ? entry.balances : [];

        balances.forEach(item => {
          const coin = pionexUpper_(item.coin);
          const free = parsePionexNumber_(item.free);
          const frozen = parsePionexNumber_(item.frozen);
          const debts = parsePionexNumber_(item.debts);
          const meta = buildPionexIsolatedMeta_(symbol, isolatedMode);

          if (!coin) return;
          if (free > 0) {
            result.assets.push({ ccy: coin, amt: free, type: 'Futures Isolated', status: 'Available', meta: meta });
            futuresBalanceRows++;
          }
          if (frozen > 0) {
            result.assets.push({ ccy: coin, amt: frozen, type: 'Futures Isolated', status: 'Frozen', meta: meta });
            futuresBalanceRows++;
          }
          if (debts > 0) {
            result.assets.push({ ccy: coin, amt: -Math.abs(debts), type: 'Loan', status: 'Debt', meta: `${meta}; isolatedDebt=true` });
            futuresBalanceRows++;
          }
        });
      });

      SyncManager.registerSourceCheck(result, { name: 'Futures Balances', required: true, success: true, rows: futuresBalanceRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Futures Balances',
        required: true,
        success: false,
        message: pionexStatus_(futuresBalanceRes.raw)
      });
    }

    // D. Futures Positions
    const futuresPosRes = fetchPionexFuturesPositions_(baseUrl, apiKey, apiSecret);
    let futuresPositionRows = 0;
    if (futuresPosRes.success && Array.isArray(futuresPosRes.data)) {
      futuresPosRes.data.forEach(item => {
        const netSize = parsePionexNumber_(item.netSize);
        const ccy = pionexBaseFromSymbol_(item.symbol);
        if (!ccy || netSize === 0) return;

        result.assets.push({
          ccy: ccy,
          amt: netSize,
          type: 'Positions',
          status: 'Position',
          meta: buildPionexPositionMeta_(item)
        });
        futuresPositionRows++;
      });
      SyncManager.registerSourceCheck(result, { name: 'Futures Positions', required: true, success: true, rows: futuresPositionRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Futures Positions',
        required: true,
        success: false,
        message: pionexStatus_(futuresPosRes.raw)
      });
    }

    // E. Earn Dual Balances
    const dualBalanceRes = fetchPionexDualBalances_(baseUrl, apiKey, apiSecret);
    let dualRows = 0;
    let dualRecordSummary = {};
    if (dualBalanceRes.success && Array.isArray(dualBalanceRes.data)) {
      const bases = uniqueNonEmptyStrings_(dualBalanceRes.data.map(item => pionexUpper_(item.base)));
      const dualRecordsRes = fetchPionexDualRecordsByBases_(baseUrl, apiKey, apiSecret, bases);

      if (dualRecordsRes.success) {
        dualRecordSummary = buildPionexDualRecordSummary_(dualRecordsRes.data || []);
        SyncManager.registerSourceCheck(result, {
          name: 'Earn Dual Records',
          required: false,
          success: true,
          rows: (dualRecordsRes.data || []).length
        });
      } else {
        SyncManager.registerSourceCheck(result, {
          name: 'Earn Dual Records',
          required: false,
          success: false,
          message: pionexStatus_(dualRecordsRes.raw)
        });
      }

      dualBalanceRes.data.forEach(item => {
        const base = pionexUpper_(item.base);
        const coin = pionexUpper_(item.coin);
        const free = parsePionexNumber_(item.free);
        const frozen = parsePionexNumber_(item.frozen);
        const meta = buildPionexDualMeta_(base, coin, dualRecordSummary);

        if (!coin) return;
        if (free > 0) {
          result.assets.push({ ccy: coin, amt: free, type: 'Earn Dual', status: 'Available', meta: meta });
          dualRows++;
        }
        if (frozen > 0) {
          result.assets.push({ ccy: coin, amt: frozen, type: 'Earn Dual', status: 'Frozen', meta: meta });
          dualRows++;
        }
      });

      SyncManager.registerSourceCheck(result, { name: 'Earn Dual Balances', required: true, success: true, rows: dualRows });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Earn Dual Balances',
        required: true,
        success: false,
        message: pionexStatus_(dualBalanceRes.raw)
      });
    }

    // F. Bot Order List
    const botListRes = fetchPionexBotOrders_(baseUrl, apiKey, apiSecret);
    let botRows = 0;
    if (botListRes.success && Array.isArray(botListRes.data)) {
      const futuresGridOrders = botListRes.data.filter(item => String(item.buOrderType) === 'futures_grid');
      const spotGridOrders = botListRes.data.filter(item => String(item.buOrderType) === 'spot_grid');
      const smartCopyOrders = botListRes.data.filter(item => String(item.buOrderType) === 'smart_copy');

      SyncManager.registerSourceCheck(result, { name: 'Bot Order List', required: true, success: true, rows: botListRes.data.length });

      if (smartCopyOrders.length > 0) {
        result.warnings.push(`Pionex smart_copy orders detected (${smartCopyOrders.length}), but no safe asset-detail mapping is implemented.`);
      }

      // F1. Futures Grid Detail
      if (futuresGridOrders.length > 0) {
        const futuresGridRes = fetchPionexFuturesGridOrders_(baseUrl, apiKey, apiSecret, futuresGridOrders);
        if (futuresGridRes.success && Array.isArray(futuresGridRes.data)) {
          futuresGridRes.data.forEach(order => {
            const rowsAdded = appendPionexFuturesGridAssets_(result.assets, order);
            botRows += rowsAdded;
          });
          SyncManager.registerSourceCheck(result, {
            name: 'Bot Futures Grid Detail',
            required: true,
            success: true,
            rows: botRows
          });
        } else {
          SyncManager.registerSourceCheck(result, {
            name: 'Bot Futures Grid Detail',
            required: true,
            success: false,
            message: pionexStatus_(futuresGridRes.raw)
          });
        }
      } else {
        SyncManager.registerSourceCheck(result, {
          name: 'Bot Futures Grid Detail',
          required: false,
          success: true,
          rows: 0,
          message: 'No running futures grid orders'
        });
      }

      // F2. Spot Grid Detail
      let spotGridRows = 0;
      if (spotGridOrders.length > 0) {
        const spotGridRes = fetchPionexSpotGridOrders_(baseUrl, apiKey, apiSecret, spotGridOrders);
        if (spotGridRes.success && Array.isArray(spotGridRes.data)) {
          spotGridRes.data.forEach(order => {
            const rowsAdded = appendPionexSpotGridAssets_(result.assets, order);
            spotGridRows += rowsAdded;
          });
          botRows += spotGridRows;
          SyncManager.registerSourceCheck(result, {
            name: 'Bot Spot Grid Detail',
            required: true,
            success: true,
            rows: spotGridRows
          });
        } else {
          SyncManager.registerSourceCheck(result, {
            name: 'Bot Spot Grid Detail',
            required: true,
            success: false,
            message: pionexStatus_(spotGridRes.raw)
          });
        }
      } else {
        SyncManager.registerSourceCheck(result, {
          name: 'Bot Spot Grid Detail',
          required: false,
          success: true,
          rows: 0,
          message: 'No running spot grid orders'
        });
      }
    } else {
      const botPermissionDenied = isPionexPermissionDenied_(botListRes.raw);
      SyncManager.registerSourceCheck(result, {
        name: 'Bot Order List',
        required: !botPermissionDenied,
        success: false,
        message: botPermissionDenied
          ? `Bot reading (Beta) permission missing. ${pionexStatus_(botListRes.raw)}`
          : pionexStatus_(botListRes.raw)
      });
    }

    SyncManager.log("INFO", `Collected ${result.assets.length} asset entries from Pionex.`, MODULE_NAME);
    return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
  });
}

function fetchPionexSpotBalances_(baseUrl, apiKey, apiSecret) {
  const res = fetchPionexApi_(baseUrl, '/api/v1/account/balances', {}, apiKey, apiSecret);
  if (res.result === true && Array.isArray(res.data && res.data.balances)) {
    return { success: true, data: res.data.balances, raw: res };
  }
  return { success: false, raw: res };
}

function fetchPionexFuturesBalances_(baseUrl, apiKey, apiSecret) {
  const res = fetchPionexApi_(baseUrl, '/uapi/v1/account/balances', {}, apiKey, apiSecret);
  if (res.result === true && res.data) {
    return { success: true, data: res.data, raw: res };
  }
  return { success: false, raw: res };
}

function fetchPionexFuturesPositions_(baseUrl, apiKey, apiSecret) {
  const res = fetchPionexApi_(baseUrl, '/uapi/v1/account/positions', {}, apiKey, apiSecret);
  if (res.result === true && Array.isArray(res.data && res.data.positions)) {
    return { success: true, data: res.data.positions, raw: res };
  }
  return { success: false, raw: res };
}

function fetchPionexFuturesAccountDetail_(baseUrl, apiKey, apiSecret) {
  const res = fetchPionexApi_(baseUrl, '/uapi/v1/account/detail', {}, apiKey, apiSecret);
  if (res.result === true && res.data) {
    return { success: true, data: res.data, raw: res };
  }
  return { success: false, raw: res };
}

function fetchPionexDualBalances_(baseUrl, apiKey, apiSecret) {
  const res = fetchPionexApi_(baseUrl, '/api/v1/earn/dual/balances', { merge: false }, apiKey, apiSecret);
  if (res.result === true && Array.isArray(res.data && res.data.balances)) {
    return { success: true, data: res.data.balances, raw: res };
  }
  return { success: false, raw: res };
}

function fetchPionexDualRecordsByBases_(baseUrl, apiKey, apiSecret, bases) {
  if (!bases || bases.length === 0) {
    return { success: true, data: [], raw: { result: true, data: { records: [] } } };
  }

  const allRecords = [];
  const endTime = Date.now();
  const startTime = endTime - (90 * 24 * 60 * 60 * 1000);

  for (let i = 0; i < bases.length; i++) {
    const base = bases[i];
    const res = fetchPionexApi_(baseUrl, '/api/v1/earn/dual/records', {
      base: base,
      startTime: startTime,
      endTime: endTime,
      limit: 20
    }, apiKey, apiSecret);

    if (!(res.result === true && Array.isArray(res.data && res.data.records))) {
      return { success: false, raw: res };
    }

    allRecords.push.apply(allRecords, res.data.records);
  }

  return { success: true, data: allRecords, raw: { result: true, data: { records: allRecords } } };
}

function fetchPionexBotOrders_(baseUrl, apiKey, apiSecret) {
  const allResults = [];
  let pageToken = '';
  let loops = 0;

  while (loops < 20) {
    const params = { status: 'running' };
    if (pageToken) params.pageToken = pageToken;

    const res = fetchPionexApi_(baseUrl, '/api/v1/bot/orders', params, apiKey, apiSecret);
    if (!(res.result === true && res.data)) {
      return { success: false, raw: res };
    }

    const pageResults = Array.isArray(res.data.results) ? res.data.results : [];
    allResults.push.apply(allResults, pageResults);

    pageToken = String(res.data.nextPageToken || '');
    if (!pageToken) break;
    loops++;
  }

  return { success: true, data: allResults, raw: { result: true, data: { results: allResults } } };
}

function fetchPionexFuturesGridOrders_(baseUrl, apiKey, apiSecret, orders) {
  const details = [];
  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    const buOrderId = String(order.buOrderId || '');
    if (!buOrderId) continue;

    const res = fetchPionexApi_(baseUrl, '/api/v1/bot/orders/futuresGrid/order', { buOrderId: buOrderId }, apiKey, apiSecret);
    if (!(res.result === true && res.data)) {
      return { success: false, raw: res };
    }
    details.push(res.data);
  }

  return { success: true, data: details, raw: { result: true, data: { results: details } } };
}

function fetchPionexSpotGridOrders_(baseUrl, apiKey, apiSecret, orders) {
  const details = [];
  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];
    const buOrderId = String(order.buOrderId || '');
    if (!buOrderId) continue;

    const res = fetchPionexApi_(baseUrl, '/api/v1/bot/orders/spotGrid/order', { buOrderId: buOrderId }, apiKey, apiSecret);
    if (!(res.result === true && res.data)) {
      return { success: false, raw: res };
    }
    details.push(res.data);
  }

  return { success: true, data: details, raw: { result: true, data: { results: details } } };
}

function fetchPionexApi_(baseUrl, endpoint, params, apiKey, apiSecret, method, body) {
  const upperMethod = String(method || 'GET').toUpperCase();
  const timestamp = Date.now();
  const signedParams = {};
  const sourceParams = params || {};

  Object.keys(sourceParams).forEach(key => {
    const value = sourceParams[key];
    if (value !== null && value !== undefined && value !== '') {
      signedParams[key] = value;
    }
  });
  signedParams.timestamp = timestamp;

  const queryString = buildPionexQueryString_(signedParams);
  const pathUrl = `${endpoint}?${queryString}`;
  const bodyString = body ? JSON.stringify(body) : '';
  const payload = `${upperMethod}${pathUrl}${bodyString}`;
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, payload, apiSecret)
    .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
    .join('');
  const url = `${baseUrl}${pathUrl}`;

  const options = {
    method: upperMethod,
    headers: {
      'PIONEX-KEY': apiKey,
      'PIONEX-SIGNATURE': signature,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (bodyString) {
    options.payload = bodyString;
  }

  try {
    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
    if (!text) return { result: false, code: 'EMPTY_RESPONSE', message: `HTTP ${response.getResponseCode()}` };

    try {
      return JSON.parse(text);
    } catch (e) {
      return {
        result: false,
        code: `HTTP_${response.getResponseCode()}`,
        message: `Non-JSON response: ${text}`
      };
    }
  } catch (e) {
    return {
      result: false,
      code: 'FETCH_ERROR',
      message: e.message
    };
  }
}

function buildPionexQueryString_(params) {
  const keys = Object.keys(params).sort();
  return keys.map(function (key) { return `${key}=${params[key]}`; }).join('&');
}

function parsePionexNumber_(value) {
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

function pionexUpper_(value) {
  return String(value || '').trim().toUpperCase();
}

function pionexBaseFromSymbol_(symbol) {
  const text = String(symbol || '').trim();
  if (!text) return '';
  return pionexUpper_(text.split('_')[0].replace('.PERP', ''));
}

function pionexStatus_(res) {
  if (!res) return 'Unknown error';
  const code = res.code || (res.result === true ? 'OK' : 'UNKNOWN');
  const msg = res.message || res.msg || 'Unknown error';
  return `Code: ${code}, Msg: ${msg}`;
}

function isPionexPermissionDenied_(res) {
  return String(res && res.code || '').trim().toUpperCase() === 'PERMISSION_DENIED';
}

function buildPionexFuturesDetailLookup_(detail) {
  const lookup = {};
  const balances = Array.isArray(detail.balances) ? detail.balances : [];
  balances.forEach(item => {
    const coin = pionexUpper_(item.coin);
    if (!coin) return;
    lookup[coin] = item;
  });
  return lookup;
}

function buildPionexFuturesCrossMeta_(coin, detailLookup) {
  const detail = detailLookup && detailLookup[coin];
  if (!detail) return '';

  const parts = [];
  const available = parsePionexNumber_(detail.available);
  const transferable = parsePionexNumber_(detail.transferable);
  const unrealizedPnL = parsePionexNumber_(detail.unrealizedPnL);
  const totalInitialMargin = parsePionexNumber_(detail.totalInitialMargin);
  const debts = parsePionexNumber_(detail.debts);

  if (available !== 0) parts.push(`available=${available}`);
  if (transferable !== 0) parts.push(`transferable=${transferable}`);
  if (unrealizedPnL !== 0) parts.push(`unrealizedPnL=${unrealizedPnL}`);
  if (totalInitialMargin !== 0) parts.push(`initMargin=${totalInitialMargin}`);
  if (debts !== 0) parts.push(`debts=${debts}`);
  return parts.join('; ');
}

function buildPionexIsolatedMeta_(symbol, isolatedMode) {
  const parts = [];
  if (symbol) parts.push(`symbol=${symbol}`);
  if (isolatedMode) parts.push(`isolatedMode=${isolatedMode}`);
  return parts.join('; ');
}

function buildPionexPositionMeta_(item) {
  const parts = [];
  if (item.symbol) parts.push(`symbol=${item.symbol}`);
  if (item.positionSide) parts.push(`side=${item.positionSide}`);
  if (item.isolatedMode) parts.push(`mode=${item.isolatedMode}`);
  if (item.avgPrice) parts.push(`avgPrice=${item.avgPrice}`);
  if (item.markPrice) parts.push(`markPrice=${item.markPrice}`);
  if (item.unrealizedPnL) parts.push(`unrealizedPnL=${item.unrealizedPnL}`);
  if (item.leverage) parts.push(`leverage=${item.leverage}`);
  if (item.liquidationPrice) parts.push(`liqPrice=${item.liquidationPrice}`);
  if (item.riskState) parts.push(`risk=${item.riskState}`);
  return parts.join('; ');
}

function buildPionexDualRecordSummary_(records) {
  const summary = {};

  records.forEach(record => {
    const base = pionexUpper_(record.base);
    const coin = pionexUpper_(record.currency);
    const key = `${base}|${coin}`;
    if (!summary[key]) {
      summary[key] = { count: 0, latestState: '', latestProductId: '', latestProfit: '' };
    }

    summary[key].count++;
    summary[key].latestState = String(record.state || '');
    summary[key].latestProductId = String(record.productId || '');
    summary[key].latestProfit = String(record.profit || '');
  });

  return summary;
}

function buildPionexDualMeta_(base, coin, recordSummary) {
  const parts = [];
  if (base) parts.push(`base=${base}`);

  const key = `${base}|${coin}`;
  const summary = recordSummary ? recordSummary[key] : null;
  if (summary) {
    if (summary.count > 0) parts.push(`recentRecords=${summary.count}`);
    if (summary.latestState) parts.push(`latestState=${summary.latestState}`);
    if (summary.latestProductId) parts.push(`latestProductId=${summary.latestProductId}`);
    if (summary.latestProfit) parts.push(`latestProfit=${summary.latestProfit}`);
  }

  return parts.join('; ');
}

function appendPionexSpotGridAssets_(assetList, order) {
  let rows = 0;
  const base = pionexUpper_(order.base);
  const quote = pionexUpper_(order.quote);
  const data = order.buOrderData || {};
  const baseAmount = parsePionexNumber_(data.baseAmount);
  const quoteAmount = parsePionexNumber_(data.quoteAmount);
  const meta = buildPionexSpotGridMeta_(order, data);

  if (base && baseAmount > 0) {
    assetList.push({ ccy: base, amt: baseAmount, type: 'Bot Spot Grid', status: 'Base', meta: meta });
    rows++;
  }
  if (quote && quoteAmount > 0) {
    assetList.push({ ccy: quote, amt: quoteAmount, type: 'Bot Spot Grid', status: 'Quote', meta: meta });
    rows++;
  }

  return rows;
}

function appendPionexFuturesGridAssets_(assetList, order) {
  let rows = 0;
  const base = pionexUpper_(order.base);
  const quote = pionexUpper_(order.quote);
  const data = order.buOrderData || {};
  const marginBalance = parsePionexNumber_(data.marginBalance);
  const extraBalance = parsePionexNumber_(data.extraBalance);
  const position = parsePionexNumber_(data.position);
  const meta = buildPionexFuturesGridMeta_(order, data);

  if (quote && marginBalance > 0) {
    assetList.push({ ccy: quote, amt: marginBalance, type: 'Bot Futures Grid', status: 'Margin', meta: meta });
    rows++;
  }
  if (quote && extraBalance > 0) {
    assetList.push({ ccy: quote, amt: extraBalance, type: 'Bot Futures Grid', status: 'Extra', meta: meta });
    rows++;
  }
  if (base && position !== 0) {
    assetList.push({ ccy: base, amt: position, type: 'Bot Futures Grid', status: 'Position', meta: meta });
    rows++;
  }

  return rows;
}

function buildPionexSpotGridMeta_(order, data) {
  const parts = [];
  if (order.buOrderId) parts.push(`buOrderId=${order.buOrderId}`);
  if (data.status) parts.push(`status=${data.status}`);
  if (data.gridType) parts.push(`gridType=${data.gridType}`);
  if (data.top) parts.push(`top=${data.top}`);
  if (data.bottom) parts.push(`bottom=${data.bottom}`);
  if (data.quoteInvestment) parts.push(`quoteInvestment=${data.quoteInvestment}`);
  if (data.gridProfit) parts.push(`gridProfit=${data.gridProfit}`);
  if (data.realizedProfit) parts.push(`realizedProfit=${data.realizedProfit}`);
  if (data.profitWithdrawn) parts.push(`profitWithdrawn=${data.profitWithdrawn}`);
  return parts.join('; ');
}

function buildPionexFuturesGridMeta_(order, data) {
  const parts = [];
  if (order.buOrderId) parts.push(`buOrderId=${order.buOrderId}`);
  if (data.status) parts.push(`status=${data.status}`);
  if (data.trend) parts.push(`trend=${data.trend}`);
  if (data.leverage) parts.push(`leverage=${data.leverage}`);
  if (data.quoteInvestment) parts.push(`quoteInvestment=${data.quoteInvestment}`);
  if (data.positionOpenPrice) parts.push(`positionOpenPrice=${data.positionOpenPrice}`);
  if (data.liquidationPrice) parts.push(`liqPrice=${data.liquidationPrice}`);
  if (data.riskStatus) parts.push(`risk=${data.riskStatus}`);
  if (data.marginStatus) parts.push(`marginStatus=${data.marginStatus}`);
  if (data.profitWithdrawn) parts.push(`profitWithdrawn=${data.profitWithdrawn}`);
  return parts.join('; ');
}

function uniqueNonEmptyStrings_(values) {
  const map = {};
  values.forEach(value => {
    const text = String(value || '').trim();
    if (text) map[text] = true;
  });
  return Object.keys(map);
}
