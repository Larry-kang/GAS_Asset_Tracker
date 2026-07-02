// =======================================================
// --- BitoPro Service Layer (v3.0 - Unified Ledger) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getBitoProBalance() {
  const MODULE_NAME = "Sync_BitoPro";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = SyncManager.createResult("BitoPro");
    const creds = Credentials.get('BITOPRO');
    const { apiKey, apiSecret } = creds;
    const baseUrl = 'https://api.bitopro.com/v3';

    if (!apiKey || !apiSecret) {
      SyncManager.registerSourceCheck(result, {
        name: 'Credentials',
        required: true,
        success: false,
        message: 'Missing BitoPro API Keys'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }

    const json = fetchBitoProApi_(baseUrl, '/accounts/balance', 'GET', {}, apiKey, apiSecret);

    if (json && json.data) {
      json.data.forEach(b => {
        const avail = parseFloat(b.available) || 0;
        const stake = parseFloat(b.stake) || 0;

        if (avail > 0) {
          result.assets.push({ ccy: b.currency.toUpperCase(), amt: avail, type: 'Spot', status: 'Available' });
        }
        if (stake > 0) {
          result.assets.push({ ccy: b.currency.toUpperCase(), amt: stake, type: 'Spot', status: 'Frozen' });
        }
      });

      SyncManager.registerSourceCheck(result, {
        name: 'Balance',
        required: true,
        success: true,
        rows: result.assets.length
      });
      SyncManager.log("INFO", `Collected ${result.assets.length} asset entries from BitoPro.`, MODULE_NAME);
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Balance',
        required: true,
        success: false,
        message: `BitoPro API Error: ${JSON.stringify(json)}`
      });
    }

    const btcCostRes = typeof syncBitoProBtcSpotCostBasis_ === 'function'
      ? syncBitoProBtcSpotCostBasis_(ss, baseUrl, apiKey, apiSecret)
      : { success: false, status: 'syncBitoProBtcSpotCostBasis_ missing' };
    SyncManager.registerSourceCheck(result, {
      name: 'BTC Spot Cost',
      required: false,
      success: !!btcCostRes.success,
      rows: btcCostRes.rowCount || 0,
      message: btcCostRes.message || btcCostRes.status || ''
    });

    return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
  });
}

// ... Helpers (English Logs) ...

function fetchBitoProApi_(baseUrl, endpoint, method, params, apiKey, apiSecret) {
  const nonce = Date.now() - 2000; // Time Skew Fix
  const payloadObj = { ...params, nonce: nonce };
  const payloadJson = JSON.stringify(payloadObj);
  const payloadBase64 = Utilities.base64Encode(payloadJson);
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_384, payloadBase64, apiSecret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  const options = {
    'method': method,
    'headers': {
      'X-BITOPRO-APIKEY': apiKey,
      'X-BITOPRO-PAYLOAD': payloadBase64,
      'X-BITOPRO-SIGNATURE': signature,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };
  if (method === 'POST') options.payload = payloadJson;

  try {
    const response = UrlFetchApp.fetch(baseUrl + endpoint, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    return { error: e.message };
  }
}

function syncBitoProBtcSpotCostBasis_(ss, baseUrl, apiKey, apiSecret) {
  try {
    const sheet = typeof getApiSummaryExportSheet_ === 'function' ? getApiSummaryExportSheet_(ss) : null;
    if (!sheet) {
      return { success: false, status: 'API_Summary_Export not found' };
    }

    const state = typeof readApiSummaryExportState_ === 'function'
      ? readApiSummaryExportState_(sheet, 'BITOPRO_BTC_Spot_')
      : {};
    const summary = fetchBitoProBtcSpotTradeSummary_(baseUrl, apiKey, apiSecret, state);
    if (!summary.success) return summary;

    const entries = buildBitoProBtcSpotSummaryEntries_(summary);
    if (typeof upsertApiSummaryExportEntries_ === 'function') {
      upsertApiSummaryExportEntries_(sheet, entries);
      if (typeof writeAggregateBtcSpotCostSummary_ === 'function') {
        writeAggregateBtcSpotCostSummary_(ss, sheet, entries);
      }
    }

    return {
      success: true,
      rowCount: summary.incrementalBuyCount || 0,
      message: summary.message || ''
    };
  } catch (e) {
    return { success: false, status: e.message || String(e) };
  }
}

function fetchBitoProBtcSpotTradeSummary_(baseUrl, apiKey, apiSecret, state) {
  const pair = 'btc_twd';
  const historyWindowDays = 90;
  const now = Date.now();
  const defaultStart = now - historyWindowDays * 24 * 60 * 60 * 1000;
  const startTimestamp = Math.max(parseInt(state.LastTradeTime || 0, 10) || defaultStart, defaultStart);
  const checkpointTradeId = String(state.LastTradeId || '').trim();
  const checkpointTradeTime = parseInt(state.LastTradeTime || 0, 10) || 0;
  const rows = fetchAllBitoProTradesForPair_(baseUrl, apiKey, apiSecret, pair, startTimestamp, now);
  if (!rows.success) return rows;

  const priorBought = parseSummaryNumber_(state.TotalBought_BTC);
  const priorInvested = parseSummaryNumber_(state.TotalInvested_TWD);
  const priorBuyCount = parseInt(state.BuyCount || 0, 10) || 0;
  const latestTrade = rows.rows.reduce((latest, row) => {
    if (!latest) return row;
    return (parseInt(row.createdTimestamp || row.timestamp || 0, 10) || 0) > (parseInt(latest.createdTimestamp || latest.timestamp || 0, 10) || 0)
      ? row
      : latest;
  }, null);

  let incrementalBoughtBtc = 0;
  let incrementalInvestedTwd = 0;
  rows.rows.forEach(row => {
    const tradeId = String(row.tradeId || '').trim();
    const tradeTime = parseInt(row.createdTimestamp || row.timestamp || 0, 10) || 0;
    if (checkpointTradeId && tradeId === checkpointTradeId) return;
    if (checkpointTradeTime && tradeTime < checkpointTradeTime) return;
    if (String(row.action || '').toUpperCase() !== 'BUY') return;
    incrementalBoughtBtc += parseFloat(row.baseAmount) || 0;
    incrementalInvestedTwd += parseFloat(row.quoteAmount) || 0;
  });

  const incrementalBuyCount = rows.rows.filter(row => {
    const tradeId = String(row.tradeId || '').trim();
    const tradeTime = parseInt(row.createdTimestamp || row.timestamp || 0, 10) || 0;
    if (checkpointTradeId && tradeId === checkpointTradeId) return false;
    if (checkpointTradeTime && tradeTime < checkpointTradeTime) return false;
    return String(row.action || '').toUpperCase() === 'BUY';
  }).length;
  const totalBoughtBtc = priorBought + incrementalBoughtBtc;
  const totalInvestedTwd = priorInvested + incrementalInvestedTwd;

  return {
    success: true,
    pair: pair,
    historyWindowDays: historyWindowDays,
    buyCount: priorBuyCount + incrementalBuyCount,
    incrementalBuyCount: incrementalBuyCount,
    totalBoughtBtc: totalBoughtBtc,
    totalInvestedTwd: totalInvestedTwd,
    derivedAvgPrice: totalBoughtBtc > 0 ? (totalInvestedTwd / totalBoughtBtc) : '',
    lastTradeId: latestTrade ? latestTrade.tradeId : (state.LastTradeId || ''),
    lastTradeTime: latestTrade ? (latestTrade.createdTimestamp || latestTrade.timestamp) : (state.LastTradeTime || ''),
    message: `pair=${pair}; windowDays=${historyWindowDays}; newRows=${incrementalBuyCount}; totalBoughtBtc=${Number(totalBoughtBtc).toFixed(8)}`
  };
}

function fetchAllBitoProTradesForPair_(baseUrl, apiKey, apiSecret, pair, startTimestamp, endTimestamp) {
  const rows = [];
  const seen = {};
  let tradeIdCursor = '';
  const maxPages = 10;

  for (let page = 0; page < maxPages; page++) {
    const params = {
      startTimestamp: startTimestamp,
      endTimestamp: endTimestamp,
      limit: 1000
    };
    if (tradeIdCursor) params.tradeId = tradeIdCursor;

    const endpoint = `/orders/trades/${pair}`;
    const res = fetchBitoProApi_(baseUrl, endpoint, 'GET', params, apiKey, apiSecret);
    if (res && res.error) {
      return { success: false, status: `BitoPro trades ${pair} failed: ${res.error}` };
    }
    const pageRows = res && Array.isArray(res.data) ? res.data : [];
    if (!pageRows.length) break;

    let oldestTradeId = '';
    pageRows.forEach(row => {
      const tradeId = String(row.tradeId || '').trim();
      if (!tradeId || seen[tradeId]) return;
      seen[tradeId] = true;
      rows.push(row);
      oldestTradeId = tradeId;
    });

    if (pageRows.length < 1000 || !oldestTradeId || oldestTradeId === tradeIdCursor) break;
    tradeIdCursor = oldestTradeId;
  }

  return { success: true, rows: rows };
}

function buildBitoProBtcSpotSummaryEntries_(summary) {
  const timestamp = new Date().toISOString();
  return [
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_BuyCount', summary.buyCount || '', 'bitopro_btc_cost_sync', timestamp, 'Cumulative BitoPro BTC spot buy trade count'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_IncrementalBuyCount', summary.incrementalBuyCount || '', 'bitopro_btc_cost_sync', timestamp, 'New BitoPro BTC spot buy trades in this sync'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_TotalBought_BTC', summary.totalBoughtBtc || '', 'bitopro_btc_cost_sync', timestamp, 'Cumulative BTC bought from BitoPro spot trades'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_TotalInvested_TWD', summary.totalInvestedTwd || '', 'bitopro_btc_cost_sync', timestamp, 'Cumulative BitoPro BTC spot cost in TWD'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_DerivedAvgPrice', summary.derivedAvgPrice || '', 'bitopro_btc_cost_sync', timestamp, 'Derived BitoPro BTC spot avg price in TWD/BTC'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_LastTradeId', summary.lastTradeId || '', 'bitopro_btc_cost_sync', timestamp, 'BitoPro BTC spot latest trade id'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_LastTradeTime', summary.lastTradeTime || '', 'bitopro_btc_cost_sync', timestamp, 'BitoPro BTC spot latest trade timestamp (ms)'),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_Pair', summary.pair || '', 'bitopro_btc_cost_sync', timestamp, summary.message || ''),
    buildSummaryExportEntry_('BITOPRO_BTC_Spot_HistoryWindowDays', summary.historyWindowDays || '', 'bitopro_btc_cost_sync', timestamp, 'BitoPro API trade history is limited to the latest 90 days')
  ];
}
