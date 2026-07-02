// 檔案: WebhookHandler.gs

/**
 * [Router] 統一入口，根據 action 分發任務
 */
function doPost(e) {
  try {
    // 防止 e 為 undefined (例如直接在編輯器按執行時)
    if (!e || !e.postData) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "No Post Data" }));
    }

    const data = JSON.parse(e.postData.contents);

    // 1. 全局驗證
    if (!isAuthorizedWebhookRequest_(data)) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "Auth Failed" }));
    }

    // 2. 路由分發
    console.log(`[Webhook] Received Action: ${data.action}`); // DEBUG LOG

    switch (data.action) {
      case 'update_tunnel_url':
        return handleTunnelUpdate(data);

      case 'trigger_balance_update':
        return handleForceUpdate(data);

      case 'log_client_error':
        return handleLogClientError(data);

      case 'trigger_report':
        return handleTriggerReport(data);

      case 'get_inventory':
        return handleGetInventory(data);

      case 'get_quick_status':
        return handleQuickSummary(data);

      case 'debug_okx_recurring':
        return handleDebugOkxRecurring(data);

      default:
        console.warn(`[Webhook] Valid Action NOT FOUND. Payload: ${JSON.stringify(data)}`);
        return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "Unknown Action", received: data.action }));
    }

  } catch (err) {
    console.error("[Webhook] Crash: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: err.toString() }));
  }
}

// --- Controllers / Handlers ---

function isAuthorizedWebhookRequest_(data) {
  const providedPassword = String((data && data.password) || '').trim();

  const allowedPasswords = [];
  const proxyPassword = Settings.get('PROXY_PASSWORD');
  if (proxyPassword) allowedPasswords.push(proxyPassword);

  // 若尚未設定任何密碼，允許第一次連線
  if (allowedPasswords.length === 0) return true;

  return allowedPasswords.indexOf(providedPassword) >= 0;
}

function handleTunnelUpdate(data) {
  const oldUrl = Settings.get('TUNNEL_URL');
  Settings.set('TUNNEL_URL', data.url);

  // ⭐ 修改建議：總是更新密碼，以保持與電腦端同步
  // 因為這一步是在全域驗證通過後(或第一次)才執行的，所以是安全的
  if (data.password) {
    Settings.set('PROXY_PASSWORD', data.password);
  }

  console.log(`[Webhook] Tunnel URL Updated: ${oldUrl} -> ${data.url}`);
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    msg: "URL Updated",
    data: {
      url: data.url,
      timestamp: new Date().toISOString()
    }
  }));
}

function handleForceUpdate(data) {
  const context = "Webhook:Sync";
  LogService.info("Discord User triggered manual balance sync.", context);

  // 這裡呼叫 Binance.gs 裡面的函式 (假設其他同步函式也在此邏輯內或由 runAutomationMaster 處理)
  // 為了精確，我們觸發全系統自動化
  if (typeof runAutomationMaster === 'function') {
    const result = runAutomationMaster();
    if (result && result.fatal) {
      LogService.error(`System-wide balance sync failed: ${result.message || result.status}`, context);
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        msg: "System-wide balance sync failed.",
        error: result.message || result.status,
        timestamp: new Date().toISOString()
      }));
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      msg: "System-wide balance sync completed successfully.",
      result: result || null,
      timestamp: new Date().toISOString()
    }));
  } else {
    // Fallback to Binance only if Master is missing
    getBinanceBalance();
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      msg: "Binance sync triggered (Master sync function missing).",
      timestamp: new Date().toISOString()
    }));
  }
}

function handleLogClientError(data) {
  Logger.log("[Client Error] " + data.message);
  return ContentService.createTextOutput(JSON.stringify({ status: "success", msg: "Error Logged" }));
}

/**
 * [手動] 強制遠端電腦重啟 Bridge
 * 在選單點擊後，會發送指令給 PowerShell，約 2-5 秒後生效
 */
function triggerRemoteRestart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseUrl = Settings.get('TUNNEL_URL');
  const proxyPassword = Settings.get('PROXY_PASSWORD');

  if (!baseUrl || !proxyPassword) {
    ss.toast("❌ 失敗：找不到連線資訊，電腦可能未開機。");
    return;
  }

  const url = `${baseUrl}/restart`;
  const options = {
    'method': 'POST', // 使用 POST
    'headers': {
      'x-proxy-auth': proxyPassword
    },
    'muteHttpExceptions': true
  };

  try {
    ss.toast("⏳ 正在發送重啟指令...");
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      ss.toast("✅ 指令已發送！PowerShell 將在幾秒內重啟。");
      console.log("Remote Restart Triggered Successfully.");
    } else {
      ss.toast(`⚠️ 發送失敗 (Code: ${code})`);
      console.log(`Remote Restart Failed: ${response.getContentText()}`);
    }

  } catch (e) {
    ss.toast("❌ 連線錯誤：無法連接到電腦。");
    console.error(e);
  }
}

function handleTriggerReport(data) {
  const context = "Webhook:Report";
  LogService.info("Discord User requested manual tactical report.", context);

  if (typeof triggerManualReport === 'function') {
    const result = triggerManualReport();
    return ContentService.createTextOutput(JSON.stringify({
      ...result,
      timestamp: new Date().toISOString()
    }));
  } else {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "StrategicEngine not loaded or function missing" }));
  }
}

function handleGetInventory(data) {
  // [ReadOnly] Export calculated context for SAR simulation
  if (typeof buildContext === 'function') {
    const context = typeof buildFreshContext === 'function' ? buildFreshContext() : buildContext();
    const exportBundle = typeof getInventoryExportBundle_ === 'function'
      ? getInventoryExportBundle_()
      : null;

    if (exportBundle && exportBundle.available) {
      context.inventoryExport = exportBundle;
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      source: "GAS_Asset_Tracker",
      timestamp: new Date().toISOString(),
      data: context
    }));
  } else {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "StrategicEngine (buildContext) not found" }));
  }
}

/**
 * [Discord Optimized] 回傳精簡版狀態快照
 */
function handleQuickSummary(data) {
  const logCtx = "Webhook:Status";
  LogService.info("Discord User requested quick status check.", logCtx);

  if (typeof buildContext !== 'function') {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "Core Engine missing" }));
  }

  const ctx = buildContext();
  const summary = {
    netWorth: ctx.netEntityValue,
    ltv: ctx.indicators.ltv,
    cryptoLtv: ctx.indicators.cryptoLTV,
    btcPrice: ctx.market.btcPrice,
    runway: ctx.indicators.survivalRunway,
    timestamp: new Date().toISOString(),
    version: Config.VERSION
  };

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: summary
  }));
}

/**
 * [ReadOnly Debug] 遠端取得 OKX recurring / BTC fills debug 結果
 * 不寫入 Unified Assets，不觸發全系統同步，只回傳 debug JSON
 */
function handleDebugOkxRecurring(data) {
  const logCtx = "Webhook:DebugOKXRecurring";
  LogService.info("Webhook requested OKX recurring debug.", logCtx);

  if (typeof Credentials !== 'object' || typeof fetchOkxRecurringBuyDebug_ !== 'function') {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      msg: "OKX debug dependencies missing"
    }));
  }

  const creds = Credentials.get('OKX');
  const { apiKey, apiSecret, apiPassphrase } = creds || {};
  const baseUrl = 'https://www.okx.com';

  if (!apiKey || !apiSecret || !apiPassphrase) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      msg: "Missing OKX keys"
    }));
  }

  const ss = getPrimarySpreadsheetForDebug_();
  const syncRes = typeof syncOkxRecurringArtifacts_ === 'function'
    ? syncOkxRecurringArtifacts_(ss, baseUrl, apiKey, apiSecret, apiPassphrase)
    : null;

  if (!syncRes || !syncRes.success) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      debugAction: "okx_recurring",
      msg: (syncRes && syncRes.status) || "Unknown Error",
      timestamp: new Date().toISOString()
    }));
  }
  return ContentService.createTextOutput(JSON.stringify(syncRes.payload));
}

function ensureSheetExists_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const requiredHeaders = headers || [];
  if (!requiredHeaders.length) return sheet;

  const headerRange = sheet.getRange(1, 1, 1, requiredHeaders.length);
  const currentHeaders = headerRange.getValues()[0];
  const matches = requiredHeaders.every((header, idx) => String(currentHeaders[idx] || '').trim() === header);

  if (!matches) {
    headerRange.setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getPrimarySpreadsheetForDebug_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const spreadsheetId = Settings.get('PRIMARY_SPREADSHEET_ID');
  if (!spreadsheetId) return null;
  return SpreadsheetApp.openById(spreadsheetId);
}

function getOkxDcaDebugHeaders_() {
  return [
    'timestamp',
    'rowType',
    'method',
    'instId',
    'buyCount',
    'totalBoughtBtc',
    'totalInvestedUsdt',
    'derivedAvgPrice',
    'pendingCount',
    'subOrdersCount',
    'historyCount',
    'lastFillBillId',
    'lastFillTime',
    'incrementalBuyCount',
    'incrementalBoughtBtc',
    'incrementalInvestedUsdt',
    'syncMode',
    'rawMessage'
  ];
}

function readOkxDcaDebugState_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return {};
  const headers = getOkxDcaDebugHeaders_();
  const row = sheet.getRange(2, 1, 1, headers.length).getValues()[0];
  const state = {};
  headers.forEach((header, idx) => {
    state[header] = row[idx];
  });
  if (String(state.rowType || '').toUpperCase() !== 'STATE') return {};
  return state;
}

function writeOkxDcaDebugArtifacts_(payload, debugPayload, providedSpreadsheet) {
  try {
    const ss = providedSpreadsheet || getPrimarySpreadsheetForDebug_();
    if (!ss) {
      LogService.error('No spreadsheet context for OKX_DCA_Debug write. Set PRIMARY_SPREADSHEET_ID or use a bound execution context.', 'Webhook:DebugOKXRecurring');
      return;
    }

    const sheet = ss.getSheetByName('OKX_DCA_Debug');
    if (!sheet) {
      LogService.error('Sheet "OKX_DCA_Debug" not found. Skipping debug snapshot write.', 'Webhook:DebugOKXRecurring');
      return;
    }

    const headers = getOkxDcaDebugHeaders_();
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    const currentHeaders = headerRange.getValues()[0];
    const matches = headers.every((header, idx) => String(currentHeaders[idx] || '').trim() === header);
    if (!matches) {
      headerRange.setValues([headers]);
      sheet.setFrozenRows(1);
    }

    const summary = payload && payload.summary ? payload.summary : {};
    const preview = payload && payload.preview ? payload.preview : {};
    const firstFill = preview.firstFill || {};
    const syncMode = summary.syncMode || '';
    const stateRow = [[
      payload.timestamp || new Date().toISOString(),
      'STATE',
      payload.method || '',
      summary.instId || '',
      preview.buyCount || summary.incrementalBuyCount || '',
      summary.totalBoughtBtc || '',
      summary.totalInvestedUsdt || '',
      summary.derivedAvgPrice || '',
      preview.pendingCount || '',
      preview.subOrdersCount || '',
      preview.historyCount || '',
      summary.lastFillBillId || firstFill.billId || '',
      summary.lastFillTime || firstFill.fillTime || firstFill.ts || '',
      summary.incrementalBuyCount || '',
      summary.incrementalBoughtBtc || '',
      summary.incrementalInvestedUsdt || '',
      syncMode,
      payload.rawMessage || ''
    ]];

    sheet.getRange(2, 1, 1, headers.length).setValues(stateRow);

    const snapshotRow = [[
      payload.timestamp || new Date().toISOString(),
      'SNAPSHOT',
      payload.method || '',
      summary.instId || '',
      preview.buyCount || summary.incrementalBuyCount || '',
      summary.totalBoughtBtc || '',
      summary.totalInvestedUsdt || '',
      summary.derivedAvgPrice || '',
      preview.pendingCount || '',
      preview.subOrdersCount || '',
      preview.historyCount || '',
      summary.lastFillBillId || firstFill.billId || '',
      summary.lastFillTime || firstFill.fillTime || firstFill.ts || '',
      summary.incrementalBuyCount || '',
      summary.incrementalBoughtBtc || '',
      summary.incrementalInvestedUsdt || '',
      syncMode,
      payload.rawMessage || ''
    ]];

    const appendRowIndex = Math.max(sheet.getLastRow() + 1, 3);
    sheet.getRange(appendRowIndex, 1, 1, headers.length).setValues(snapshotRow);
  } catch (e) {
    LogService.error(`Failed to write OKX_DCA_Debug row: ${e.message || e}`, 'Webhook:DebugOKXRecurring');
  }
}

function buildOkxRecurringResponsePayload_(debugRes, debugPayload) {
  const payload = {
    status: "success",
    debugAction: "okx_recurring",
    method: "",
    summary: {},
    rawMessage: debugRes && debugRes.message ? debugRes.message : "",
    timestamp: new Date().toISOString()
  };

  if (debugPayload) {
    payload.method = debugPayload.method || "";
    payload.summary = debugPayload.derivedSummary || {};
    payload.preview = {
      pendingCount: debugPayload.pendingCount,
      subOrdersCount: debugPayload.subOrdersCount,
      historyCount: debugPayload.historyCount,
      buyCount: debugPayload.buyCount,
      firstFill: debugPayload.firstFill || null,
      firstPending: debugPayload.firstPending || null
    };
  }

  return payload;
}

function syncOkxRecurringArtifacts_(ss, baseUrl, apiKey, apiSecret, apiPassphrase) {
  const spreadsheet = ss || getPrimarySpreadsheetForDebug_();
  const debugSheet = spreadsheet ? spreadsheet.getSheetByName('OKX_DCA_Debug') : null;
  const debugState = debugSheet ? readOkxDcaDebugState_(debugSheet) : {};

  const debugRes = typeof fetchOkxSpotBtcFillIncrementalDebug_ === 'function'
    ? fetchOkxSpotBtcFillIncrementalDebug_(baseUrl, apiKey, apiSecret, apiPassphrase, debugState)
    : fetchOkxRecurringBuyDebug_(baseUrl, apiKey, apiSecret, apiPassphrase);

  if (!debugRes.success) {
    return {
      success: false,
      status: debugRes.status || "Unknown Error"
    };
  }

  const debugPayload = typeof getLastOkxRecurringDebugPayload_ === 'function'
    ? getLastOkxRecurringDebugPayload_()
    : null;
  const payload = buildOkxRecurringResponsePayload_(debugRes, debugPayload);

  writeOkxDcaDebugArtifacts_(payload, debugPayload, spreadsheet);
  const exportWriteRes = writeOkxDcaSummaryExport_(payload, spreadsheet);

  return {
    success: true,
    rowCount: debugRes.rowCount || 0,
    exportRowCount: exportWriteRes && exportWriteRes.rowCount ? exportWriteRes.rowCount : 0,
    message: debugRes.message || "",
    payload: payload
  };
}

function writeOkxDcaSummaryExport_(payload, providedSpreadsheet) {
  try {
    const ss = providedSpreadsheet || getPrimarySpreadsheetForDebug_();
    if (!ss) return { success: false, status: 'No spreadsheet context' };

    const sheet = getApiSummaryExportSheet_(ss);
    if (!sheet) {
      LogService.warn('API_Summary_Export not found. Skipping OKX DCA summary export write.', 'Webhook:DebugOKXRecurring');
      return { success: false, status: 'Summary export sheet not found' };
    }

    const entries = buildOkxDcaSummaryExportEntries_(payload);
    if (!entries.length) return { success: true, rowCount: 0 };

    upsertApiSummaryExportEntries_(sheet, entries);
    writeAggregateBtcSpotCostSummary_(ss, sheet, entries);
    return { success: true, rowCount: entries.length };
  } catch (e) {
    LogService.error(`Failed to write API_Summary_Export rows for OKX DCA: ${e.message || e}`, 'Webhook:DebugOKXRecurring');
    return { success: false, status: e.message || String(e) };
  }
}

function buildOkxDcaSummaryExportEntries_(payload) {
  const summary = payload && payload.summary ? payload.summary : {};
  const timestamp = payload && payload.timestamp ? payload.timestamp : new Date().toISOString();
  const method = payload && payload.method ? payload.method : '';
  const source = method || 'okx_dca_sync';
  const syncMode = summary.syncMode || '';
  const preview = payload && payload.preview ? payload.preview : {};
  const cumulativeBuyCount = preview.buyCount !== undefined && preview.buyCount !== null && preview.buyCount !== ''
    ? preview.buyCount
    : '';
  const incrementalBuyCount = summary.incrementalBuyCount !== undefined && summary.incrementalBuyCount !== null && summary.incrementalBuyCount !== ''
    ? summary.incrementalBuyCount
    : '';
  const entries = [
    { key: 'OKX_BTC_DCA_BuyCount', value: cumulativeBuyCount, source: source, asOf: timestamp, note: 'Cumulative OKX BTC-USDT buy fill count' },
    { key: 'OKX_BTC_DCA_IncrementalBuyCount', value: incrementalBuyCount, source: source, asOf: timestamp, note: syncMode ? `syncMode=${syncMode}` : '' },
    { key: 'OKX_BTC_DCA_TotalBought_BTC', value: summary.totalBoughtBtc || '', source: source, asOf: timestamp, note: 'Cumulative BTC bought from OKX BTC-USDT spot fills' },
    { key: 'OKX_BTC_DCA_TotalInvested_USDT', value: summary.totalInvestedUsdt || '', source: source, asOf: timestamp, note: 'Cumulative USDT invested from OKX BTC-USDT spot fills' },
    { key: 'OKX_BTC_DCA_DerivedAvgPrice', value: summary.derivedAvgPrice || '', source: source, asOf: timestamp, note: 'Derived avg cost in USDT/BTC from OKX BTC-USDT spot fills' },
    { key: 'OKX_BTC_DCA_LastFillBillId', value: summary.lastFillBillId || '', source: source, asOf: timestamp, note: 'Checkpoint for incremental OKX DCA sync' },
    { key: 'OKX_BTC_DCA_LastFillTime', value: summary.lastFillTime || '', source: source, asOf: timestamp, note: 'OKX fill timestamp (ms)' },
    { key: 'OKX_BTC_DCA_SyncMode', value: syncMode, source: source, asOf: timestamp, note: payload && payload.rawMessage ? payload.rawMessage : '' }
  ];

  return entries;
}

function upsertApiSummaryExportEntries_(sheet, entries) {
  const maxColumns = Math.max(sheet.getLastColumn(), 5);
  const existingValues = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), maxColumns).getValues();
  const headerRow = existingValues[0] || [];
  const headerIndexes = mapSummaryExportHeaderIndexes_(headerRow);
  const rowsByKey = {};

  for (let rowIndex = 1; rowIndex < existingValues.length; rowIndex++) {
    const existingKey = String(existingValues[rowIndex][headerIndexes.key] || '').trim();
    if (!existingKey) continue;
    rowsByKey[existingKey] = rowIndex + 1;
  }

  const rowsToWrite = [];
  entries.forEach(entry => {
    const targetRow = rowsByKey[entry.key] || Math.max(sheet.getLastRow() + rowsToWrite.length + 1, 2);
    const row = new Array(maxColumns).fill('');

    if (rowsByKey[entry.key]) {
      const existingRowValues = sheet.getRange(targetRow, 1, 1, maxColumns).getValues()[0];
      for (let i = 0; i < maxColumns; i++) row[i] = existingRowValues[i];
    }

    row[headerIndexes.key] = entry.key;
    row[headerIndexes.value] = entry.value;
    if (headerIndexes.source >= 0) row[headerIndexes.source] = entry.source || '';
    if (headerIndexes.asOf >= 0) row[headerIndexes.asOf] = entry.asOf || '';
    if (headerIndexes.note >= 0) row[headerIndexes.note] = entry.note || '';

    rowsToWrite.push({ rowIndex: targetRow, values: row });
  });

  rowsToWrite.forEach(item => {
    sheet.getRange(item.rowIndex, 1, 1, maxColumns).setValues([item.values]);
  });
}

function mapSummaryExportHeaderIndexes_(headerRow) {
  const normalizedHeaders = (headerRow || []).map(value => String(value || '').trim().toLowerCase());
  const keyIndex = normalizedHeaders.indexOf('key');
  const valueIndex = normalizedHeaders.indexOf('value');

  if (keyIndex < 0 || valueIndex < 0) {
    throw new Error('API_Summary_Export missing required key/value headers');
  }

  return {
    key: keyIndex,
    value: valueIndex,
    source: normalizedHeaders.indexOf('source'),
    asOf: normalizedHeaders.indexOf('asof'),
    note: normalizedHeaders.indexOf('note')
  };
}

function getApiSummaryExportSheet_(ss) {
  const spreadsheet = ss || getPrimarySpreadsheetForDebug_();
  if (!spreadsheet) return null;

  const sheetName = (Config && Config.SHEET_NAMES && Config.SHEET_NAMES.INVENTORY_EXPORT_SUMMARY)
    ? Config.SHEET_NAMES.INVENTORY_EXPORT_SUMMARY
    : 'API_Summary_Export';
  return spreadsheet.getSheetByName(sheetName);
}

function buildApiSummaryExportEntryMap_(sheet, pendingEntries) {
  const entryMap = {};
  if (sheet) {
    const maxColumns = Math.max(sheet.getLastColumn(), 5);
    const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), maxColumns).getValues();
    const headerIndexes = mapSummaryExportHeaderIndexes_(values[0] || []);

    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const key = String(row[headerIndexes.key] || '').trim();
      if (!key) continue;
      entryMap[key] = {
        key: key,
        value: row[headerIndexes.value],
        source: headerIndexes.source >= 0 ? row[headerIndexes.source] : '',
        asOf: headerIndexes.asOf >= 0 ? row[headerIndexes.asOf] : '',
        note: headerIndexes.note >= 0 ? row[headerIndexes.note] : ''
      };
    }
  }

  (pendingEntries || []).forEach(entry => {
    if (!entry || !entry.key) return;
    entryMap[entry.key] = entry;
  });

  return entryMap;
}

function getApiSummaryExportValue_(entryMap, key) {
  const entry = entryMap && entryMap[key];
  return entry ? entry.value : '';
}

function readApiSummaryExportState_(sheet, keyPrefix) {
  const entryMap = buildApiSummaryExportEntryMap_(sheet);
  const prefix = String(keyPrefix || '');
  const state = {};

  Object.keys(entryMap).forEach(key => {
    if (!prefix || key.indexOf(prefix) !== 0) return;
    state[key.slice(prefix.length)] = entryMap[key].value;
  });

  return state;
}

function buildSummaryExportEntry_(key, value, source, asOf, note) {
  return {
    key: key,
    value: value,
    source: source || '',
    asOf: asOf || '',
    note: note || ''
  };
}

function parseSummaryNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

function writeAggregateBtcSpotCostSummary_(ss, sheet, pendingEntries) {
  const spreadsheet = ss || getPrimarySpreadsheetForDebug_();
  const exportSheet = sheet || getApiSummaryExportSheet_(spreadsheet);
  if (!spreadsheet || !exportSheet) return { success: false, status: 'Summary export sheet not found' };

  const entryMap = buildApiSummaryExportEntryMap_(exportSheet, pendingEntries);
  const usdTwd = lookupUsdTwdRateOrFallback_(spreadsheet);

  const okxBought = parseSummaryNumber_(getApiSummaryExportValue_(entryMap, 'OKX_BTC_DCA_TotalBought_BTC'));
  const okxCostTwd = parseSummaryNumber_(getApiSummaryExportValue_(entryMap, 'OKX_BTC_DCA_TotalInvested_USDT')) * usdTwd;

  const binanceBought = parseSummaryNumber_(getApiSummaryExportValue_(entryMap, 'BINANCE_BTC_Spot_TotalBought_BTC'));
  const binanceCostTwd = parseSummaryNumber_(getApiSummaryExportValue_(entryMap, 'BINANCE_BTC_Spot_TotalInvested_USDLike')) * usdTwd;

  const bitoproBought = parseSummaryNumber_(getApiSummaryExportValue_(entryMap, 'BITOPRO_BTC_Spot_TotalBought_BTC'));
  const bitoproCostTwd = parseSummaryNumber_(getApiSummaryExportValue_(entryMap, 'BITOPRO_BTC_Spot_TotalInvested_TWD'));

  const totalBought = okxBought + binanceBought + bitoproBought;
  const totalCostTwd = okxCostTwd + binanceCostTwd + bitoproCostTwd;
  const derivedAvgCostTwd = totalBought > 0 ? (totalCostTwd / totalBought) : '';
  const timestamp = new Date().toISOString();
  const coverageNote = buildAllBtcCostCoverageNote_(entryMap);

  const aggregateEntries = [
    buildSummaryExportEntry_('ALL_BTC_Spot_TotalBought_BTC', totalBought || '', 'btc_cost_aggregate', timestamp, coverageNote),
    buildSummaryExportEntry_('ALL_BTC_Spot_TotalCost_TWD', totalCostTwd || '', 'btc_cost_aggregate', timestamp, `USD/TWD=${usdTwd}`),
    buildSummaryExportEntry_('ALL_BTC_Spot_DerivedAvgCost_TWD', derivedAvgCostTwd || '', 'btc_cost_aggregate', timestamp, 'Cross-exchange derived BTC spot cost in TWD'),
    buildSummaryExportEntry_('ALL_BTC_Spot_USDT_TWD_Rate', usdTwd || '', 'btc_cost_aggregate', timestamp, 'USDT treated as USD for BTC spot cost normalization')
  ];

  upsertApiSummaryExportEntries_(exportSheet, aggregateEntries);
  return { success: true, rowCount: aggregateEntries.length };
}

function buildAllBtcCostCoverageNote_(entryMap) {
  const notes = [];
  const bitoproWindow = getApiSummaryExportValue_(entryMap, 'BITOPRO_BTC_Spot_HistoryWindowDays');
  if (bitoproWindow) {
    notes.push(`BitoPro limited to latest ${bitoproWindow} days via API`);
  }
  const binanceSymbols = getApiSummaryExportValue_(entryMap, 'BINANCE_BTC_Spot_Symbols');
  if (binanceSymbols) {
    notes.push(`Binance symbols=${binanceSymbols}`);
  }
  return notes.join('; ');
}

function lookupUsdTwdRateOrFallback_(ss) {
  try {
    if (typeof SettingsMatrixRepo !== 'undefined' && SettingsMatrixRepo.lookupRate) {
      const lookup = SettingsMatrixRepo.lookupRate(ss, 'USD', 'TWD');
      const rate = parseSummaryNumber_(lookup && lookup.value);
      if (rate > 0) return rate;
    }
  } catch (e) { }

  return 32;
}
