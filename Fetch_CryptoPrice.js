// =================================================================
// == 自動化價格更新系統                                         ==
// == 1. updateAllPrices: 價格目錄 discovery + 安全更新 orchestrator ==
// == 2. fetchCryptoPrice: 數位幣價格調度中心                     ==
// == 3. fetchStockPrice: 股票價格調度中心                         ==
// =================================================================

const PRICE_UPDATE_MODULE_NAME = "Fetch_CryptoPrice";
const PRICE_UPDATE_LOCK_TIMEOUT_MS = 5000;
const PRICE_ASSET_ALLOCATION_SHEET_NAME = "資產配置";
const PRICE_DISCOVERY_HEADER_SCAN_ROWS = 20;
const PRICE_DISCOVERY_TICKER_ALIASES = ["股票代號/銀行", "股票代號 / 銀行", "股票代號", "Ticker", "Symbol"];
const PRICE_DISCOVERY_TYPE_ALIASES = ["類型", "資產類型", "Asset Type"];
const PRICE_DISCOVERY_ALLOWED_TYPES = ["股票", "數位幣", "數位穩定幣"];
const PRICE_LOCK_FUTURE_TOLERANCE_MS = 1000;

// =======================================================
// 主更新函式 - 請將此函式設定為每15或30分鐘觸發一次
// =======================================================
function updateAllPrices(options) {
  const opts = options || {};
  const summary = createPriceUpdateSummary_();
  const attemptStartedAt = new Date();
  let lock = null;

  logPriceUpdate_("info", "Starting Market Price Update...");

  try {
    const lockResult = acquirePriceUpdateLock_(opts.lockTimeoutMs || PRICE_UPDATE_LOCK_TIMEOUT_MS);
    if (!lockResult.acquired) {
      summary.status = "SKIPPED_LOCKED";
      summary.ok = false;
      summary.fatal = false;
      summary.message = lockResult.message || "Another price update is already running.";
      logPriceUpdate_("warn", `Price update skipped: ${summary.message}`);
      return summary;
    }

    lock = lockResult.lock;
    const ss = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const priceCacheOptions = { sheetName: opts.priceCacheSheetName };
    let rows = PriceCacheRepo.readRows(ss, priceCacheOptions);
    const discovery = syncPriceCatalogFromAssetAllocation_(ss, rows, opts);
    rows = rows.concat(discovery.addedRows || []);
    summary.added = discovery.addedCount || 0;
    summary.discoveryWarnings = discovery.warnings || [];

    const cacheDurationMinutes = (Config && Config.THRESHOLDS && Config.THRESHOLDS.PRICE_CACHE_MINUTES) || 1;
    const cacheDurationMs = cacheDurationMinutes * 60 * 1000;

    if (rows.length === 0) {
      summary.status = "COMPLETE";
      summary.ok = true;
      summary.message = "Price cache sheet is empty and no asset allocation candidates were discovered.";
      logPriceUpdate_("info", summary.message);
      return summary;
    }

    const currentTime = new Date();
    const updates = [];

    rows.forEach(row => {
      const ticker = normalizePriceTicker_(row.ticker, row.assetType);
      if (!ticker) {
        summary.skippedInvalid++;
        return;
      }

      if (isPriceCacheRowLocked_(row, currentTime)) {
        summary.locked++;
        logPriceUpdate_("info", `Price row locked by future updatedAt: ${ticker}`);
        return;
      }

      if (!isPriceCacheRowDueForUpdate_(row, currentTime, cacheDurationMs)) {
        summary.fresh++;
        return;
      }

      const price = fetchPriceForCatalogRow_(row, { skipTempLock: true });
      if (isPositivePrice_(price)) {
        updates.push({
          rowNumber: row.rowNumber,
          ticker: ticker,
          expectedPrice: row.price,
          expectedUpdatedAt: row.updatedAt,
          price: price,
          updatedAt: currentTime
        });
      } else {
        summary.failed++;
        summary.failedTickers.push(ticker);
        logPriceUpdate_("warn", `[FAILED] 更新 ${ticker} 的價格失敗。`);
      }
    });

    const safeUpdates = filterSafePriceUpdatesBeforeWrite_(ss, updates, priceCacheOptions, attemptStartedAt, summary);
    if (safeUpdates.length > 0) {
      PriceCacheRepo.writePriceUpdates(ss, safeUpdates, priceCacheOptions);
    }
    summary.updated = safeUpdates.length;

    summary.status = (summary.failed > 0 || summary.skippedChanged > 0) ? "WARNING" : "COMPLETE";
    summary.ok = summary.failed === 0 && summary.skippedChanged === 0;
    summary.fatal = false;
    summary.message = buildPriceUpdateSummaryMessage_(summary);
    logPriceUpdate_(summary.status === "WARNING" ? "warn" : "info", summary.message);
    return summary;
  } catch (e) {
    summary.status = "FAILED";
    summary.ok = false;
    summary.fatal = true;
    summary.message = e.message || String(e);
    logPriceUpdate_("error", `Price Update Failed: ${summary.message}`);
    return summary;
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (e) {
        logPriceUpdate_("warn", `Price update lock release failed: ${e.message || e}`);
      }
    }
  }
}

function createPriceUpdateSummary_() {
  return {
    status: "RUNNING",
    ok: false,
    fatal: false,
    added: 0,
    updated: 0,
    failed: 0,
    locked: 0,
    fresh: 0,
    skippedChanged: 0,
    skippedInvalid: 0,
    failedTickers: [],
    discoveryWarnings: [],
    message: ""
  };
}

function buildPriceUpdateSummaryMessage_(summary) {
  const parts = [
    `Added: ${summary.added}`,
    `Updated: ${summary.updated}`,
    `Failed: ${summary.failed}`,
    `Locked: ${summary.locked}`,
    `Fresh: ${summary.fresh}`
  ];

  if (summary.skippedChanged > 0) parts.push(`Changed Before Write: ${summary.skippedChanged}`);
  if (summary.skippedInvalid > 0) parts.push(`Invalid: ${summary.skippedInvalid}`);
  if (summary.discoveryWarnings && summary.discoveryWarnings.length > 0) {
    parts.push(`Discovery Warnings: ${summary.discoveryWarnings.length}`);
  }
  if (summary.failedTickers.length > 0) {
    parts.push(`Failed Tickers: ${summary.failedTickers.join(", ")}`);
  }

  return `Price Update Completed. ${parts.join(", ")}`;
}

function acquirePriceUpdateLock_(timeoutMs) {
  try {
    let lock = null;
    if (typeof LockService !== "undefined" && LockService.getDocumentLock) {
      lock = LockService.getDocumentLock();
    }
    if (!lock && typeof LockService !== "undefined" && LockService.getScriptLock) {
      lock = LockService.getScriptLock();
    }
    if (!lock) {
      return {
        acquired: false,
        unavailable: true,
        lock: null,
        message: "LockService did not provide a document/script lock."
      };
    }

    if (lock.tryLock(timeoutMs || PRICE_UPDATE_LOCK_TIMEOUT_MS)) {
      return {
        acquired: true,
        unavailable: false,
        lock: lock,
        message: "OK"
      };
    }

    return {
      acquired: false,
      unavailable: false,
      lock: null,
      message: "Price update lock is currently held by another execution."
    };
  } catch (e) {
    return {
      acquired: false,
      unavailable: true,
      lock: null,
      message: `Lock unavailable: ${e.message || e}`
    };
  }
}

// Kept for tests and future callers that only need the lock object.
function tryAcquirePriceUpdateLock_(timeoutMs) {
  const result = acquirePriceUpdateLock_(timeoutMs);
  return result.acquired ? result.lock : null;
}

function syncPriceCatalogFromAssetAllocation_(ss, rows, options) {
  const opts = options || {};
  const warnings = [];
  const existingLookup = {};
  const additions = [];

  (rows || []).forEach(row => {
    const ticker = normalizePriceTicker_(row.ticker, row.assetType);
    if (ticker) existingLookup[ticker] = true;
  });

  const candidates = readAssetAllocationPriceCandidates_(ss, opts);
  candidates.warnings.forEach(message => warnings.push(message));

  candidates.rows.forEach(candidate => {
    if (existingLookup[candidate.ticker]) return;
    existingLookup[candidate.ticker] = true;
    additions.push({
      ticker: candidate.ticker,
      assetType: candidate.assetType,
      price: "",
      updatedAt: ""
    });
  });

  if (additions.length === 0) {
    return {
      addedCount: 0,
      addedRows: [],
      warnings: warnings
    };
  }

  const appendResult = PriceCacheRepo.appendRows(ss, additions, { sheetName: opts.priceCacheSheetName });
  appendResult.rows.forEach(row => {
    logPriceUpdate_("info", `Added price catalog row: ${row.ticker} (${row.assetType})`);
  });

  return {
    addedCount: appendResult.rowCount,
    addedRows: appendResult.rows,
    warnings: warnings
  };
}

function filterSafePriceUpdatesBeforeWrite_(ss, updates, priceCacheOptions, attemptStartedAt, summary) {
  if (!updates || updates.length === 0) return [];

  const latestRows = PriceCacheRepo.readRows(ss, priceCacheOptions);
  const rowLookup = {};
  latestRows.forEach(row => {
    rowLookup[row.rowNumber] = row;
  });

  const now = new Date();
  return updates.filter(update => {
    const latestRow = rowLookup[update.rowNumber];
    if (!latestRow) {
      summary.skippedChanged++;
      logPriceUpdate_("warn", `Skipped price write for ${update.ticker}: row no longer exists.`);
      return false;
    }

    const latestTicker = normalizePriceTicker_(latestRow.ticker, latestRow.assetType);
    if (latestTicker !== update.ticker) {
      summary.skippedChanged++;
      logPriceUpdate_("warn", `Skipped price write for ${update.ticker}: row ticker changed to ${latestTicker || "(blank)"}.`);
      return false;
    }

    if (isPriceCacheRowLocked_(latestRow, now)) {
      summary.skippedChanged++;
      logPriceUpdate_("warn", `Skipped price write for ${update.ticker}: row was manually locked before write.`);
      return false;
    }

    if (!priceCacheCellValuesMatch_(latestRow.price, update.expectedPrice) ||
      !priceCacheCellValuesMatch_(latestRow.updatedAt, update.expectedUpdatedAt)) {
      summary.skippedChanged++;
      logPriceUpdate_("warn", `Skipped price write for ${update.ticker}: row price/timestamp changed before write.`);
      return false;
    }

    const latestUpdatedAt = parsePriceCacheDate_(latestRow.updatedAt);
    if (latestUpdatedAt && latestUpdatedAt.getTime() > attemptStartedAt.getTime()) {
      summary.skippedChanged++;
      logPriceUpdate_("warn", `Skipped price write for ${update.ticker}: row was updated after this run started.`);
      return false;
    }

    return true;
  });
}

function readAssetAllocationPriceCandidates_(ss, options) {
  const opts = options || {};
  const sheetName = opts.assetAllocationSheetName || PRICE_ASSET_ALLOCATION_SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);
  const warnings = [];

  if (!sheet) {
    const message = `Asset allocation sheet "${sheetName}" not found. Price discovery skipped.`;
    logPriceUpdate_("warn", message);
    return { rows: [], warnings: [message] };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 0 || lastCol <= 0) {
    return { rows: [], warnings: [] };
  }

  const scanRows = Math.min(Math.max(opts.headerScanRows || PRICE_DISCOVERY_HEADER_SCAN_ROWS, 1), lastRow);
  const headerValues = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const headerMatch = findAssetAllocationPriceHeader_(headerValues);

  if (!headerMatch) {
    const message = `Asset allocation sheet "${sheetName}" header not found. Price discovery skipped.`;
    logPriceUpdate_("warn", message);
    return { rows: [], warnings: [message] };
  }

  const dataStartRow = headerMatch.headerRowIndex + 1;
  const dataRowCount = Math.max(lastRow - headerMatch.headerRowIndex, 0);
  if (dataRowCount <= 0) {
    return { rows: [], warnings: [] };
  }

  const values = sheet.getRange(dataStartRow, 1, dataRowCount, lastCol).getValues();
  const lookup = {};
  const candidates = [];

  values.forEach(row => {
    const rawAssetType = row[headerMatch.assetTypeColumnIndex - 1];
    const assetType = normalizePriceAssetType_(rawAssetType);
    if (!isPriceEligibleAssetType_(assetType)) return;

    const ticker = normalizePriceTicker_(row[headerMatch.tickerColumnIndex - 1], assetType);
    if (!ticker || lookup[ticker]) return;

    lookup[ticker] = true;
    candidates.push({
      ticker: ticker,
      assetType: assetType
    });
  });

  return {
    rows: candidates,
    warnings: warnings
  };
}

function findAssetAllocationPriceHeader_(values) {
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex] || [];
    const tickerColumnIndex = findPriceHeaderColumnIndex_(row, PRICE_DISCOVERY_TICKER_ALIASES, "first");
    if (tickerColumnIndex < 0) continue;

    const typeIndexes = findAllPriceHeaderColumnIndexes_(row, PRICE_DISCOVERY_TYPE_ALIASES);
    if (typeIndexes.length === 0) continue;

    let assetTypeColumnIndex = -1;
    for (let i = typeIndexes.length - 1; i >= 0; i--) {
      if (typeIndexes[i] < tickerColumnIndex) {
        assetTypeColumnIndex = typeIndexes[i];
        break;
      }
    }

    if (assetTypeColumnIndex < 0) {
      assetTypeColumnIndex = typeIndexes[0];
    }

    return {
      headerRowIndex: rowIndex + 1,
      tickerColumnIndex: tickerColumnIndex + 1,
      assetTypeColumnIndex: assetTypeColumnIndex + 1
    };
  }

  return null;
}

function findPriceHeaderColumnIndex_(row, aliases, matchMode) {
  const indexes = findAllPriceHeaderColumnIndexes_(row, aliases);
  if (indexes.length === 0) return -1;
  return matchMode === "last" ? indexes[indexes.length - 1] : indexes[0];
}

function findAllPriceHeaderColumnIndexes_(row, aliases) {
  const normalizedAliases = (aliases || []).map(normalizePriceHeaderToken_).filter(Boolean);
  const indexes = [];

  (row || []).forEach((value, index) => {
    const normalizedValue = normalizePriceHeaderToken_(value);
    if (!normalizedValue) return;

    const matched = normalizedAliases.some(alias => {
      return normalizedValue === alias || normalizedValue.indexOf(alias) >= 0 || alias.indexOf(normalizedValue) >= 0;
    });

    if (matched) indexes.push(index);
  });

  return indexes;
}

function normalizePriceHeaderToken_(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/[\s\u3000_\-:/()（）\[\]【】]/g, "");
}

function normalizePriceAssetType_(value) {
  return String(value == null ? "" : value).trim();
}

function isPriceEligibleAssetType_(assetType) {
  return PRICE_DISCOVERY_ALLOWED_TYPES.indexOf(normalizePriceAssetType_(assetType)) >= 0;
}

function normalizePriceTicker_(value, assetType) {
  let text = String(value == null ? "" : value).trim();
  if (!text) return "";

  text = text.replace(/\s+/g, "");
  if (normalizePriceAssetType_(assetType) === "股票" && /^\d+$/.test(text) && text.length < 4) {
    while (text.length < 4) text = `0${text}`;
  }

  return text.toUpperCase();
}

function isPriceCacheRowLocked_(row, now) {
  const updatedAt = parsePriceCacheDate_(row && row.updatedAt);
  if (!updatedAt) return false;
  return updatedAt.getTime() > (now || new Date()).getTime() + PRICE_LOCK_FUTURE_TOLERANCE_MS;
}

function isPriceCacheRowDueForUpdate_(row, now, cacheDurationMs) {
  const updatedAt = parsePriceCacheDate_(row && row.updatedAt);
  if (!updatedAt) return true;
  return (now || new Date()).getTime() - updatedAt.getTime() > cacheDurationMs;
}

function parsePriceCacheDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value;
  }

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function priceCacheCellValuesMatch_(currentValue, expectedValue) {
  if (!currentValue && !expectedValue) return true;

  if (isPriceCacheNumericLiteral_(currentValue) || isPriceCacheNumericLiteral_(expectedValue)) {
    const currentNumber = parseFloat(currentValue);
    const expectedNumber = parseFloat(expectedValue);
    return isFinite(currentNumber) && isFinite(expectedNumber) && Math.abs(currentNumber - expectedNumber) < 1e-12;
  }

  const currentDate = parsePriceCacheDate_(currentValue);
  const expectedDate = parsePriceCacheDate_(expectedValue);
  if (currentDate || expectedDate) {
    return !!currentDate && !!expectedDate && currentDate.getTime() === expectedDate.getTime();
  }

  return String(currentValue == null ? "" : currentValue).trim() === String(expectedValue == null ? "" : expectedValue).trim();
}

function isPriceCacheNumericLiteral_(value) {
  if (typeof value === "number") return isFinite(value);
  return /^-?\d+(\.\d+)?$/.test(String(value == null ? "" : value).trim());
}

function fetchPriceForCatalogRow_(row, options) {
  const opts = options || {};
  const ticker = normalizePriceTicker_(row && row.ticker, row && row.assetType);
  const assetType = normalizePriceAssetType_(row && row.assetType);

  if (!ticker) return null;

  if (assetType === "數位幣" || assetType === "數位穩定幣") {
    return fetchCryptoPrice(ticker);
  }

  if (assetType === "股票") {
    return fetchStockPrice(ticker, false, opts);
  }

  logPriceUpdate_("warn", `Unknown price asset type for ${ticker}: ${assetType || "(blank)"}. Trying stock then crypto.`);
  const stockPrice = fetchStockPrice(ticker, false, opts);
  if (isPositivePrice_(stockPrice)) return stockPrice;
  return fetchCryptoPrice(ticker);
}

function isPositivePrice_(value) {
  const num = parseFloat(value);
  return isFinite(num) && num > 0;
}

// =======================================================
// == 輔助函式 (Helper Functions)
// =======================================================

/**
 * 數位幣價格調度中心
 * 依次嘗試多個數據來源，直到成功獲取價格為止。
 */
function fetchCryptoPrice(ticker, currency = "USD", bypassCache = false) {
  const normalizedTicker = normalizePriceTicker_(ticker);
  const normalizedCurrency = String(currency || "USD").trim().toUpperCase();
  if (!normalizedTicker) return null;

  const cacheKey = `PRICE_CRYPTO_${normalizedTicker}_${normalizedCurrency}`;

  if (normalizedTicker === "BITO" && typeof fetchBitoProBitoPrice_ === "function") {
    const bitoPrice = fetchBitoProBitoPrice_(normalizedCurrency, bypassCache);
    if (isPositivePrice_(bitoPrice)) {
      ScriptCache.put(cacheKey, bitoPrice, 300);
      Logger.log(`  - [Source 0: BitoPro] OK for ${normalizedTicker}: ${bitoPrice}`);
      return bitoPrice;
    }
  }

  // 1. Check Cache (unless bypass is requested)
  if (!bypassCache) {
    const cached = ScriptCache.get(cacheKey);
    if (isPositivePrice_(cached)) return parseFloat(cached);
  }

  let price = null;

  // --- 策略1：嘗試從 cryptoprices.cc (快速、免費) ---
  try {
    const url = `https://cryptoprices.cc/${normalizedTicker}/`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (response.getResponseCode() === 200) {
      price = parseFloat(response.getContentText());
      if (isPositivePrice_(price)) {
        Logger.log(`  - [Source 1: cryptoprices.cc] OK for ${normalizedTicker}: ${price}`);
        ScriptCache.put(cacheKey, price, 900);
        return price;
      }
    }
  } catch (e) {
    Logger.log(`  - [Source 1: cryptoprices.cc] FAILED for ${normalizedTicker}: ${e.toString()}`);
  }

  // --- 策略2：如果策略1失敗，則呼叫 CoinMarketCap (穩定、可靠) ---
  Logger.log(`  - [Source 1 FAILED] Falling back to Source 2 (CoinMarketCap) for ${normalizedTicker}...`);
  price = getCryptoPrice(normalizedTicker, normalizedCurrency);

  if (isPositivePrice_(price)) {
    Logger.log(`  - [Source 2: CoinMarketCap] OK for ${normalizedTicker}: ${price}`);
    ScriptCache.put(cacheKey, price, 900);
    return price;
  }

  return null;
}

/**
 * [穩健版] 從 CoinMarketCap 獲取加密貨幣價格 (已重新命名)
 */
function getCryptoPrice(symbol, convert) {
  var convertCurrency = convert || "USD";
  var apiKey = Settings.get("CMC_API_KEY");

  if (!apiKey) {
    Logger.log("  - ERROR: CoinMarketCap API Key not found in Script Properties.");
    return null;
  }

  var url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=" + symbol.toUpperCase() + "&convert=" + convertCurrency.toUpperCase();
  var options = { "method": "GET", "headers": { "X-CMC_PRO_API_KEY": apiKey, "Accept": "application/json" }, "muteHttpExceptions": true };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200 && json.status.error_code === 0) {
      if (json.data && json.data[symbol.toUpperCase()]) {
        var price = json.data[symbol.toUpperCase()].quote[convertCurrency.toUpperCase()].price;
        if (isPositivePrice_(price)) {
          return price;
        }
      } else {
        Logger.log(`  - CoinMarketCap ERROR: Symbol ${symbol} not found in response.`);
      }
    } else {
      Logger.log(`  - CoinMarketCap API ERROR: ${json.status.error_message || "Request failed"}`);
    }
  } catch (e) {
    Logger.log("  - CoinMarketCap Exception: " + e.toString());
  }

  return null;
}

/**
 * 股票價格調度中心
 * 依次嘗試 GOOGLEFINANCE 和網頁爬蟲。
 */
function fetchStockPrice(ticker, bypassCache = false, options) {
  const opts = options || {};
  const normalizedTicker = normalizePriceTicker_(ticker, "股票");
  if (!normalizedTicker) return null;

  const cacheKey = `PRICE_STOCK_${normalizedTicker}`;

  if (!bypassCache) {
    const cached = ScriptCache.get(cacheKey);
    if (isPositivePrice_(cached)) return parseFloat(cached);
  }

  let price = null;

  // --- 策略1：嘗試 GOOGLEFINANCE (最穩定) ---
  let tempLockResult = { acquired: false, unavailable: false, lock: null };
  if (!opts.skipTempLock) {
    tempLockResult = acquirePriceUpdateLock_(opts.lockTimeoutMs || PRICE_UPDATE_LOCK_TIMEOUT_MS);
  }

  const canUseTemp = opts.skipTempLock || tempLockResult.acquired || tempLockResult.unavailable;
  if (canUseTemp) {
    try {
      const tempSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Temp");
      if (!tempSheet) throw new Error("找不到名為 'Temp' 的臨時工作表。");
      const tempRange = tempSheet.getRange("A1");
      tempRange.setFormula(`=GOOGLEFINANCE("${normalizedTicker}","price")`);
      SpreadsheetApp.flush();
      price = tempRange.getValue();
      if (isPositivePrice_(price)) {
        Logger.log(`  - [Source 1: GOOGLEFINANCE] OK for ${normalizedTicker}: ${price}`);
        ScriptCache.put(cacheKey, price, 3600);
        return price;
      }
    } catch (e) {
      Logger.log(`  - [Source 1: GOOGLEFINANCE] FAILED for ${normalizedTicker}: ${e.toString()}`);
    } finally {
      if (tempLockResult.lock) {
        try {
          tempLockResult.lock.releaseLock();
        } catch (e) {
          Logger.log(`  - [Source 1: GOOGLEFINANCE] lock release failed for ${normalizedTicker}: ${e.toString()}`);
        }
      }
    }
  } else {
    Logger.log(`  - [Source 1: GOOGLEFINANCE] SKIPPED for ${normalizedTicker}: ${tempLockResult.message}`);
  }

  // --- 策略2：如果策略1失敗，嘗試爬取 cnyes.com (適用台股) ---
  Logger.log(`  - [Source 1 FAILED] Falling back to Source 2 (cnyes.com) for ${normalizedTicker}...`);
  try {
    const url = `https://www.cnyes.com/twstock/${normalizedTicker}`;
    const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
    const match = html.match(/<h3 class="[^"]*">([0-9,]+\.?[0-9]*)<\/h3>/);
    if (match && match[1]) {
      price = parseFloat(match[1].replace(/,/g, ""));
      if (isPositivePrice_(price)) {
        Logger.log(`  - [Source 2: cnyes.com] OK for ${normalizedTicker}: ${price}`);
        ScriptCache.put(cacheKey, price, 3600);
        return price;
      }
    }
  } catch (e) {
    Logger.log(`  - [Source 2: cnyes.com] FAILED for ${normalizedTicker}: ${e.toString()}`);
  }

  return null;
}

function logPriceUpdate_(level, message) {
  if (typeof LogService !== "undefined" && LogService[level]) {
    LogService[level](message, PRICE_UPDATE_MODULE_NAME);
    return;
  }

  Logger.log(`[${PRICE_UPDATE_MODULE_NAME}] ${message}`);
}
