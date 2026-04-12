// =================================================================
// == FX Rate Sync for 參數設定                                   ==
// == Keeps the currency matrix fresh without changing its layout. ==
// =================================================================

const FX_UPDATE_MODULE_NAME = "Sync_FxRates";
const FX_UPDATE_LOCK_TIMEOUT_MS = 5000;
const FX_TEMP_START_ROW = 1;
const FX_TEMP_START_COLUMN = 13; // M
const FX_TEMP_OWNED_COLUMN_COUNT = 2;
const FX_TEMP_MIN_CLEAR_ROWS = 200;
const FX_RATE_CACHE_MINUTES = 60;
const FX_RATE_EPSILON = 1e-12;

function updateAllFxRates(options) {
  const opts = normalizeFxRateUpdateOptions_(options);
  const summary = createFxRateUpdateSummary_();
  const attemptStartedAt = opts.now || new Date();
  let lock = null;

  logFxRateUpdate_("info", "Starting FX Rate Update...");

  try {
    const lockResult = acquireFxRateUpdateLock_(opts.lockTimeoutMs);
    if (!lockResult.acquired) {
      summary.status = "SKIPPED_LOCKED";
      summary.ok = false;
      summary.fatal = false;
      summary.message = lockResult.message || "Another FX rate update is already running.";
      logFxRateUpdate_("warn", `FX rate update skipped: ${summary.message}`);
      return summary;
    }

    lock = lockResult.lock;
    const ss = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const settingsOptions = { sheetName: opts.settingsSheetName };
    const discoveredRequiredPairs = opts.requiredPairs || collectRequiredCurrencyPairs_(ss, {
      settingsSheetName: opts.settingsSheetName,
      sheetNames: opts.sheetNames
    }).pairs;
    const requiredPairs = normalizeFxRequiredPairs_(discoveredRequiredPairs, opts);
    const requiredPairLookup = buildFxRequiredPairLookup_(requiredPairs);

    const appendResult = SettingsMatrixRepo.appendMissingPairs(ss, requiredPairs, settingsOptions);
    appendResult.addedFromCurrencies.forEach(currency => {
      logFxRateUpdate_("info", `Added From-Currency before FX update: ${currency}`);
    });
    appendResult.addedToCurrencies.forEach(currency => {
      logFxRateUpdate_("info", `Added To-Currency before FX update: ${currency}`);
    });

    const pairResult = SettingsMatrixRepo.listPairs(ss, settingsOptions);
    const pairs = (pairResult.pairs || []).map(pair => {
      const key = buildFxPairKey_(pair.from, pair.to);
      pair.required = !!requiredPairLookup[key];
      return pair;
    });
    if (pairs.length === 0) {
      summary.status = "COMPLETE";
      summary.ok = true;
      summary.fatal = false;
      summary.message = "Settings matrix has no FX pairs to update.";
      logFxRateUpdate_("info", summary.message);
      return summary;
    }

    const cacheDurationMinutes = getFxRateCacheMinutes_();
    const cacheDurationMs = cacheDurationMinutes * 60 * 1000;
    const currentTime = opts.now || new Date();
    const duePairs = [];
    const updates = [];

    pairs.forEach(pair => {
      if (!isFxRatePairDueForUpdate_(pair, currentTime, cacheDurationMs)) {
        summary.fresh++;
        return;
      }

      if (pair.from === pair.to) {
        updates.push(createFxRateUpdate_(pair, 1, currentTime, "identity"));
        return;
      }

      duePairs.push(pair);
    });

    const fetchedRates = fetchFxRatesForPairs_(ss, duePairs, opts);
    duePairs.forEach(pair => {
      const key = buildFxPairKey_(pair.from, pair.to);
      const fetched = fetchedRates[key];
      if (fetched && isPositiveFxRate_(fetched.rate)) {
        updates.push(createFxRateUpdate_(pair, fetched.rate, currentTime, fetched.source));
        return;
      }

      summary.failed++;
      summary.failedPairs.push(`${pair.from}/${pair.to}`);
      logFxRateUpdate_("warn", `[FAILED] 更新 ${pair.from}/${pair.to} 匯率失敗。`);

      if (!isFxRatePairUsableAfterFailure_(pair, currentTime)) {
        if (!pair.required) {
          addFxOptionalUnusablePair_(summary, `${pair.from}/${pair.to}`);
          return;
        }
        addFxMissingPair_(summary, `${pair.from}/${pair.to}`);
      }
    });

    const safeUpdates = filterSafeFxRateUpdatesBeforeWrite_(ss, updates, settingsOptions, attemptStartedAt, summary);
    if (safeUpdates.length > 0) {
      SettingsMatrixRepo.writeRateUpdates(ss, safeUpdates, settingsOptions);
    }

    summary.updated = safeUpdates.length;
    summary.status = summary.missingPairs.length > 0
      ? "FAILED"
      : (summary.failed > 0 || summary.skippedChanged > 0 ? "WARNING" : "COMPLETE");
    summary.fatal = summary.missingPairs.length > 0;
    summary.ok = summary.status === "COMPLETE";
    summary.message = buildFxRateUpdateSummaryMessage_(summary);

    logFxRateUpdate_(summary.status === "COMPLETE" ? "info" : "warn", summary.message);
    if (summary.fatal && opts.throwOnFatal) throw new Error(summary.message);
    return summary;
  } catch (e) {
    summary.status = "FAILED";
    summary.ok = false;
    summary.fatal = true;
    summary.message = e.message || String(e);
    logFxRateUpdate_("error", `FX Rate Update Failed: ${summary.message}`);
    if (opts.throwOnFatal) throw e;
    return summary;
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (e) {
        logFxRateUpdate_("warn", `FX rate update lock release failed: ${e.message || e}`);
      }
    }
  }
}

function normalizeFxRateUpdateOptions_(options) {
  const opts = options || {};
  return {
    spreadsheet: opts.spreadsheet || null,
    settingsSheetName: opts.settingsSheetName || SettingsMatrixRepo.DEFAULT_SHEET_NAME,
    tempSheetName: opts.tempSheetName || "Temp",
    tempStartRow: opts.tempStartRow || FX_TEMP_START_ROW,
    tempStartColumn: opts.tempStartColumn || FX_TEMP_START_COLUMN,
    tempOwnedColumnCount: opts.tempOwnedColumnCount || FX_TEMP_OWNED_COLUMN_COUNT,
    lockTimeoutMs: opts.lockTimeoutMs || FX_UPDATE_LOCK_TIMEOUT_MS,
    now: opts.now || null,
    throwOnFatal: !!opts.throwOnFatal,
    requiredPairs: opts.requiredPairs || null,
    sheetNames: opts.sheetNames || null,
    rateProvider: opts.rateProvider || null,
    flushWaitMs: opts.flushWaitMs === undefined ? 500 : opts.flushWaitMs,
    formulaRetryCount: opts.formulaRetryCount || 3,
    formulaRetryWaitMs: opts.formulaRetryWaitMs || 1000,
    includeMandatoryPairs: opts.includeMandatoryPairs === undefined ? true : !!opts.includeMandatoryPairs
  };
}

function createFxRateUpdateSummary_() {
  return {
    status: "RUNNING",
    ok: false,
    fatal: false,
    updated: 0,
    fresh: 0,
    failed: 0,
    skippedChanged: 0,
    missingPairs: [],
    failedPairs: [],
    optionalUnusablePairs: [],
    message: ""
  };
}

function buildFxRateUpdateSummaryMessage_(summary) {
  const parts = [
    `Updated: ${summary.updated}`,
    `Fresh: ${summary.fresh}`,
    `Failed: ${summary.failed}`
  ];

  if (summary.skippedChanged > 0) parts.push(`Changed Before Write: ${summary.skippedChanged}`);
  if (summary.missingPairs.length > 0) parts.push(`Missing Usable Rates: ${summary.missingPairs.join(", ")}`);
  if (summary.optionalUnusablePairs.length > 0) parts.push(`Optional Unusable Rates: ${summary.optionalUnusablePairs.join(", ")}`);
  if (summary.failedPairs.length > 0) parts.push(`Failed Pairs: ${summary.failedPairs.join(", ")}`);

  return `FX Rate Update Completed. ${parts.join(", ")}`;
}

function acquireFxRateUpdateLock_(timeoutMs) {
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

    if (lock.tryLock(timeoutMs || FX_UPDATE_LOCK_TIMEOUT_MS)) {
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
      message: "FX rate update lock is currently held by another execution."
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

function getFxRateCacheMinutes_() {
  if (typeof Config !== "undefined" && Config.THRESHOLDS && Config.THRESHOLDS.FX_RATE_CACHE_MINUTES) {
    return Config.THRESHOLDS.FX_RATE_CACHE_MINUTES;
  }

  return FX_RATE_CACHE_MINUTES;
}

function isFxRatePairDueForUpdate_(pair, now, cacheDurationMs) {
  if (!pair) return false;
  if (!isPositiveFxRate_(pair.currentValue)) return true;

  const updatedAt = parseFxRateDate_(pair.currentTimestamp);
  if (!updatedAt) return true;

  return (now || new Date()).getTime() - updatedAt.getTime() > cacheDurationMs;
}

function createFxRateUpdate_(pair, rate, updatedAt, source) {
  return {
    from: pair.from,
    to: pair.to,
    rowNumber: pair.rowNumber,
    valueColumn: pair.valueColumn,
    timestampColumn: pair.timestampColumn,
    expectedValue: pair.currentValue,
    expectedTimestamp: pair.currentTimestamp,
    rate: rate,
    updatedAt: updatedAt,
    source: source || "",
    required: !!pair.required
  };
}

function fetchFxRatesForPairs_(ss, pairs, options) {
  const opts = options || {};
  const result = {};
  const normalizedPairs = (pairs || []).filter(pair => pair && pair.from && pair.to && pair.from !== pair.to);

  if (normalizedPairs.length === 0) return result;

  if (typeof opts.rateProvider === "function") {
    normalizedPairs.forEach(pair => {
      const direct = opts.rateProvider({
        from: pair.from,
        to: pair.to,
        mode: "direct",
        pair: pair
      });
      const directRate = typeof direct === "object" && direct !== null ? direct.rate : direct;
      const directSource = typeof direct === "object" && direct !== null ? direct.source : "provider:direct";
      if (isPositiveFxRate_(directRate)) {
        result[buildFxPairKey_(pair.from, pair.to)] = {
          rate: parseFloat(directRate),
          source: directSource || "provider:direct"
        };
        return;
      }

      const inverse = opts.rateProvider({
        from: pair.to,
        to: pair.from,
        mode: "inverse",
        pair: pair
      });
      const inverseRate = typeof inverse === "object" && inverse !== null ? inverse.rate : inverse;
      const inverseSource = typeof inverse === "object" && inverse !== null ? inverse.source : "provider:inverse";
      if (isPositiveFxRate_(inverseRate)) {
        result[buildFxPairKey_(pair.from, pair.to)] = {
          rate: 1 / parseFloat(inverseRate),
          source: inverseSource || "provider:inverse"
        };
      }
    });
    return result;
  }

  const tempSheet = getFxTempSheet_(ss, opts.tempSheetName);
  const requestRows = [];

  normalizedPairs.forEach(pair => {
    requestRows.push({
      pair: pair,
      mode: "direct",
      label: `${pair.from}/${pair.to}:direct`,
      formula: `=GOOGLEFINANCE("CURRENCY:${pair.from}${pair.to}")`
    });
    requestRows.push({
      pair: pair,
      mode: "inverse",
      label: `${pair.from}/${pair.to}:inverse`,
      formula: `=GOOGLEFINANCE("CURRENCY:${pair.to}${pair.from}")`
    });
  });

  prepareFxTempBlock_(tempSheet, requestRows.length, opts);
  if (requestRows.length === 0) return result;

  const labels = requestRows.map(row => [row.label]);
  const formulas = requestRows.map(row => [row.formula]);
  tempSheet.getRange(opts.tempStartRow, opts.tempStartColumn, requestRows.length, 1).setValues(labels);
  tempSheet.getRange(opts.tempStartRow, opts.tempStartColumn + 1, requestRows.length, 1).setFormulas(formulas);
  SpreadsheetApp.flush();
  const values = readFxTempFormulaValuesWithRetry_(tempSheet, requestRows.length, opts);
  const lookup = {};

  requestRows.forEach((request, index) => {
    const pairKey = buildFxPairKey_(request.pair.from, request.pair.to);
    if (!lookup[pairKey]) lookup[pairKey] = {};
    lookup[pairKey][request.mode] = values[index][0];
  });

  normalizedPairs.forEach(pair => {
    const pairKey = buildFxPairKey_(pair.from, pair.to);
    const direct = lookup[pairKey] && lookup[pairKey].direct;
    const inverse = lookup[pairKey] && lookup[pairKey].inverse;
    const directRate = parsePositiveFxRate_(direct);
    if (directRate !== null) {
      result[pairKey] = {
        rate: directRate,
        source: "googlefinance:direct"
      };
      return;
    }

    const inverseRate = parsePositiveFxRate_(inverse);
    if (inverseRate !== null) {
      result[pairKey] = {
        rate: 1 / inverseRate,
        source: "googlefinance:inverse"
      };
    }
  });

  return result;
}

function readFxTempFormulaValuesWithRetry_(sheet, rowCount, options) {
  const opts = options || {};
  const retryCount = Math.max(Number(opts.formulaRetryCount || 3), 1);
  const retryWaitMs = Math.max(Number(opts.formulaRetryWaitMs || opts.flushWaitMs || 0), 0);
  let values = [];

  for (let attempt = 0; attempt < retryCount; attempt++) {
    SpreadsheetApp.flush();
    if (retryWaitMs > 0 && typeof Utilities !== "undefined" && Utilities.sleep) {
      Utilities.sleep(retryWaitMs);
      SpreadsheetApp.flush();
    }

    values = sheet.getRange(opts.tempStartRow, opts.tempStartColumn + 1, rowCount, 1).getValues();
    if (!hasPendingFxFormulaValues_(values)) {
      return values;
    }
  }

  return values;
}

function hasPendingFxFormulaValues_(values) {
  return (values || []).some(row => {
    const text = String(row && row[0] == null ? "" : row[0]).trim().toLowerCase();
    return !text || text.indexOf("loading") >= 0 || text.indexOf("calculating") >= 0 || text.indexOf("載入") >= 0;
  });
}

function getFxTempSheet_(ss, sheetName) {
  const name = sheetName || "Temp";
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    try {
      sheet.hideSheet();
    } catch (e) {
      logFxRateUpdate_("warn", `Unable to hide temp sheet ${name}: ${e.message || e}`);
    }
  }
  return sheet;
}

function prepareFxTempBlock_(sheet, requiredRows, options) {
  const opts = options || {};
  const startRow = opts.tempStartRow || FX_TEMP_START_ROW;
  const startColumn = opts.tempStartColumn || FX_TEMP_START_COLUMN;
  const columnCount = opts.tempOwnedColumnCount || FX_TEMP_OWNED_COLUMN_COUNT;
  const rowsNeeded = Math.max(requiredRows || 0, 1);

  ensureFxTempSheetSize_(sheet, startRow + rowsNeeded - 1, startColumn + columnCount - 1);

  const lastOwnedRowCount = findLastFxOwnedRowCount_(sheet, startRow, startColumn, columnCount);
  const clearRows = Math.max(rowsNeeded, FX_TEMP_MIN_CLEAR_ROWS, lastOwnedRowCount);
  ensureFxTempSheetSize_(sheet, startRow + clearRows - 1, startColumn + columnCount - 1);
  sheet.getRange(startRow, startColumn, clearRows, columnCount).clearContent();
}

function ensureFxTempSheetSize_(sheet, minRows, minColumns) {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < minColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minColumns - sheet.getMaxColumns());
  }
}

function findLastFxOwnedRowCount_(sheet, startRow, startColumn, columnCount) {
  const maxRows = Math.max(sheet.getLastRow() - startRow + 1, 0);
  const inspectRows = Math.max(maxRows, FX_TEMP_MIN_CLEAR_ROWS);
  ensureFxTempSheetSize_(sheet, startRow + inspectRows - 1, startColumn + columnCount - 1);

  const values = sheet.getRange(startRow, startColumn, inspectRows, columnCount).getValues();
  for (let index = values.length - 1; index >= 0; index--) {
    const hasValue = values[index].some(value => String(value == null ? "" : value).trim() !== "");
    if (hasValue) return index + 1;
  }

  return 0;
}

function filterSafeFxRateUpdatesBeforeWrite_(ss, updates, settingsOptions, attemptStartedAt, summary) {
  if (!updates || updates.length === 0) return [];

  return updates.filter(update => {
    const latest = SettingsMatrixRepo.readPairAt(ss, update, settingsOptions);
    if (!latest.exists) {
      summary.skippedChanged++;
      logFxRateUpdate_("warn", `Skipped FX write for ${update.from}/${update.to}: ${latest.message || "pair no longer exists."}`);
      return false;
    }

    if (!fxRateCellValuesMatch_(latest.value, update.expectedValue) ||
      !fxRateCellValuesMatch_(latest.timestamp, update.expectedTimestamp)) {
      summary.skippedChanged++;
      logFxRateUpdate_("warn", `Skipped FX write for ${update.from}/${update.to}: value/timestamp changed before write.`);
      if (!isFxRatePairUsableAfterFailure_({
        currentValue: latest.value,
        currentTimestamp: latest.timestamp
      }, new Date())) {
        if (update.required) {
          addFxMissingPair_(summary, `${update.from}/${update.to}`);
        } else {
          addFxOptionalUnusablePair_(summary, `${update.from}/${update.to}`);
        }
      }
      return false;
    }

    const latestUpdatedAt = parseFxRateDate_(latest.timestamp);
    if (latestUpdatedAt && latestUpdatedAt.getTime() > attemptStartedAt.getTime()) {
      summary.skippedChanged++;
      logFxRateUpdate_("warn", `Skipped FX write for ${update.from}/${update.to}: row was updated after this run started.`);
      if (!isFxRatePairUsableAfterFailure_({
        currentValue: latest.value,
        currentTimestamp: latest.timestamp
      }, new Date())) {
        if (update.required) {
          addFxMissingPair_(summary, `${update.from}/${update.to}`);
        } else {
          addFxOptionalUnusablePair_(summary, `${update.from}/${update.to}`);
        }
      }
      return false;
    }

    return true;
  });
}

function addFxMissingPair_(summary, pairLabel) {
  if (!summary || !pairLabel) return;
  if (summary.missingPairs.indexOf(pairLabel) < 0) {
    summary.missingPairs.push(pairLabel);
  }
}

function addFxOptionalUnusablePair_(summary, pairLabel) {
  if (!summary || !pairLabel) return;
  if (summary.optionalUnusablePairs.indexOf(pairLabel) < 0) {
    summary.optionalUnusablePairs.push(pairLabel);
  }
}

function normalizeFxRequiredPairs_(pairs, options) {
  const opts = options || {};
  const required = [];
  const lookup = {};

  (pairs || []).forEach(pair => {
    const from = String(pair && pair.from || "").trim().toUpperCase();
    const to = String(pair && pair.to || "").trim().toUpperCase();
    if (!from || !to) return;
    const key = buildFxPairKey_(from, to);
    if (lookup[key]) return;
    lookup[key] = true;
    required.push({ from: from, to: to });
  });

  if (opts.includeMandatoryPairs !== false) {
    const mandatory = { from: "USD", to: "TWD" };
    const mandatoryKey = buildFxPairKey_(mandatory.from, mandatory.to);
    if (!lookup[mandatoryKey]) {
      required.push(mandatory);
    }
  }

  return required;
}

function buildFxRequiredPairLookup_(pairs) {
  const lookup = {};
  (pairs || []).forEach(pair => {
    const key = buildFxPairKey_(pair.from, pair.to);
    if (key !== "__") lookup[key] = true;
  });
  return lookup;
}

function isFxRatePairUsableAfterFailure_(pair, now) {
  if (!pair || !isPositiveFxRate_(pair.currentValue)) return false;

  const updatedAt = parseFxRateDate_(pair.currentTimestamp);
  if (!updatedAt) return false;

  const staleHours = (typeof Config !== "undefined" && Config.THRESHOLDS && Config.THRESHOLDS.FX_RATE_STALE_HOURS) || 24;
  return (now || new Date()).getTime() - updatedAt.getTime() <= staleHours * 60 * 60 * 1000;
}

function buildFxPairKey_(from, to) {
  return `${String(from || "").trim().toUpperCase()}__${String(to || "").trim().toUpperCase()}`;
}

function parsePositiveFxRate_(value) {
  const num = parseFloat(value);
  return isFinite(num) && num > 0 ? num : null;
}

function isPositiveFxRate_(value) {
  return parsePositiveFxRate_(value) !== null;
}

function parseFxRateDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function fxRateCellValuesMatch_(currentValue, expectedValue) {
  if (!currentValue && !expectedValue) return true;

  if (isFxRateNumericLiteral_(currentValue) || isFxRateNumericLiteral_(expectedValue)) {
    const currentNumber = parseFloat(currentValue);
    const expectedNumber = parseFloat(expectedValue);
    return isFinite(currentNumber) && isFinite(expectedNumber) && Math.abs(currentNumber - expectedNumber) < FX_RATE_EPSILON;
  }

  const currentDate = parseFxRateDate_(currentValue);
  const expectedDate = parseFxRateDate_(expectedValue);
  if (currentDate || expectedDate) {
    return !!currentDate && !!expectedDate && currentDate.getTime() === expectedDate.getTime();
  }

  return String(currentValue == null ? "" : currentValue).trim() === String(expectedValue == null ? "" : expectedValue).trim();
}

function isFxRateNumericLiteral_(value) {
  if (typeof value === "number") return isFinite(value);
  return /^-?\d+(\.\d+)?$/.test(String(value == null ? "" : value).trim());
}

function logFxRateUpdate_(level, message) {
  if (typeof LogService !== "undefined" && LogService[level]) {
    LogService[level](message, FX_UPDATE_MODULE_NAME);
    return;
  }

  Logger.log(`[${FX_UPDATE_MODULE_NAME}] ${message}`);
}
