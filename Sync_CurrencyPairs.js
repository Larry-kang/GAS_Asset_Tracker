/**
 * 智慧同步函式 (v2.0 - 支援儲存格引用)
 * 掃描所有工作表，找出所有 MY_GOOGLEFINANCE 的用法，
 * 無論參數是文字還是儲存格引用，都能解析出真實的貨幣對，
 * 並自動在 '參數設定' 工作表中新增缺失的行和列。
 */
function syncCurrencyPairs(silentOrOptions) {
  const MODULE_NAME = 'Sync_CurrencyPairs';
  const opts = normalizeCurrencyPairSyncOptions_(silentOrOptions);
  if (!opts.silent) logCurrencyPairSync_('info', 'Starting Currency Pair Sync...', MODULE_NAME);

  try {
    const ss = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    SettingsMatrixRepo.getSheet_(ss, { sheetName: opts.settingsSheetName });
    const requiredPairsResult = collectRequiredCurrencyPairs_(ss, opts);

    const appendResult = SettingsMatrixRepo.appendMissingPairs(
      ss,
      requiredPairsResult.pairs,
      { sheetName: opts.settingsSheetName }
    );
    const newFromAdded = appendResult.newFromAdded;
    const newToAdded = appendResult.newToAdded;

    appendResult.addedFromCurrencies.forEach(currency => {
      logCurrencyPairSync_('info', `Added From-Currency: ${currency}`, MODULE_NAME);
    });

    appendResult.addedToCurrencies.forEach(currency => {
      logCurrencyPairSync_('info', `Added To-Currency: ${currency}`, MODULE_NAME);
    });

    if (!opts.silent) {
      if (newFromAdded || newToAdded) {
        logCurrencyPairSync_('info', 'Sync complete. New pairs added.', MODULE_NAME);
        SpreadsheetApp.getUi().alert('成功新增缺失的貨幣對！稍後匯率將會自動更新。');
      } else {
        logCurrencyPairSync_('info', 'Sync complete. No new pairs needed.', MODULE_NAME);
        SpreadsheetApp.getUi().alert('所有貨幣對均已存在，無需新增。');
      }
    } else {
      if (newFromAdded || newToAdded) {
        logCurrencyPairSync_('info', 'Sync complete (Silent). New pairs added.', MODULE_NAME);
      } else {
        // Silent mode often used in automated runs, maybe skip info logging to reduce noise, or keep as info
        // Keeping as info for audit trail
      }
    }

    return {
      status: 'COMPLETE',
      ok: true,
      fatal: false,
      requiredPairCount: requiredPairsResult.pairs.length,
      addedFromCurrencies: appendResult.addedFromCurrencies,
      addedToCurrencies: appendResult.addedToCurrencies,
      message: newFromAdded || newToAdded ? 'New pairs added.' : 'No new pairs needed.'
    };
  } catch (e) {
    const message = e.message || String(e);
    logCurrencyPairSync_('error', `Sync failed: ${message}`, MODULE_NAME);
    if (!opts.silent) SpreadsheetApp.getUi().alert(`同步失敗：${message}`);
    if (opts.throwOnError) throw e;

    return {
      status: 'FAILED',
      ok: false,
      fatal: true,
      requiredPairCount: 0,
      addedFromCurrencies: [],
      addedToCurrencies: [],
      message: message
    };
  }
}

function normalizeCurrencyPairSyncOptions_(silentOrOptions) {
  if (typeof silentOrOptions === 'boolean') {
    return {
      silent: silentOrOptions,
      throwOnError: false,
      spreadsheet: null,
      settingsSheetName: null,
      sheetNames: null
    };
  }

  const opts = silentOrOptions || {};
  return {
    silent: opts.silent === undefined ? false : !!opts.silent,
    throwOnError: !!opts.throwOnError,
    spreadsheet: opts.spreadsheet || null,
    settingsSheetName: opts.settingsSheetName || null,
    sheetNames: opts.sheetNames || null
  };
}

function collectRequiredCurrencyPairs_(ss, options) {
  const opts = options || {};
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheetName = opts.settingsSheetName || SettingsMatrixRepo.DEFAULT_SHEET_NAME;
  const requiredPairs = new Set();
  const formulaRegex = /MY_GOOGLEFINANCE\(([^,)]+),([^,)]+)\)/gi;
  const sheets = resolveCurrencyPairScanSheets_(spreadsheet, opts);

  sheets.forEach(sheet => {
    if (!sheet) return;
    if (sheet.getName() === settingsSheetName) return;

    const formulas = sheet.getDataRange().getFormulas();

    formulas.forEach(row => {
      row.forEach(formula => {
        if (!formula || !String(formula).toUpperCase().includes('MY_GOOGLEFINANCE')) return;

        formulaRegex.lastIndex = 0;
        let match = formulaRegex.exec(formula);
        while (match) {
          const argText = match[1].trim();
          const arg2Text = match[2].trim();

          const fromCurrency = normalizeCurrencyPairValue_(resolveArgument(argText, sheet));
          const toCurrency = normalizeCurrencyPairValue_(resolveArgument(arg2Text, sheet));

          if (fromCurrency && toCurrency) {
            requiredPairs.add(JSON.stringify({ from: fromCurrency, to: toCurrency }));
          }

          match = formulaRegex.exec(formula);
        }
      });
    });
  });

  return {
    pairs: Array.from(requiredPairs).map(pairStr => JSON.parse(pairStr))
  };
}

function resolveCurrencyPairScanSheets_(spreadsheet, options) {
  const opts = options || {};
  if (!opts.sheetNames || opts.sheetNames.length === 0) return spreadsheet.getSheets();

  return opts.sheetNames.map(sheetName => spreadsheet.getSheetByName(sheetName)).filter(Boolean);
}

function normalizeCurrencyPairValue_(value) {
  return String(value == null ? "" : value).trim().toUpperCase();
}

/**
 * 輔助函式：解析公式參數
 * @param {string} argText - 公式中的參數文字, e.g., "USD" or A1 or '工作表2'!B2
 * @param {Sheet} currentSheet - 該公式所在的工作表對象
 * @return {string | null} 解析出的貨幣代碼, e.g., "USD", or null if failed
 */
function resolveArgument(argText, currentSheet) {
  // 情況1：參數是文字, e.g., "USD" or 'USD'
  if ((argText.startsWith('"') && argText.endsWith('"')) || (argText.startsWith("'") && argText.endsWith("'"))) {
    return argText.substring(1, argText.length - 1);
  }

  // 情況2：參數是儲存格引用
  try {
    const spreadsheet = currentSheet.getParent ? currentSheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();
    // 檢查是否包含工作表名稱, e.g., '工作表2'!B2
    if (argText.includes('!')) {
      return spreadsheet.getRange(argText).getValue();
    } else {
      // 引用當前工作表的儲存格, e.g., A1
      return currentSheet.getRange(argText).getValue();
    }
  } catch (e) {
    Logger.log(`無法解析參數 "${argText}" 於工作表 "${currentSheet.getName()}": ${e.toString()}`);
    return null;
  }
}

function logCurrencyPairSync_(level, message, moduleName) {
  if (typeof LogService !== 'undefined' && LogService[level]) {
    LogService[level](message, moduleName);
    return;
  }

  Logger.log(`[${moduleName}] [${String(level || 'info').toUpperCase()}] ${message}`);
}
