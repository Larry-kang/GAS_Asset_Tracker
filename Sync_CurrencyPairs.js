/**
 * 智慧同步函式 (v2.0 - 支援儲存格引用)
 * 掃描所有工作表，找出所有 MY_GOOGLEFINANCE 的用法，
 * 無論參數是文字還是儲存格引用，都能解析出真實的貨幣對，
 * 並自動在 '參數設定' 工作表中新增缺失的行和列。
 */
function syncCurrencyPairs(silent = false) {
  const MODULE_NAME = 'Sync_CurrencyPairs';
  if (!silent) LogService.info('Starting Currency Pair Sync...', MODULE_NAME);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SettingsMatrixRepo.getSheet_(ss);
    const allSheets = ss.getSheets();
    const requiredPairs = new Set();
    const formulaRegex = /MY_GOOGLEFINANCE\(([^,)]+),([^,)]+)\)/gi;

    // 遍歷所有工作表
    allSheets.forEach(sheet => {
      // 跳過參數設定表本身
      if (sheet.getName() === '參數設定') return;

      const formulas = sheet.getDataRange().getFormulas();

      formulas.forEach((row, rIndex) => {
        row.forEach((formula, cIndex) => {
          if (formula && formula.toUpperCase().includes('MY_GOOGLEFINANCE')) {
            formulaRegex.lastIndex = 0;
            let match = formulaRegex.exec(formula);

            if (match) {
              const argText = match[1].trim();
              const arg2Text = match[2].trim();

              const fromCurrency = resolveArgument(argText, sheet);
              const toCurrency = resolveArgument(arg2Text, sheet);

              if (fromCurrency && toCurrency) {
                requiredPairs.add(JSON.stringify({ from: fromCurrency, to: toCurrency }));
              }
            }
          }
        });
      });
    });

    const appendResult = SettingsMatrixRepo.appendMissingPairs(
      ss,
      Array.from(requiredPairs).map(pairStr => JSON.parse(pairStr))
    );
    const newFromAdded = appendResult.newFromAdded;
    const newToAdded = appendResult.newToAdded;

    appendResult.addedFromCurrencies.forEach(currency => {
      LogService.info(`Added From-Currency: ${currency}`, MODULE_NAME);
    });

    appendResult.addedToCurrencies.forEach(currency => {
      LogService.info(`Added To-Currency: ${currency}`, MODULE_NAME);
    });

    if (!silent) {
      if (newFromAdded || newToAdded) {
        LogService.info('Sync complete. New pairs added.', MODULE_NAME);
        SpreadsheetApp.getUi().alert('成功新增缺失的貨幣對！稍後匯率將會自動更新。');
      } else {
        LogService.info('Sync complete. No new pairs needed.', MODULE_NAME);
        SpreadsheetApp.getUi().alert('所有貨幣對均已存在，無需新增。');
      }
    } else {
      if (newFromAdded || newToAdded) {
        LogService.info('Sync complete (Silent). New pairs added.', MODULE_NAME);
      } else {
        // Silent mode often used in automated runs, maybe skip info logging to reduce noise, or keep as info
        // Keeping as info for audit trail
      }
    }
  } catch (e) {
    LogService.error(`Sync failed: ${e.message}`, MODULE_NAME);
    if (!silent) SpreadsheetApp.getUi().alert(`同步失敗：${e.message}`);
  }
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
