/**
 * 智慧同步資產函式
 * 掃描所有工作表，找出所有 GET_PRICE 的用法，
 * 並自動在 '價格暫存' 工作表中新增缺失的資產標的。
 */
function syncAssets() {
  const MODULE_NAME = 'Sync_Assets';
  LogService.info('Starting Asset & Currency Pair Sync...', MODULE_NAME);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName('價格暫存');
    if (!targetSheet) {
      const msg = "錯誤：找不到名為 '價格暫存' 的工作表。";
      LogService.error(msg, MODULE_NAME);
      SpreadsheetApp.getUi().alert(msg);
      return;
    }
    const allSheets = ss.getSheets();

    // 獲取 '價格暫存' 表中已有的資產列表
    const existingAssets = targetSheet.getRange('A2:A').getValues().flat().filter(String);

    const requiredAssets = new Set();
    const formulaRegex = /GET_PRICE\(([^,)]+)/gi;

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName === '價格暫存' || sheetName === '參數設定' || sheetName === 'Temp') return;

      const formulas = sheet.getDataRange().getFormulas();

      formulas.forEach((row, rIndex) => {
        row.forEach((formula, cIndex) => {
          if (formula && formula.toUpperCase().includes('GET_PRICE')) {
            formulaRegex.lastIndex = 0;
            let match = formulaRegex.exec(formula);

            if (match) {
              const argText = match[1].trim();
              const assetTicker = resolveArgument(argText, sheet);
              if (assetTicker) requiredAssets.add(assetTicker);
            }
          }
        });
      });
    });

    let newAssetsAddedCount = 0;

    requiredAssets.forEach(ticker => {
      if (!existingAssets.includes(ticker)) {
        targetSheet.appendRow([ticker, '請手動填寫類型']);
        existingAssets.push(ticker);
        newAssetsAddedCount++;
        LogService.info(`Added new asset ticker: ${ticker}`, MODULE_NAME);
      }
    });

    if (newAssetsAddedCount > 0) {
      LogService.info(`Sync complete. Added ${newAssetsAddedCount} new assets.`, MODULE_NAME);
      SpreadsheetApp.getUi().alert(`成功新增 ${newAssetsAddedCount} 個資產！請記得前往 "價格暫存" 工作表為它們填寫正確的 "類型"。`);
    } else {
      LogService.info('Sync complete. No new assets found.', MODULE_NAME);
      SpreadsheetApp.getUi().alert('所有資產均已存在，無需新增。');
    }
  } catch (e) {
    LogService.error(`Sync failed: ${e.message}`, MODULE_NAME);
    SpreadsheetApp.getUi().alert(`同步失敗：${e.message}`);
  }
}
