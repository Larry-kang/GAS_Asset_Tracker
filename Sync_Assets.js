/**
 * 智慧同步資產函式
 * 掃描所有工作表，找出所有 GET_PRICE 的用法，
 * 並自動在 '價格暫存' 工作表中新增缺失的資產標的。
 */
function syncAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName('價格暫存');
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到名為 '價格暫存' 的工作表。");
    return;
  }
  const allSheets = ss.getSheets();

  // 獲取 '價格暫存' 表中已有的資產列表
  const existingAssets = targetSheet.getRange('A2:A').getValues().flat().filter(String);
  
  const requiredAssets = new Set();
  // 正則表達式只捕獲 GET_PRICE 的第一個參數
  const formulaRegex = /GET_PRICE\(([^,)]+)/gi;

  // 遍歷所有工作表
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // 跳過系統工作表
    if (sheetName === '價格暫存' || sheetName === '參數設定' || sheetName === 'Temp') return;

    const formulas = sheet.getDataRange().getFormulas();
    
    formulas.forEach((row, rIndex) => {
      row.forEach((formula, cIndex) => {
        if (formula && formula.toUpperCase().includes('GET_PRICE')) {
          formulaRegex.lastIndex = 0; 
          let match = formulaRegex.exec(formula);
          
          if (match) {
            const argText = match[1].trim();
            const assetTicker = resolveArgument(argText, sheet); // 重用我們已有的智慧解析函式
            
            if (assetTicker) {
              requiredAssets.add(assetTicker);
            }
          }
        }
      });
    });
  });

  let newAssetsAddedCount = 0;

  // 處理所有需要的新資產
  requiredAssets.forEach(ticker => {
    if (!existingAssets.includes(ticker)) {
      // 新增資產，並提示使用者手動填寫類型
      targetSheet.appendRow([ticker, '請手動填寫類型']); 
      existingAssets.push(ticker); // 更新記憶體中的列表，避免重複新增
      newAssetsAddedCount++;
      Logger.log(`新增資產標的: ${ticker}`);
    }
  });

  if (newAssetsAddedCount > 0) {
    SpreadsheetApp.getUi().alert(`成功新增 ${newAssetsAddedCount} 個資產！請記得前往 "價格暫存" 工作表為它們填寫正確的 "類型"。`);
  } else {
    SpreadsheetApp.getUi().alert('所有資產均已存在，無需新增。');
  }
}
