function MY_GOOGLEFINANCE(fromCurrency, toCurrency) {
  // 支援英文或中文工作表名稱以維持向後兼容性
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Market Settings') || ss.getSheetByName('參數設定');

  if (!sheet) {
    return "#SHEET_MISSING! - 參數設定";
  }

  const values = sheet.getDataRange().getValues();

  // 尋找 fromCurrency 的行索引
  let fromRowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === fromCurrency) {
      fromRowIndex = i;
      break;
    }
  }

  // 尋找 toCurrency 的列索引
  let toColumnIndex = -1;
  for (let j = 1; j < values[0].length; j++) {
    if (values[0][j] === toCurrency) {
      toColumnIndex = j;
      break;
    }
  }

  // ⭐ 如果找不到，返回一個清晰的提示，而不是 #N/A
  if (fromRowIndex === -1 || toColumnIndex === -1) {
    return `#NEW_PAIR! - ${fromCurrency}/${toCurrency}`;
  }

  const storedRate = values[fromRowIndex][toColumnIndex];

  // 檢查匯率是否為數字，並返回
  if (typeof storedRate === 'number' && !isNaN(storedRate)) {
    return storedRate;
  } else {
    // 如果表格沒有值，可能是因為還沒更新
    return 'Updating...';
  }
}
