function MY_GOOGLEFINANCE(fromCurrency, toCurrency) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lookup = null;

  try {
    lookup = SettingsMatrixRepo.lookupRate(ss, fromCurrency, toCurrency, { allowLegacyName: true });
  } catch (e) {
    if (e.message && e.message.includes('missing required sheet')) {
      return "#SHEET_MISSING! - 參數設定";
    }
    return "#CONFIG_ERROR! - 參數設定";
  }

  if (!lookup || !lookup.sheet) {
    return "#SHEET_MISSING! - 參數設定";
  }

  if (!lookup.hasPair) {
    return `#NEW_PAIR! - ${fromCurrency}/${toCurrency}`;
  }

  // 檢查匯率是否為數字，並返回
  if (typeof lookup.value === 'number' && !isNaN(lookup.value)) {
    return lookup.value;
  } else {
    // 如果表格沒有值，可能是因為還沒更新
    return 'Updating...';
  }
}
