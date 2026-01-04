// =================================================================
// == 自動化價格更新系統 (v2.0 - 2025-10-10)                   ==
// == 核心功能：                                              ==
// == 1. updateAllPrices: 主引擎，按時觸發，遍歷所有資產        ==
// == 2. fetchCryptoPrice: 數位幣價格調度中心，具備備援機制     ==
// == 3. getCryptoPrice: 從 CoinMarketCap 獲取價格的穩健函式   ==
// == 4. fetchStockPrice: 股票價格調度中心，具備備援機制        ==
// =================================================================


// =======================================================
// 主更新函式 - 請將此函式設定為每15或30分鐘觸發一次
// =======================================================
function updateAllPrices() {
  const MODULE_NAME = 'Fetch_CryptoPrice';
  const context = 'updateAllPrices';
  LogService.info('Starting Market Price Update...', MODULE_NAME);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('價格暫存');
    if (!sheet) {
      const msg = "錯誤：找不到名為 '價格暫存' 的工作表。";
      LogService.error(msg, MODULE_NAME);
      return;
    }

    const dataRange = sheet.getRange('A2:D' + sheet.getLastRow());
    const data = dataRange.getValues();
    const currentTime = new Date();
    const CACHE_DURATION = 1 * 60 * 1000; // 快取時間：1分鐘 (原15分似太久, 配合高頻監控調整)

    let updatedCount = 0;
    let failedCount = 0;

    for (let i = 0; i < data.length; i++) {
      const ticker = data[i][0];
      const type = data[i][1];
      const lastUpdated = data[i][3];

      // 檢查標的是否存在，以及是否需要更新
      if (ticker && (!lastUpdated || (currentTime.getTime() - new Date(lastUpdated).getTime()) > CACHE_DURATION)) {
        let price = null;

        // 根據資產類型，呼叫不同的價格獲取“調度中心”
        if (type === "數位幣" || type === "數位穩定幣") {
          price = fetchCryptoPrice(ticker);
        } else { // 假設其他皆為股票
          price = fetchStockPrice(ticker);
        }

        // 如果成功獲取價格，則更新到工作表中
        if (typeof price === 'number' && !isNaN(price)) {
          sheet.getRange(i + 2, 3).setValue(price); // 更新價格
          sheet.getRange(i + 2, 4).setValue(currentTime); // 更新時間戳
          updatedCount++;
          // Logger.log(`[OK] 成功更新 ${ticker} 的價格: ${price}`); // 減少細節日誌以節省空間
        } else {
          failedCount++;
          LogService.warn(`[FAILED] 更新 ${ticker} 的價格失敗。`, MODULE_NAME);
        }
      }
    }

    LogService.info(`Price Update Completed. Updated: ${updatedCount}, Failed: ${failedCount}`, MODULE_NAME);

  } catch (e) {
    LogService.error(`Price Update Failed: ${e.message}`, MODULE_NAME);
  }
}

// =======================================================
// == 輔助函式 (Helper Functions)
// =======================================================

/**
 * 數位幣價格調度中心
 * 依次嘗試多個數據來源，直到成功獲取價格為止。
 */
function fetchCryptoPrice(ticker, currency = 'USD', bypassCache = false) {
  const cacheKey = `PRICE_CRYPTO_${ticker.toUpperCase()}_${currency.toUpperCase()}`;

  // 1. Check Cache (unless bypass is requested)
  if (!bypassCache) {
    const cached = ScriptCache.get(cacheKey);
    if (cached) return cached;
  }

  let price = null;

  // --- 策略1：嘗試從 cryptoprices.cc (快速、免費) ---
  try {
    const url = `https://cryptoprices.cc/${ticker.toUpperCase()}/`;
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });

    if (response.getResponseCode() === 200) {
      price = parseFloat(response.getContentText());
      if (typeof price === 'number' && !isNaN(price)) {
        Logger.log(`  - [Source 1: cryptoprices.cc] OK for ${ticker}: ${price}`);
        return price;
      }
    }
  } catch (e) {
    Logger.log(`  - [Source 1: cryptoprices.cc] FAILED for ${ticker}: ${e.toString()}`);
  }

  // --- 策略2：如果策略1失敗，則呼叫 CoinMarketCap (穩定、可靠) ---
  Logger.log(`  - [Source 1 FAILED] Falling back to Source 2 (CoinMarketCap) for ${ticker}...`);
  price = getCryptoPrice(ticker, currency); // 呼叫重新命名的穩健函式

  if (price !== null) {
    Logger.log(`  - [Source 2: CoinMarketCap] OK for ${ticker}: ${price}`);
    // Store in cache for 15 minutes
    ScriptCache.put(cacheKey, price, 900);
    return price;
  }

  return null; // 所有方法都失敗
}

/**
 * [穩健版] 從 CoinMarketCap 獲取加密貨幣價格 (已重新命名)
 */
function getCryptoPrice(symbol, convert) {
  var convertCurrency = convert || "USD";
  var apiKey = Settings.get('CMC_API_KEY');

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
        if (typeof price === 'number' && !isNaN(price)) {
          return price;
        }
      } else {
        Logger.log(`  - CoinMarketCap ERROR: Symbol ${symbol} not found in response.`);
      }
    } else {
      Logger.log(`  - CoinMarketCap API ERROR: ${json.status.error_message || 'Request failed'}`);
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
function fetchStockPrice(ticker, bypassCache = false) {
  const cacheKey = `PRICE_STOCK_${ticker.toUpperCase()}`;

  if (!bypassCache) {
    const cached = ScriptCache.get(cacheKey);
    if (cached) return cached;
  }

  let price = null;

  // --- 策略1：嘗試 GOOGLEFINANCE (最穩定) ---
  try {
    const tempSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temp');
    if (!tempSheet) throw new Error("找不到名為 'Temp' 的臨時工作表。");
    const tempRange = tempSheet.getRange('A1');
    tempRange.setFormula(`=GOOGLEFINANCE("${ticker}","price")`);
    SpreadsheetApp.flush();
    price = tempRange.getValue();
    if (typeof price === 'number' && !isNaN(price) && price > 0) {
      Logger.log(`  - [Source 1: GOOGLEFINANCE] OK for ${ticker}: ${price}`);
      return price;
    }
  } catch (e) {
    Logger.log(`  - [Source 1: GOOGLEFINANCE] FAILED for ${ticker}: ${e.toString()}`);
  }

  // --- 策略2：如果策略1失敗，嘗試爬取 cnyes.com (適用台股) ---
  Logger.log(`  - [Source 1 FAILED] Falling back to Source 2 (cnyes.com) for ${ticker}...`);
  try {
    const url = `https://www.cnyes.com/twstock/${ticker}`;
    const html = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true }).getContentText();
    const match = html.match(/<h3 class="[^"]*">([0-9,]+\.?[0-9]*)<\/h3>/);
    if (match && match[1]) {
      price = parseFloat(match[1].replace(/,/g, ''));
      if (typeof price === 'number' && !isNaN(price)) {
        Logger.log(`  - [Source 2: cnyes.com] OK for ${ticker}: ${price}`);
        return price;
      }
    }
  } catch (e) {
    Logger.log(`  - [Source 2: cnyes.com] FAILED for ${ticker}: ${e.toString()}`);
  }

  // --- 策略3：未來可在此增加 Yahoo Finance 等其他爬蟲來源 ---

  if (price !== null) {
    ScriptCache.put(cacheKey, price, 3600); // Stock prices can cache longer (1hr)
    return price;
  }

  return null; // 所有方法都失敗
}
