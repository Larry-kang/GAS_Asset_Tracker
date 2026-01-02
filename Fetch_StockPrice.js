function getStockPrice(stockSymbol) {
  var price = null;

  // 1. 嘗試使用 GOOGLEFINANCE
  price = getPriceFromGoogleFinance(stockSymbol);
  if (price !== null) {
    return price;
  }

  // 2. 嘗試從銳享網抓取數據（僅適用於台灣股票）
  if (stockSymbol.match(/^\d+$/)) { // 檢查是否為台灣股票代號
    price = getPriceFromCnyes(stockSymbol);
    if (price !== null) {
      return price;
    }
  }

  // 3. 嘗試從 Yahoo 股市抓取數據
  price = getPriceFromYahoo(stockSymbol);
  if (price !== null) {
    return price;
  }

  // 如果所有嘗試都失敗，返回錯誤信息
  return "無法取得股價";
}

function getPriceFromGoogleFinance(stockSymbol) {
  try {
    var googleFinanceUrl = "https://www.google.com/finance/quote/" + stockSymbol;
    var response = UrlFetchApp.fetch(googleFinanceUrl);
    var content = response.getContentText();
    var regex = /<div[^>]*class="YMlKec fxKbKc">(\d+(?:\.\d+)?)<\/div>/;
    var match = content.match(regex);
    if (match && match[1]) {
      return parseFloat(match[1]);
    }
  } catch (e) {
    Logger.log("GOOGLEFINANCE 不可用，錯誤: " + e);
  }
  return null;
}

function getPriceFromCnyes(stockSymbol) {
  try {
    var cnyesUrl = "https://www.cnyes.com/twstock/" + stockSymbol;
    var response = UrlFetchApp.fetch(cnyesUrl);
    var content = response.getContentText();
    var regex = /<h3[^>]*>(\d+(?:\.\d+)?)<\/h3>/;
    var match = content.match(regex);
    if (match && match[1]) {
      return parseFloat(match[1]);
    }
  } catch (e) {
    Logger.log("銳享網不可用，錯誤: " + e);
  }
  return null;
}

function getPriceFromYahoo(stockSymbol) {
  try {
    var yahooUrl = "https://tw.stock.yahoo.com/quote/" + stockSymbol;
    var response = UrlFetchApp.fetch(yahooUrl);
    var content = response.getContentText();
    var regex = /<span[^>]*>(\d+(?:\.\d+)?)<\/span>/;
    var match = content.match(regex);
    if (match && match[1]) {
      return parseFloat(match[1]);
    }
  } catch (e) {
    Logger.log("Yahoo 股市不可用，錯誤: " + e);
  }
  return null;
}

/**
 * Test function for manual execution - tests complete stock price fetching flow.
 * @example Run in Apps Script editor to test all fallback mechanisms
 * @private
 */
function testGetStockPrice() {
  var stockSymbol = "2330"; // 輸入股票代號
  var price = getStockPrice(stockSymbol);
  Logger.log("股票代號 " + stockSymbol + " 的股價為: " + price);
}

/**
 * Test function - tests Google Finance price fetching mechanism.
 * @private
 */
function testGetPriceFromGoogleFinance() {
  var stockSymbol = "2330";
  var price = getPriceFromGoogleFinance(stockSymbol);
  Logger.log("GOOGLEFINANCE 股票代號 " + stockSymbol + " 的股價為: " + price);
}

/**
 * Test function - tests Cnyes (銳享網) price fetching mechanism.
 * @private
 */
function testGetPriceFromCnyes() {
  var stockSymbol = "2330";
  var price = getPriceFromCnyes(stockSymbol);
  Logger.log("銳享網 股票代號 " + stockSymbol + " 的股價為: " + price);
}

/**
 * Test function - tests Yahoo Finance price fetching mechanism.
 * @private
 */
function testGetPriceFromYahoo() {
  var stockSymbol = "2330";
  var price = getPriceFromYahoo(stockSymbol);
  Logger.log("Yahoo 股市 股票代號 " + stockSymbol + " 的股價為: " + price);
}

// End of Stock Price module
