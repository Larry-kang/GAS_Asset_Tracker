/**
 * 正式價格查詢相容層。
 * 供活表既有公式 `GET_PRICE(ticker)` 使用。
 *
 * 優先順序：
 * 1. 從 `價格暫存` / `Price Cache` 讀取
 * 2. 未命中時依標的類型即時抓價
 *
 * @param {string} ticker
 * @return {number|string|null}
 * @customfunction
 */
function GET_PRICE(ticker) {
  if (ticker === null || ticker === undefined || ticker === '') return null;

  const normalizedTicker = String(ticker).trim().toUpperCase();
  if (!normalizedTicker) return null;

  const cacheKey = `UDF_GET_PRICE_${normalizedTicker}`;

  try {
    const cached = ScriptCache.get(cacheKey);
    if (cached !== null && cached !== undefined && cached !== '') {
      const cachedNum = parseFloat(cached);
      if (!isNaN(cachedNum) && cachedNum > 0) return cachedNum;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const priceEntry = lookupPriceCacheEntry_(ss, normalizedTicker);

    if (priceEntry && typeof priceEntry.price === 'number' && !isNaN(priceEntry.price) && priceEntry.price > 0) {
      ScriptCache.put(cacheKey, String(priceEntry.price), 60);
      return priceEntry.price;
    }

    let price = null;
    if (priceEntry && priceEntry.assetType) {
      price = fetchPriceByAssetType_(normalizedTicker, priceEntry.assetType);
    } else {
      price = inferAndFetchPrice_(normalizedTicker);
    }

    if (typeof price === 'number' && !isNaN(price) && price > 0) {
      ScriptCache.put(cacheKey, String(price), 60);
      return price;
    }

    return null;
  } catch (e) {
    return 'Err: ' + String(e.message || e).substring(0, 50);
  }
}

function lookupPriceCacheEntry_(ss, ticker) {
  const rows = readPriceCacheRowsSafe_(ss);
  for (let i = 0; i < rows.length; i++) {
    const rowTicker = String(rows[i].ticker || '').trim().toUpperCase();
    if (rowTicker === ticker) {
      return {
        ticker: rowTicker,
        assetType: rows[i].assetType || '',
        price: parseLookupPrice_(rows[i].price)
      };
    }
  }
  return null;
}

function readPriceCacheRowsSafe_(ss) {
  try {
    if (typeof PriceCacheRepo !== 'undefined' && PriceCacheRepo.readRows) {
      return PriceCacheRepo.readRows(ss);
    }
  } catch (e) {
    // Fall through to manual legacy-compatible read below.
  }

  const sheet = ss.getSheetByName('價格暫存') || ss.getSheetByName('Price Cache');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 4).getValues().map(function (row, index) {
    return {
      rowNumber: index + 2,
      ticker: row[0],
      assetType: row[1],
      price: row[2],
      updatedAt: row[3]
    };
  });
}

function fetchPriceByAssetType_(ticker, assetType) {
  const type = String(assetType || '').trim();
  if (type === '數位幣' || type === '數位穩定幣') {
    if (typeof fetchCryptoPrice === 'function') {
      const cryptoPrice = fetchCryptoPrice(ticker);
      if (typeof cryptoPrice === 'number' && !isNaN(cryptoPrice) && cryptoPrice > 0) return cryptoPrice;
    }
    if (typeof fetchStockPrice === 'function') {
      return fetchStockPrice(ticker);
    }
    return null;
  }

  if (typeof fetchStockPrice === 'function') {
    const stockPrice = fetchStockPrice(ticker);
    if (typeof stockPrice === 'number' && !isNaN(stockPrice) && stockPrice > 0) return stockPrice;
  }

  if (typeof fetchCryptoPrice === 'function') {
    return fetchCryptoPrice(ticker);
  }

  return null;
}

function inferAndFetchPrice_(ticker) {
  if (/^\d+$/.test(ticker)) {
    return typeof fetchStockPrice === 'function' ? fetchStockPrice(ticker) : null;
  }

  if (typeof fetchCryptoPrice === 'function') {
    const cryptoPrice = fetchCryptoPrice(ticker);
    if (typeof cryptoPrice === 'number' && !isNaN(cryptoPrice) && cryptoPrice > 0) return cryptoPrice;
  }

  return typeof fetchStockPrice === 'function' ? fetchStockPrice(ticker) : null;
}

function parseLookupPrice_(value) {
  const num = parseFloat(value);
  return isNaN(num) ? null : num;
}
