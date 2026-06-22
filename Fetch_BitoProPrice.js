// =================================================================
// == BitoPro Public Market Price Source                          ==
// == 用於 BITO 等公開行情，不依賴帳戶 API key。                  ==
// =================================================================

const BITOPRO_PUBLIC_API_BASE_URL = "https://api.bitopro.com/v3";
const BITOPRO_TICKER_FAILURE_BACKOFF_SECONDS = 180;

function fetchBitoProTickerLastPrice_(pair, bypassCache) {
  const normalizedPair = String(pair || "").trim().toLowerCase();
  if (!normalizedPair) return null;

  const cacheKey = `PRICE_BITOPRO_TICKER_${normalizedPair.toUpperCase()}`;
  const failureCacheKey = `PRICE_BITOPRO_TICKER_FAIL_${normalizedPair.toUpperCase()}`;
  if (!bypassCache) {
    const cached = ScriptCache.get(cacheKey);
    const cachedPrice = parseBitoProPositiveNumber_(cached);
    if (cachedPrice !== null) return cachedPrice;

    if (ScriptCache.get(failureCacheKey)) return null;
  }

  try {
    const url = `${BITOPRO_PUBLIC_API_BASE_URL}/tickers/${encodeURIComponent(normalizedPair)}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code !== 200) {
      ScriptCache.put(failureCacheKey, "1", BITOPRO_TICKER_FAILURE_BACKOFF_SECONDS);
      logBitoProPrice_("warn", `BitoPro ticker ${normalizedPair} failed: HTTP ${code}`);
      return null;
    }

    const json = JSON.parse(body);
    const payload = extractBitoProTickerPayload_(json, normalizedPair);
    const price = parseBitoProPositiveNumber_(payload && (payload.lastPrice || payload.last || payload.close));

    if (price === null) {
      ScriptCache.put(failureCacheKey, "1", BITOPRO_TICKER_FAILURE_BACKOFF_SECONDS);
      logBitoProPrice_("warn", `BitoPro ticker ${normalizedPair} missing valid lastPrice.`);
      return null;
    }

    ScriptCache.put(cacheKey, price, 60);
    return price;
  } catch (e) {
    ScriptCache.put(failureCacheKey, "1", BITOPRO_TICKER_FAILURE_BACKOFF_SECONDS);
    logBitoProPrice_("warn", `BitoPro ticker ${normalizedPair} exception: ${e.message || e}`);
    return null;
  }
}

function extractBitoProTickerPayload_(json, normalizedPair) {
  const payload = json && json.data ? json.data : json;
  if (Array.isArray(payload)) {
    return payload.find(item => String(item && item.pair || "").trim().toLowerCase() === normalizedPair) || payload[0] || null;
  }

  return payload || null;
}

function fetchBitoProBitoPrice_(currency, bypassCache) {
  const targetCurrency = String(currency || "USD").trim().toUpperCase();

  if (targetCurrency === "TWD") {
    return fetchBitoProTickerLastPrice_("bito_twd", bypassCache);
  }

  const usdtPrice = fetchBitoProTickerLastPrice_("bito_usdt", bypassCache);
  if (usdtPrice !== null) return usdtPrice;

  const twdPrice = fetchBitoProTickerLastPrice_("bito_twd", bypassCache);
  if (twdPrice === null) return null;

  return convertBitoProTwdPriceToUsd_(twdPrice);
}

function convertBitoProTwdPriceToUsd_(twdPrice) {
  const price = parseBitoProPositiveNumber_(twdPrice);
  if (price === null) return null;

  try {
    if (typeof SettingsMatrixRepo === "undefined" || !SettingsMatrixRepo.lookupRate) return null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const twdToUsd = SettingsMatrixRepo.lookupRate(ss, "TWD", "USD", { allowLegacyName: true });
    const directRate = parseBitoProPositiveNumber_(twdToUsd && twdToUsd.value);
    if (directRate !== null) return price * directRate;

    const usdToTwd = SettingsMatrixRepo.lookupRate(ss, "USD", "TWD", { allowLegacyName: true });
    const inverseRate = parseBitoProPositiveNumber_(usdToTwd && usdToTwd.value);
    if (inverseRate !== null) return price / inverseRate;
  } catch (e) {
    logBitoProPrice_("warn", `BitoPro TWD->USD conversion failed: ${e.message || e}`);
  }

  return null;
}

function parseBitoProPositiveNumber_(value) {
  const num = parseFloat(value);
  return isFinite(num) && num > 0 ? num : null;
}

function logBitoProPrice_(level, message) {
  const moduleName = "Fetch_BitoProPrice";
  if (typeof LogService !== "undefined" && LogService[level]) {
    LogService[level](message, moduleName);
    return;
  }

  Logger.log(`[${moduleName}] ${message}`);
}
