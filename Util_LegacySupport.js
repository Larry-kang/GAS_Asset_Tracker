/**
 * SAP 舊版兼容支援: GET_PRICE
 * 為試算表公式提供向後兼容性。
 * 
 * @param {string} ticker - 資產代碼 (例如: "00713", "BTC")
 * @return {number} 從緩存或外部獲取的當前價格。
 * @customfunction
 */
function GET_PRICE(ticker) {
    if (!ticker) return null;
    const t = ticker.toString().trim().toUpperCase();

    try {
        // --- 策略 1: 從「價格暫存」讀取 (最快且最安全) ---
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // 優先嘗試英文 "Price Cache"，若無則嘗試中文 "價格暫存"
        const cacheSheet = ss.getSheetByName('Price Cache') || ss.getSheetByName('價格暫存');

        if (cacheSheet) {
            const data = cacheSheet.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
                if (data[i][0] && data[i][0].toString().toUpperCase() === t) {
                    const price = parseFloat(data[i][2]);
                    if (!isNaN(price) && price > 0) return price;
                }
            }
        }

        // --- 策略 2: 即時抓取 (若緩存未命中則啟動備援) ---
        // 台股偵測 (純數字)
        if (/^\d+$/.test(t)) {
            return typeof fetchStockPrice === 'function' ? fetchStockPrice(t) : null;
        }

        // 加密貨幣抓取 (內部備援: cryptoprices.cc -> CoinMarketCap)
        if (typeof fetchCryptoPrice === 'function') {
            const price = fetchCryptoPrice(t);
            if (price) return price;
        }

        // 最終備援: 股票抓取 (GoogleFinance 等)
        if (typeof fetchStockPrice === 'function') {
            return fetchStockPrice(t);
        }

        return null;
    } catch (e) {
        return "Err: " + e.toString().substring(0, 50);
    }
}
