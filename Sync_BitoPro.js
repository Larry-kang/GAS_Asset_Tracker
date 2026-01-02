// =======================================================
// --- BitoPro Service Layer (v3.0 - Unified Ledger) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getBitoProBalance() {
  const MODULE_NAME = "Sync_BitoPro";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const creds = Credentials.get('BITOPRO');
    const { apiKey, apiSecret } = creds;
    const baseUrl = 'https://api.bitopro.com/v3';

    if (!apiKey || !apiSecret) {
      SyncManager.log("ERROR", "Missing BitoPro API Keys", MODULE_NAME);
      return;
    }

    const json = fetchBitoProApi_(baseUrl, '/accounts/balance', 'GET', {}, apiKey, apiSecret);
    const assetList = [];

    if (json && json.data) {
      json.data.forEach(b => {
        const avail = parseFloat(b.available) || 0;
        const stake = parseFloat(b.stake) || 0;

        if (avail > 0) {
          assetList.push({ ccy: b.currency.toUpperCase(), amt: avail, type: 'Spot', status: 'Available' });
        }
        if (stake > 0) {
          assetList.push({ ccy: b.currency.toUpperCase(), amt: stake, type: 'Spot', status: 'Frozen' });
        }
      });

      SyncManager.log("INFO", `Collected ${assetList.length} asset entries from BitoPro.`, MODULE_NAME);
      SyncManager.updateUnifiedLedger(ss, "BitoPro", assetList);

    } else {
      SyncManager.log("ERROR", `BitoPro API Error: ${JSON.stringify(json)}`, MODULE_NAME);
    }
  });
}

// ... Helpers (English Logs) ...

function fetchBitoProApi_(baseUrl, endpoint, method, params, apiKey, apiSecret) {
  const nonce = Date.now() - 2000; // Time Skew Fix
  const payloadObj = { ...params, nonce: nonce };
  const payloadJson = JSON.stringify(payloadObj);
  const payloadBase64 = Utilities.base64Encode(payloadJson);
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_384, payloadBase64, apiSecret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  const options = {
    'method': method,
    'headers': {
      'X-BITOPRO-APIKEY': apiKey,
      'X-BITOPRO-PAYLOAD': payloadBase64,
      'X-BITOPRO-SIGNATURE': signature,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };
  if (method === 'POST') options.payload = payloadJson;

  try {
    const response = UrlFetchApp.fetch(baseUrl + endpoint, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    return { error: e.message };
  }
}
