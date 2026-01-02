// =======================================================
// --- Pionex Service Layer (v3.0 - Unified Ledger) ---
// --- Refactored to use SyncManager.updateUnifiedLedger
// =======================================================

function getPionexBalance() {
  const MODULE_NAME = "Sync_Pionex";

  SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const creds = Credentials.get('PIONEX');
    const { apiKey, apiSecret } = creds;
    const baseUrl = 'https://api.pionex.com';

    if (!apiKey || !apiSecret) {
      SyncManager.log("ERROR", "Missing Pionex API Keys", MODULE_NAME);
      return;
    }

    const json = fetchPionexApi_(baseUrl, '/api/v1/account/balances', {}, apiKey, apiSecret);
    const assetList = [];

    if (json.result === true && json.data && json.data.balances) {
      json.data.balances.forEach(b => {
        const free = parseFloat(b.free) || 0;
        const frozen = parseFloat(b.frozen) || 0;

        // Available
        if (free > 0) {
          assetList.push({ ccy: b.coin, amt: free, type: 'Spot', status: 'Available' });
        }
        // Frozen
        if (frozen > 0) {
          assetList.push({ ccy: b.coin, amt: frozen, type: 'Spot', status: 'Frozen' });
        }
      });

      SyncManager.log("INFO", `Collected ${assetList.length} asset entries from Pionex.`, MODULE_NAME);
      SyncManager.updateUnifiedLedger(ss, "Pionex", assetList);

    } else {
      SyncManager.log("ERROR", `Pionex API Error: ${JSON.stringify(json)}`, MODULE_NAME);
    }
  });
}

// ... Helpers (English Logs) ...

function fetchPionexApi_(baseUrl, endpoint, params, apiKey, apiSecret) {
  const method = 'GET';
  const timestamp = new Date().getTime();
  const signatureParams = { ...params, timestamp: timestamp };
  const sortedKeys = Object.keys(signatureParams).sort();
  const queryString = sortedKeys.map(key => key + '=' + signatureParams[key]).join('&');
  const fullPath = endpoint + '?' + queryString;
  const url = baseUrl + fullPath;
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, method + fullPath, apiSecret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  const options = {
    'method': method,
    'headers': {
      'PIONEX-KEY': apiKey,
      'PIONEX-SIGNATURE': signature,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const txt = response.getContentText();
    return txt.startsWith('{') ? JSON.parse(txt) : { result: false, message: "Non-JSON: " + txt };
  } catch (e) {
    return { result: false, message: "Fetch error: " + e.message };
  }
}