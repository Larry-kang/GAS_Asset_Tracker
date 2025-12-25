// =======================================================
// --- Binance Balance Sync (v7.1 - Debugging Edition) ---
// =======================================================

function getBinanceBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Binance Balance');
  if (!sheet) {
    sheet = ss.insertSheet('Binance Balance');
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('BINANCE_API_KEY');
  const apiSecret = props.getProperty('BINANCE_API_SECRET');
  const baseUrl = props.getProperty('TUNNEL_URL');
  const proxyPassword = props.getProperty('PROXY_PASSWORD');

  if (!apiKey || !apiSecret) {
    ss.toast("[ERROR] Binance API Key or Secret not set.");
    return;
  }
  if (!baseUrl || !proxyPassword) {
    ss.toast("[WAIT] Connection failure: Waiting for PC tunnel report...");
    return;
  }

  const combinedTotals = new Map();
  const logMessages = [];

  try {
    // --- Step A: Spot Balances ---
    ss.toast('Fetching Binance Spot balances...');
    const spotRes = fetchBinanceApi_(baseUrl, '/api/v3/account', {}, apiKey, apiSecret, proxyPassword);

    if (spotRes.code === "0" && spotRes.data && spotRes.data.balances) {
      logMessages.push("Spot: Success");
      spotRes.data.balances.forEach(b => {
        if (b.asset.startsWith('LD')) return;
        const total = parseFloat(b.free) + parseFloat(b.locked);
        if (total > 0) {
          const currentTotal = combinedTotals.get(b.asset) || 0;
          combinedTotals.set(b.asset, currentTotal + total);
        }
      });
    } else {
      logMessages.push(`Spot Error: ${spotRes.msg || 'Unknown'}`);
    }

    // --- Step B: Simple Earn Balances ---
    ss.toast('Fetching Binance Earn balances...');
    const earnRes = fetchBinanceApi_(baseUrl, '/sapi/v1/simple-earn/flexible/position', { size: 100 }, apiKey, apiSecret, proxyPassword);

    if (earnRes.code === "0" && earnRes.data && Array.isArray(earnRes.data.rows)) {
      logMessages.push("Earn: Success");
      earnRes.data.rows.forEach(p => {
        const total = parseFloat(p.totalAmount);
        if (total > 0) {
          const currentTotal = combinedTotals.get(p.asset) || 0;
          combinedTotals.set(p.asset, currentTotal + total);
        }
      });
    } else {
      logMessages.push(`Earn Info: ${earnRes.msg || 'No Position'}`);
    }

    // --- Step C: Write to Sheet ---
    const sheetData = [];
    combinedTotals.forEach((val, asset) => {
      if (val > 0.00000001) {
        sheetData.push([asset, val]);
      }
    });

    sheetData.sort((a, b) => b[1] - a[1]);

    sheet.getRange('A2:C').clearContent();
    if (sheetData.length > 0) {
      sheet.getRange(2, 1, sheetData.length, 2).setValues(sheetData);
      sheet.getRange(2, 3).setValue(new Date());
    }

    ss.toast(`Binance sync complete! (${sheetData.length} assets found)`);

  } catch (e) {
    Logger.log(`Critical Error: ${e.toString()}`);
    ss.toast(`Execution Error: ${e.message}`);
  }
}

/**
 * Core Request Helper
 */
function fetchBinanceApi_(baseUrl, endpoint, params, apiKey, apiSecret, proxyPassword) {
  const timestamp = new Date().getTime();
  params.timestamp = timestamp;
  params.recvWindow = 10000;

  const queryString = Object.keys(params).map(key => {
    return `${key}=${encodeURIComponent(params[key])}`;
  }).join('&');

  const signature = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    queryString,
    apiSecret
  ).map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');

  const url = `${baseUrl}${endpoint}?${queryString}&signature=${signature}`;

  const options = {
    'method': 'GET',
    'headers': {
      'X-MBX-APIKEY': apiKey,
      'Content-Type': 'application/json',
      'x-proxy-auth': proxyPassword
    },
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
    const code = response.getResponseCode();

    if (code === 403 && text.includes('Access Denied')) {
      return { code: "-403", msg: "Proxy Auth Failed" };
    }

    let json;
    try {
      json = JSON.parse(text);
    } catch (e) {
      return { code: "-2", msg: "Invalid JSON response", raw: text };
    }

    if (json.msg || (json.code && json.code !== 0)) {
      return { code: json.code || "-999", msg: json.msg || "API Error", data: json };
    }

    return { code: "0", data: json };

  } catch (e) {
    return { code: "-1", msg: "Connection Error: " + e.message };
  }
}