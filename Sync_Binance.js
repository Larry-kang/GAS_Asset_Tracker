// =======================================================
// --- Binance 業務邏輯層 (Service Layer) ---
// --- 專注於執行查詢與寫入試算表，不處理 Webhook 路由
// =======================================================

/**
 * [核心功能] 獲取 Binance 餘額
 * 此函式由 WebhookHandler 或 每日排程 呼叫
 */
function getBinanceBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Binance Balance');
  if (!sheet) sheet = ss.insertSheet('Binance Balance');

  // 1. 讀取設定 (資料由 WebhookHandler 更新)
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('BINANCE_API_KEY');
  const apiSecret = props.getProperty('BINANCE_API_SECRET');
  const baseUrl = props.getProperty('TUNNEL_URL');
  const proxyPassword = props.getProperty('PROXY_PASSWORD');

  // 防呆與連線檢查
  if (!apiKey || !apiSecret) {
    ss.toast("❌ 錯誤：請先設定 BINANCE_API_KEY 與 SECRET");
    return;
  }

  // 若電腦沒開 (沒有網址)，直接中止，保護舊資料
  if (!baseUrl || !proxyPassword) {
    console.log("⚠️ 跳過：尚未接收到 Tunnel 網址 (電腦可能未開機)，保留舊資料。");
    // ss.toast("⚠️ 跳過更新：等待電腦連線..."); // 註解掉以免每日排程一直跳通知
    return;
  }

  const combinedTotals = new Map();
  const logMessages = [];
  let isSpotSuccess = false;

  try {
    // --- 步驟 A: Spot (現貨) ---
    const spotRes = fetchBinanceApi_(baseUrl, '/api/v3/account', {}, apiKey, apiSecret, proxyPassword);

    if (spotRes.code === "0" && spotRes.data && spotRes.data.balances) {
      logMessages.push("Spot: OK");
      isSpotSuccess = true;
      spotRes.data.balances.forEach(b => {
        if (b.asset.startsWith('LD')) return; // 忽略 LD
        const total = parseFloat(b.free) + parseFloat(b.locked);
        if (total > 0) {
          const currentTotal = combinedTotals.get(b.asset) || 0;
          combinedTotals.set(b.asset, currentTotal + total);
        }
      });
    } else {
      console.log(`Spot Failed: ${spotRes.msg} (可能電腦關機或 Tunnel 失效)`);
      return; // 失敗直接中止
    }

    // --- 步驟 B: Earn (理財) ---
    const earnRes = fetchBinanceApi_(baseUrl, '/sapi/v1/simple-earn/flexible/position', { size: 100 }, apiKey, apiSecret, proxyPassword);

    if (earnRes.code === "0" && earnRes.data && Array.isArray(earnRes.data.rows)) {
      logMessages.push("Earn: OK");
      earnRes.data.rows.forEach(p => {
        const total = parseFloat(p.totalAmount);
        if (total > 0) {
          const currentTotal = combinedTotals.get(p.asset) || 0;
          combinedTotals.set(p.asset, currentTotal + total);
        }
      });
    }

    // --- 步驟 C: 寫入試算表 ---
    console.log("===== Binance Update Log =====");
    console.log(logMessages.join("\n"));

    const sheetData = [];
    combinedTotals.forEach((val, asset) => {
      if (val > 0.00000001) sheetData.push([asset, val]);
    });
    sheetData.sort((a, b) => b[1] - a[1]);

    if (sheetData.length > 0) {
      sheet.getRange('A2:C').clearContent();
      sheet.getRange(2, 1, sheetData.length, 2).setValues(sheetData);
      sheet.getRange(2, 3).setValue(new Date());
      ss.toast(`Binance 更新成功！(${sheetData.length} 幣種)`);
    }

  } catch (e) {
    console.log(`Critical Error: ${e.toString()}`);
  }
}

/**
 * [私有工具] 連線核心
 */
function fetchBinanceApi_(baseUrl, endpoint, params, apiKey, apiSecret, proxyPassword) {
  const timestamp = new Date().getTime();
  params.timestamp = timestamp;
  params.recvWindow = 10000;

  const queryString = Object.keys(params).map(key => `${key}=${encodeURIComponent(params[key])}`).join('&');
  const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, queryString, apiSecret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

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
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 403) return { code: "-403", msg: "Proxy Auth Failed" };

    let json;
    try { json = JSON.parse(text); } catch (e) { return { code: "-2", msg: "Invalid JSON" }; }

    if (json.msg || (json.code && json.code !== 0)) {
      return { code: json.code || "-999", msg: json.msg, data: json };
    }
    return { code: "0", data: json };

  } catch (e) {
    return { code: "-1", msg: "Connection Error: " + e.message };
  }
}