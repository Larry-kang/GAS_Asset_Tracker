// 檔案: WebhookHandler.gs

/**
 * [Router] 統一入口，根據 action 分發任務
 */
function doPost(e) {
  try {
    // 防止 e 為 undefined (例如直接在編輯器按執行時)
    if (!e || !e.postData) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "No Post Data" }));
    }

    const data = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    
    // 1. 全局驗證
    const savedPassword = props.getProperty('PROXY_PASSWORD');
    // 若尚未設定密碼，則允許第一次連線 (或您可以手動先去設定好)
    if (savedPassword && data.password !== savedPassword) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "Auth Failed" }));
    }

    // 2. 路由分發
    switch (data.action) {
      case 'update_tunnel_url':
        return handleTunnelUpdate(data, props);
        
      case 'trigger_balance_update': 
        return handleForceUpdate(data);
        
      case 'log_client_error':       
        return handleLogClientError(data);

      default:
        return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: "Unknown Action" }));
    }
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", msg: err.toString() }));
  }
}

// --- Controllers / Handlers ---

function handleTunnelUpdate(data, props) {
  props.setProperty('TUNNEL_URL', data.url);
  
  // ⭐ 修改建議：總是更新密碼，以保持與電腦端同步
  // 因為這一步是在全域驗證通過後(或第一次)才執行的，所以是安全的
  if (data.password) {
     props.setProperty('PROXY_PASSWORD', data.password);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", msg: "URL & Password Updated" }));
}

function handleForceUpdate(data) {
  // 這裡呼叫 Binance.gs 裡面的函式
  getBinanceBalance(); 
  return ContentService.createTextOutput(JSON.stringify({ status: "success", msg: "Balance Update Triggered" }));
}

function handleLogClientError(data) {
  Logger.log("[Client Error] " + data.message);
  return ContentService.createTextOutput(JSON.stringify({ status: "success", msg: "Error Logged" }));
}

/**
 * [手動] 強制遠端電腦重啟 Bridge
 * 在選單點擊後，會發送指令給 PowerShell，約 2-5 秒後生效
 */
function triggerRemoteRestart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const baseUrl = props.getProperty('TUNNEL_URL');
  const proxyPassword = props.getProperty('PROXY_PASSWORD');

  if (!baseUrl || !proxyPassword) {
    ss.toast("❌ 失敗：找不到連線資訊，電腦可能未開機。");
    return;
  }

  const url = `${baseUrl}/restart`;
  const options = {
    'method': 'POST', // 使用 POST
    'headers': {
      'x-proxy-auth': proxyPassword 
    },
    'muteHttpExceptions': true
  };

  try {
    ss.toast("⏳ 正在發送重啟指令...");
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      ss.toast("✅ 指令已發送！PowerShell 將在幾秒內重啟。");
      console.log("Remote Restart Triggered Successfully.");
    } else {
      ss.toast(`⚠️ 發送失敗 (Code: ${code})`);
      console.log(`Remote Restart Failed: ${response.getContentText()}`);
    }

  } catch (e) {
    ss.toast("❌ 連線錯誤：無法連接到電腦。");
    console.error(e);
  }
}