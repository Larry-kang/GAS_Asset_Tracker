function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('SAP 指揮中心');

  // --- [Command] 戰略指令 ---
  menu.addItem('執行即時戰略報告', 'showStrategicReportUI')
    .addItem('一鍵全系統同步', 'runAutomationMaster')
    .addSeparator();

  // --- [Control] 系統維運管理 ---
  const sysMenu = ui.createMenu('系統維運管理')
    .addItem('執行系統健康檢查', 'runSystemHealthCheck')
    .addItem('查看系統日誌', 'openLogSheet')
    .addSeparator()
    .addItem('核心引導設定', 'setup')
    .addItem('設定 CI/CD (GitHub)', 'setup_cicd')
    .addSeparator()
    .addItem('強制執行每日快照', 'Record_DailySnapshot');

  // --- [Data] 進階同步工具 ---
  const dataMenu = ui.createMenu('進階同步工具')
    .addItem('同步 - 幣安餘額', 'getBinanceBalance')
    .addItem('同步 - OKX 餘額', 'getOkxBalance')
    .addItem('同步 - 幣託餘額', 'getBitoProBalance')
    .addItem('同步 - 派網餘額', 'getPionexBalance')
    .addSeparator()
    .addItem('更新 - 僅市價行情', 'updateAllPrices')
    .addItem('更新 - 資產與貨幣對', 'syncAssets');

  menu.addSubMenu(sysMenu)
    .addSubMenu(dataMenu)
    .addSeparator()
    .addItem('緊急停止觸發器', 'stopScheduler')
    .addToUi();
}

/**
 * 快速跳轉至日誌工作表
 */
function openLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("System_Logs");
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert("找不到日誌工作表。請先執行同步。");
  }
}

/**
 * 每日快照持久化包裝器
 */
function Record_DailySnapshot() {
  if (typeof autoRecordDailyValues === 'function') {
    autoRecordDailyValues();
  } else {
    SpreadsheetApp.getUi().alert("找不到 autoRecordDailyValues 函數。");
  }
}

/**
 * CI/CD 設定引導
 * 由於 CI/CD 是在 GitHub Actions 執行，此處僅跳出指引視窗
 */
function setup_cicd() {
  const ui = SpreadsheetApp.getUi();
  const msg = "CI/CD 自動化部署設定須在本地端執行。\n\n" +
    "請在您的電腦上執行以下 PowerShell 腳本：\n" +
    "> .\\setup_cicd.ps1\n\n" +
    "該腳本將自動讀取您的 .clasprc.json 憑證並加密上傳至 GitHub Secrets，" +
    "完成後即可啟用 GitHub Actions 自動部署功能。";
  ui.alert("GitHub CI/CD 設定指引", msg, ui.ButtonSet.OK);
}