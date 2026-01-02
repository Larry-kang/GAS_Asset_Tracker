/**
 * Spreadsheet onOpen trigger - creates custom menu on spreadsheet load.
 * Automatically called by Google Sheets when the spreadsheet is opened.
 * @public
 */
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
 * Opens the System Logs sheet for quick access.
 * Displays an alert if the sheet is not found.
 * @public
 */
function openLogSheet() {
  LogService.info('User accessed System Logs sheet', 'UI:openLogSheet');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("System_Logs");
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    LogService.warn('System_Logs sheet not found', 'UI:openLogSheet');
    SpreadsheetApp.getUi().alert("找不到日誌工作表。請先執行同步。");
  }
}

/**
 * Wrapper function for manual daily snapshot execution.
 * Calls the core snapshot function with error handling and user feedback.
 * @public
 */
function Record_DailySnapshot() {
  const ui = SpreadsheetApp.getUi();
  const context = 'UI:Record_DailySnapshot';
  LogService.info('User triggered manual daily snapshot', context);

  try {
    if (typeof autoRecordDailyValues === 'function') {
      autoRecordDailyValues();
      ui.alert(
        '✅ 快照記錄完成',
        '每日資產快照已成功寫入 Daily History 工作表。',
        ui.ButtonSet.OK
      );
      LogService.info('Manual snapshot completed successfully', context);
    } else {
      throw new Error("找不到 autoRecordDailyValues 函數");
    }
  } catch (e) {
    LogService.error(`Execution failed: ${e.message}`, context);
    ui.alert(
      '❌ 快照記錄失敗',
      `執行時發生錯誤：\n${e.message}\n\n請檢查 Record_DailySnapshot.js 是否正確載入。`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * CI/CD 設定引導
 * 由於 CI/CD 是在 GitHub Actions 執行，此處僅跳出指引視窗
 */
function setup_cicd() {
  LogService.info('User requested CI/CD setup guide', 'UI:setup_cicd');
  const ui = SpreadsheetApp.getUi();
  const msg = "CI/CD 自動化部署設定須在本地端執行。\n\n" +
    "請在您的電腦上執行以下 PowerShell 腳本：\n" +
    "> .\\setup_cicd.ps1\n\n" +
    "該腳本將自動讀取您的 .clasprc.json 憑證並加密上傳至 GitHub Secrets，" +
    "完成後即可啟用 GitHub Actions 自動部署功能。";
  ui.alert("GitHub CI/CD 設定指引", msg, ui.ButtonSet.OK);
}
