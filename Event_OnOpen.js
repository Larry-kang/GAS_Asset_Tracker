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
    .addSeparator() // 分隔線
    .addItem('設定 Discord Webhook', 'uiSetupDiscord');

  menu.addSeparator();

  // --- [Control] 系統維運管理 ---
  const sysMenu = ui.createMenu('系統維運管理')
    .addItem('執行系統健康檢查', 'runSystemHealthCheck')
    .addItem('查看系統日誌', 'openLogSheet')
    .addSeparator()
    .addItem('核心引導設定', 'setup')
    .addSeparator()
    .addItem('強制執行每日快照', 'Record_DailySnapshot');

  // --- [Data] 進階同步工具 ---
  const dataMenu = ui.createMenu('進階同步工具');
  ExchangeRegistry.getActive().forEach(function (entry) {
    dataMenu.addItem(`同步 - ${entry.displayName}餘額`, entry.functionName);
  });
  dataMenu
    .addSeparator()
    .addItem('更新 - 僅市價行情', 'updateAllPrices');

  const guardedTestMenu = createGuardedStagingTestMenu_(ui);

  menu.addSubMenu(sysMenu)
    .addSubMenu(dataMenu);

  if (guardedTestMenu) {
    menu.addSubMenu(guardedTestMenu);
  }

  menu.addSeparator()
    .addItem('緊急停止觸發器', 'stopScheduler')
    .addToUi();
}

function createGuardedStagingTestMenu_(ui) {
  if (typeof SafeLedgerTestHarness === 'undefined' ||
    !SafeLedgerTestHarness ||
    typeof SafeLedgerTestHarness.getConfig_ !== 'function') {
    return null;
  }

  const config = SafeLedgerTestHarness.getConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!config.enabled || !config.spreadsheetId || ss.getId() !== config.spreadsheetId) {
    return null;
  }

  return ui.createMenu('Staging 驗證')
    .addItem('執行 Guarded 測試套件', 'runSafeLedgerSyncTestSuiteUi');
}

/**
 * Opens the System Logs sheet for quick access.
 * Displays an alert if the sheet is not found.
 * @public
 */
function openLogSheet() {
  LogService.info('User accessed System Logs sheet', 'UI:openLogSheet');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LogService.SHEET_NAME);
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
 * UI to setup Discord Webhook URL
 * @public
 */
function uiSetupDiscord() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '設定 Discord Webhook',
    '請貼上您的 Webhook URL (留空則停用)：',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const url = result.getResponseText().trim();
    if (url) {
      Settings.set('DISCORD_WEBHOOK_URL', url);
      ui.alert('✅ 設定成功！', 'Discord 通知已啟用。\n系統將優先使用 Discord 發送警報。', ui.ButtonSet.OK);
    } else {
      Settings.set('DISCORD_WEBHOOK_URL', '');
      ui.alert('⚠️ 已停用', 'Discord Webhook 已清除。\n系統將降級回復為 Email 通知模式。', ui.ButtonSet.OK);
    }
  }
}
