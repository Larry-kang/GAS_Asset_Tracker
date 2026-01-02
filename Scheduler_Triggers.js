/**
 * Scheduler_Triggers.js
 * Handles the creation and management of time-based triggers for SAP v24.5.
 */

/**
 * Initializes high-frequency monitoring triggers.
 * Creates two main triggers:
 * 1. Strategic Monitor - runs every 30 minutes
 * 2. Daily Close Routine - runs daily at 1 AM
 * @public
 */
function setupScheduledTriggers() {
  // Try to get UI environment, fallback to Logger if running in standalone/headless
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log("[Info] UI environment not available. Using Logger instead.");
  }

  const props = PropertiesService.getScriptProperties();

  // 1. Clear Existing Triggers to prevent duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // 2. Create "Heartbeat" Trigger (Every 30 Minutes)
  // Calls runAutomationMaster -> runStrategicMonitor
  ScriptApp.newTrigger('runAutomationMaster')
    .timeBased()
    .everyMinutes(30)
    .create();

  // 3. Create "Daily Close" Trigger (Daily at 1 AM)
  // Calls runDailyCloseRoutine -> Snapshot + Daily Email
  ScriptApp.newTrigger('runDailyCloseRoutine')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  const msg = "系統升級完成 (SAP v24.5)\n\n" +
    "- 戰略監控 (Strategic Monitor): 每 30 分鐘執行一次 (Dashboard 寫回 + 警報)\n" +
    "- 每日結算 (Daily Close): 每天凌晨 01:00 (快照 + 日報)\n\n" +
    "Digital Sovereignty Established.";

  if (ui) {
    ui.alert(msg);
  } else {
    Logger.log(msg);
  }
}

/**
 * Emergency stop function - deletes all project triggers.
 * Provides user confirmation dialog before stopping scheduled automation.
 * Displays count of deleted triggers and instructions for re-enabling.
 * @public
 */
function stopScheduler() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '⚠️ 確認停止排程？',
    '此操作將刪除所有自動觸發器，包括：\n\n' +
    '- 戰略監控 (每 30 分鐘)\n' +
    '- 每日結算 (凌晨 01:00)\n\n' +
    '⚡ 停止後系統將不再自動更新。\n\n' +
    '確定要繼續嗎？',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    const count = triggers.length;

    triggers.forEach(t => ScriptApp.deleteTrigger(t));

    ui.alert(
      '✅ 已停止所有排程',
      `成功刪除 ${count} 個觸發器。\n\n` +
      '如需重新啟用，請執行：\n' +
      '「系統維運管理」→「核心引導設定」',
      ui.ButtonSet.OK
    );

    console.log(`[Scheduler] All ${count} triggers deleted by user request.`);
  } else {
    console.log('[Scheduler] Stop operation cancelled by user.');
  }
}