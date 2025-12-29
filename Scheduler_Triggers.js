/**
 * Scheduler_Triggers.js
 * Handles the creation and management of time-based triggers for SAP v24.5.
 */

function setupScheduledTriggers() {
    const ui = SpreadsheetApp.getUi();

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

    ui.alert("系統升級完成 (SAP v24.5)\n\n" +
        "- 戰略監控 (Strategic Monitor): 每 30 分鐘執行一次 (Dashboard 寫回 + 警報)\n" +
        "- 每日結算 (Daily Close): 每天凌晨 01:00 (快照 + 日報)\n\n" +
        "Digital Sovereignty Established.");
}
