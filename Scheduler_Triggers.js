/**
 * 排程器觸發設定
 * 用於管理所有自動化觸發程序
 */

function setupScheduledTriggers() {
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();

    // 讀取設定
    const mode = props.getProperty('SCHEDULER_MODE') || 'INTERVAL'; // INTERVAL or DAILY
    const hour = parseInt(props.getProperty('SCHEDULER_HOUR') || '9');
    const interval = SCHEDULER_CONFIG.INTERVAL_HOURS;

    // 清除舊的觸發器
    deleteAllSapTriggers_();

    // 建立新的觸發器
    if (SCHEDULER_CONFIG.IS_ENABLED) {
        if (mode === 'DAILY') {
            // Daily Mode: Run main sync once a day at specific hour
            ScriptApp.newTrigger('runAutomationMaster').timeBased().everyDays(1).atHour(hour).create();

            // Daily Snapshot (Fixed at 6 AM after US Market Close) linked to Daily Close Routine
            ScriptApp.newTrigger('runDailyCloseRoutine').timeBased().everyDays(1).atHour(6).create();
        } else {
            // Interval Mode: Run main sync every X hours
            ScriptApp.newTrigger('runAutomationMaster').timeBased().everyHours(interval).create();

            // Daily Snapshot (Fixed at 6 AM after US Market Close) linked to Daily Close Routine
            ScriptApp.newTrigger('runDailyCloseRoutine').timeBased().everyDays(1).atHour(6).create();
        }

        // System Health Check (Weekly)
        ScriptApp.newTrigger('updateDailyRoutine').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

        ui.alert(`[OK] 排程已設定。\n模式: ${mode}\n頻率/時間: ${mode === 'DAILY' ? hour + ':00' : '每 ' + interval + ' 小時'}`);
    } else {
        ui.alert("[INFO] 排程功能目前已停用 (SCHEDULER_CONFIG.IS_ENABLED = false)");
    }
}

function deleteAllSapTriggers_() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        const func = trigger.getHandlerFunction();
        if (func === 'runDailyInvestmentCheck' || func === 'autoRecordDailyValues' || func === 'runDailyCloseRoutine' || func === 'updateAllPrices' || func === 'runAutomationMaster' || func === 'updateDailyRoutine') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
}

function stopScheduler() {
    const ui = SpreadsheetApp.getUi();
    deleteAllSapTriggers_();
    PropertiesService.getScriptProperties().setProperty('SCHEDULER_ENABLED', 'false');
    ui.alert("[OFF] All automated schedules disabled.");
}

function showSystemDashboard() {
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Key Market Indicators');
    const d = sheet ? sheet.getDataRange().getValues() : [];
    const ind = {}; d.forEach(r => { if (r[0]) ind[r[0]] = r[1]; });

    const status = SCHEDULER_CONFIG.IS_ENABLED ? "[ACTIVE]" : "[OFF]";
    const mode = props.getProperty('SCHEDULER_MODE') || 'INTERVAL';
    const interval = SCHEDULER_CONFIG.INTERVAL_HOURS;
    const hour = props.getProperty('SCHEDULER_HOUR');

    let scheduleInfo = "";
    if (mode === 'DAILY') {
        scheduleInfo = `Daily at ${hour}:00`;
    } else {
        scheduleInfo = `Every ${interval} hour(s)`;
    }

    const email = props.getProperty('ADMIN_EMAIL') || "Not Set";
    const reserve = props.getProperty('TREASURY_RESERVE_TWD') || "100000";
    const btcHigh = ind.BTC_Recent_High || "No Data";

    let m = "--- [SAP v24.5 System Dashboard] ---\n\n";
    m += "1. System Status: " + status + "\n";
    m += "2. Schedule: " + scheduleInfo + "\n";
    m += "3. Admin Email: " + email + "\n";
    m += "4. Treasury Reserve: TWD " + parseInt(reserve).toLocaleString() + "\n";
    m += "5. BTC Reference High: $" + (isNaN(parseFloat(btcHigh)) ? btcHigh : parseFloat(btcHigh).toLocaleString()) + "\n\n";
    m += "Digital Sovereignty is non-negotiable.";

    ui.alert("SAP: System Audit", m, ui.ButtonSet.OK);
}
