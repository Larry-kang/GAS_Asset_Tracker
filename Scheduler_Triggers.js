/**
 * @OnlyCurrentDoc
 * SAP Dynamic Scheduler (v24.5) - Technical Debugging Edition
 */

const SCHEDULER_CONFIG = {
    get INTERVAL_HOURS() {
        const val = PropertiesService.getScriptProperties().getProperty('SCHEDULER_INTERVAL_HOURS');
        return val ? parseInt(val) : 24;
    },
    get IS_ENABLED() {
        return PropertiesService.getScriptProperties().getProperty('SCHEDULER_ENABLED') === 'true';
    }
};

function deleteAllSapTriggers_() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        const func = t.getHandlerFunction();
        if (func === 'runDailyInvestmentCheck' || func === 'autoRecordDailyValues' || func === 'updateAllPrices' || func === 'runAutomationMaster' || func === 'updateDailyRoutine') {
            ScriptApp.deleteTrigger(t);
        }
    });
}

function startScheduler() {
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();

    // Choose Mode
    const modeRes = ui.alert(
        "[Scheduler] 戰略排程設定",
        "請選擇自動化報告的觸發模式：\n\n[是 (Yes)] 指定每日固定時間 (例如: 每天早上 6 點)\n[否 (No)] 指定間隔循環 (例如: 每 8 小時一次)",
        ui.ButtonSet.YES_NO
    );

    let isDailyMode = (modeRes == ui.Button.YES);

    if (isDailyMode) {
        // Daily Mode: Ask for Hour
        const hourRes = ui.prompt("[Scheduler] 指定執行時間", "請輸入每天執行的小時 (0-23):\n(建議: 輸入 6 代表早上 6 點，美股收盤後)", ui.ButtonSet.OK_CANCEL);
        if (hourRes.getSelectedButton() == ui.Button.OK) {
            let hour = parseInt(hourRes.getResponseText());
            if (isNaN(hour) || hour < 0 || hour > 23) hour = 6; // Default to 6 AM

            props.setProperty('SCHEDULER_MODE', 'DAILY');
            props.setProperty('SCHEDULER_HOUR', hour.toString());
            props.setProperty('SCHEDULER_ENABLED', 'true');
            deleteAllSapTriggers_();

            // Daily Trigger at specific hour
            ScriptApp.newTrigger('runAutomationMaster').timeBased().everyDays(1).atHour(hour).create();
            // Snapshot Trigger (Fixed at 1 AM)
            ScriptApp.newTrigger('autoRecordDailyValues').timeBased().everyDays(1).atHour(1).create();

            ui.alert(`[OK] 排程已啟動。\n- 戰略報告: 每天 ${hour}:00 - ${hour}:59 之間執行\n- 資產快照: 每天 01:00 AM`);
        }
    } else {
        // Interval Mode: Ask for Duration
        const response = ui.prompt("[Scheduler] 設定循環間隔", "請輸入執行間隔 (小時):", ui.ButtonSet.OK_CANCEL);
        if (response.getSelectedButton() == ui.Button.OK) {
            let interval = parseInt(response.getResponseText());
            if (isNaN(interval) || interval <= 0) interval = 24;

            props.setProperty('SCHEDULER_MODE', 'INTERVAL');
            props.setProperty('SCHEDULER_INTERVAL_HOURS', interval.toString());
            props.setProperty('SCHEDULER_ENABLED', 'true');
            deleteAllSapTriggers_();

            // Interval Trigger
            ScriptApp.newTrigger('runAutomationMaster').timeBased().everyHours(interval).create();
            // Snapshot Trigger
            ScriptApp.newTrigger('autoRecordDailyValues').timeBased().everyDays(1).atHour(1).create();

            ui.alert(`[OK] 排程已啟動。\n- 戰略報告: 每 ${interval} 小時執行一次\n- 資產快照: 每天 01:00 AM`);
        }
    }
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
