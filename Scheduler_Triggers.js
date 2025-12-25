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
    const response = ui.prompt("[Scheduler] Set Frequency", "Enter automated update interval (Hours):", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
        let interval = parseInt(response.getResponseText());
        if (isNaN(interval) || interval <= 0) interval = 24;
        props.setProperty('SCHEDULER_INTERVAL_HOURS', interval.toString());
        props.setProperty('SCHEDULER_ENABLED', 'true');
        deleteAllSapTriggers_();

        // Dynamic Sync Trigger (Prices/Balances/Strategic Check)
        ScriptApp.newTrigger('runAutomationMaster').timeBased().everyHours(interval).create();

        // Fixed Daily Snapshot Trigger (Record Assets - 1:00 AM)
        ScriptApp.newTrigger('autoRecordDailyValues').timeBased().everyDays(1).atHour(1).create();

        ui.alert("[OK] All schedulers started.\n- Master Sync: Every " + interval + " hour(s)\n- Daily Snapshot: Every day at 01:00 AM");
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
    const interval = SCHEDULER_CONFIG.INTERVAL_HOURS;
    const email = props.getProperty('ADMIN_EMAIL') || "Not Set";
    const reserve = props.getProperty('TREASURY_RESERVE_TWD') || "100000";
    const btcHigh = ind.BTC_Recent_High || "No Data";

    let m = "--- [SAP v24.5 System Dashboard] ---\n\n";
    m += "1. System Status: " + status + "\n";
    m += "2. Update Interval: Every " + interval + " hour(s)\n";
    m += "3. Admin Email: " + email + "\n";
    m += "4. Treasury Reserve: TWD " + parseInt(reserve).toLocaleString() + "\n";
    m += "5. BTC Reference High: $" + (isNaN(parseFloat(btcHigh)) ? btcHigh : parseFloat(btcHigh).toLocaleString()) + "\n\n";
    m += "Digital Sovereignty is non-negotiable.";

    ui.alert("SAP: System Audit", m, ui.ButtonSet.OK);
}
