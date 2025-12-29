/**
 * Core_MainMaster.js
 * 
 * [Automation Master]
 * Main driver for the scheduler. Integrates price updates, balance sync, and strategic check.
 */
function runAutomationMaster() {
    const context = "MasterLoop";
    console.log(`[${context}] Starting Automation V24.5 Routine...`);

    try {
        // 1. Update Market Prices (Crypto + Stock)
        // Assumes these functions exist in Fetch_CryptoPrice.js / Fetch_StockPrice.js mechanism
        // In this upgraded version, we might just fetch what's needed for the dashboard.
        // Calling existing price update hooks if available.
        try {
            if (typeof updateAllPrices === 'function') updateAllPrices();
        } catch (e) { console.warn("Price update skipped/failed: " + e.message); }

        // 2. Execute Strategic Monitor (The 30-min heart)
        // This reads indicators, checks logic, and updates dashboard.
        // It assumes Core_StrategicEngine.js is loaded.
        if (typeof runStrategicMonitor === 'function') {
            runStrategicMonitor();
        } else {
            throw new Error("runStrategicMonitor function not found in Core_StrategicEngine.");
        }

        console.log(`[${context}] Routine Completed.`);
    } catch (e) {
        console.error(`[${context}] CRITICAL FAILURE: ${e.toString()}`);
        // Optional: Email admin on system failure?
    }
}

/**
 * [Daily Close Routine]
 * Runs once a day (e.g. 1 AM) to record snapshot history.
 */
function runDailyCloseRoutine() {
    const context = "DailySnapshot";
    console.log(`[${context}] Recording Daily History...`);

    try {
        // Force a full check first
        runAutomationMaster();

        // Wait for spreadsheet calculation propagation
        Utilities.sleep(5000);

        // Record snapshot
        if (typeof autoRecordDailyValues === 'function') {
            autoRecordDailyValues();
        }

        // Send Daily Report (Email)
        // This calls the detailed daily report logic (Core_StrategicEngine)
        if (typeof runDailyInvestmentCheck === 'function') {
            runDailyInvestmentCheck();
        }

    } catch (e) {
        console.error(`[${context}] Failed: ${e.toString()}`);
        const email = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
        if (email) MailApp.sendEmail(email, "[System Error] Daily Close Failed", e.toString());
    }
}
