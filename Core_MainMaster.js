/**
 * Core_MainMaster.js
 * 
 * [Automation Master]
 * Main driver for the scheduler. Integrates price updates, balance sync, and strategic check.
 */
function runAutomationMaster() {
    const context = "MasterLoop";
    console.log(`[${context}] Starting Automation ${Config.VERSION} Routine...`);

    try {
        // 1. Update Market Prices (Crypto + Stock)
        try {
            if (typeof updateAllPrices === 'function') updateAllPrices();
        } catch (e) { console.warn("Price update skipped/failed: " + e.message); }

        // [NEW] 2. Sync Asset Balances
        // Ensure balances are fresh BEFORE running strategy
        syncAllAssets_();

        // 3. Execute Strategic Monitor (The 30-min heart)
        // This reads indicators, checks logic, and updates dashboard.
        if (typeof runStrategicMonitor === 'function') {
            runStrategicMonitor();
        } else {
            throw new Error("runStrategicMonitor function not found in Core_StrategicEngine.");
        }

        console.log(`[${context}] Routine Completed.`);
    } catch (e) {
        console.error(`[${context}] CRITICAL FAILURE: ${e.toString()}`);
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
        // Force a full check first (Syncs Assets + Updates Dashboard)
        runAutomationMaster();

        // Wait for spreadsheet calculation propagation
        Utilities.sleep(5000);

        // [NEW] Build Context for Snapshot
        // We rebuild it to ensure we capture the state exactly as it is after sync
        let snapshotContext = null;
        try {
            if (typeof buildContext === 'function') {
                snapshotContext = buildContext();
            }
        } catch (e) {
            console.warn("[Snapshot] Context build failed, falling back to sheet read: " + e.message);
        }

        // Record snapshot with Context
        if (typeof autoRecordDailyValues === 'function') {
            autoRecordDailyValues(snapshotContext);
        }

        // Send Daily Report (Email)
        if (typeof runDailyInvestmentCheck === 'function') {
            runDailyInvestmentCheck();
        }

    } catch (e) {
        console.error(`[${context}] Failed: ${e.toString()}`);
        const email = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
        if (email) MailApp.sendEmail(email, "[System Error] Daily Close Failed", e.toString());
    }
}

/**
 * [Helper] Sync all exchange balances
 * Orchestrates all individual sync modules safely.
 */
function syncAllAssets_() {
    console.log("[Master] Syncing Assets...");

    // [Clean Logs] Keep only last 7 days
    if (typeof LogService !== 'undefined' && LogService.cleanupOldLogs) {
        LogService.cleanupOldLogs(7);
    }

    // Binance
    try {
        if (typeof getBinanceBalance === 'function') {
            getBinanceBalance();
        } else {
            console.warn("getBinanceBalance not found");
        }
    } catch (e) { console.error("Binance Sync Failed", e); }

    // OKX
    try {
        if (typeof getOkxBalance === 'function') {
            getOkxBalance();
        } else {
            console.warn("getOkxBalance not found");
        }
    } catch (e) { console.error("OKX Sync Failed", e); }

    // Pionex
    try {
        if (typeof getPionexBalance === 'function') getPionexBalance();
    } catch (e) { console.error("Pionex Sync Failed", e); }

    // BitoPro
    try {
        if (typeof getBitoProBalance === 'function') getBitoProBalance();
    } catch (e) { console.error("BitoPro Sync Failed", e); }
}
