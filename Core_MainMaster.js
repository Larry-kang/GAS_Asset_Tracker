/**
 * Core_MainMaster.js
 * 
 * [Automation Master]
 * Main driver for the scheduler. Integrates price updates, balance sync, and strategic check.
 */
function runAutomationMaster(options) {
    const opts = options || {};
    const context = "MasterLoop";
    const result = {
        status: "RUNNING",
        ok: false,
        fatal: false,
        message: ""
    };
    console.log(`[${context}] Starting Automation ${Config.VERSION} Routine...`);

    try {
        // 1. Discover and update FX rates before price/strategy calculations.
        try {
            if (typeof syncCurrencyPairs === 'function') {
                const pairSyncResult = syncCurrencyPairs({ silent: true });
                if (pairSyncResult && pairSyncResult.fatal) {
                    throw new Error("Currency pair sync failed before strategy run: " + (pairSyncResult.message || pairSyncResult.status));
                }
                if (pairSyncResult && pairSyncResult.status && pairSyncResult.status !== 'COMPLETE') {
                    throw new Error("Currency pair sync did not complete: " + (pairSyncResult.message || pairSyncResult.status));
                }
            }

            if (typeof updateAllFxRates === 'function') {
                const fxUpdateResult = updateAllFxRates();
                if (fxUpdateResult && fxUpdateResult.fatal) {
                    throw new Error("FX rate update failed before strategy run: " + (fxUpdateResult.message || fxUpdateResult.status));
                }
                if (fxUpdateResult && fxUpdateResult.status && fxUpdateResult.status !== 'COMPLETE') {
                    console.warn("FX rate update completed with warning: " + (fxUpdateResult.message || fxUpdateResult.status));
                }
            }
        } catch (e) {
            throw new Error("FX update skipped/failed: " + e.message);
        }

        // 2. Update Market Prices (Crypto + Stock)
        try {
            if (typeof updateAllPrices === 'function') {
                const priceUpdateResult = updateAllPrices();
                if (priceUpdateResult && priceUpdateResult.fatal) {
                    throw new Error("Price update failed before strategy run: " + (priceUpdateResult.message || priceUpdateResult.status));
                }
                if (priceUpdateResult && priceUpdateResult.status && priceUpdateResult.status !== 'COMPLETE') {
                    console.warn("Price update completed with warning: " + (priceUpdateResult.message || priceUpdateResult.status));
                }
            }
        } catch (e) {
            throw new Error("Price update skipped/failed: " + e.message);
        }

        // [NEW] 3. Sync Asset Balances
        // Ensure balances are fresh BEFORE running strategy
        syncAllAssets_();

        // 4. Execute Strategic Monitor (The 30-min heart)
        // This reads indicators, checks logic, and updates dashboard.
        if (typeof runStrategicMonitor === 'function') {
            runStrategicMonitor();
        } else {
            throw new Error("runStrategicMonitor function not found in Core_StrategicEngine.");
        }

        console.log(`[${context}] Routine Completed.`);
        result.status = "COMPLETE";
        result.ok = true;
        result.fatal = false;
        result.message = "Routine Completed.";
        return result;
    } catch (e) {
        result.status = "FAILED";
        result.ok = false;
        result.fatal = true;
        result.message = e.toString();
        console.error(`[${context}] CRITICAL FAILURE: ${result.message}`);
        if (opts.throwOnFatal) throw e;
        return result;
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
        const masterResult = runAutomationMaster({ throwOnFatal: true });
        if (masterResult && masterResult.fatal) {
            throw new Error(`Automation master failed before daily close: ${masterResult.message || masterResult.status}`);
        }

        // Wait for spreadsheet calculation propagation
        Utilities.sleep(5000);

        // [NEW] Build Context for Snapshot
        // We rebuild it to ensure we capture the state exactly as it is after sync
        let snapshotContext = null;
        try {
            if (typeof buildFreshContext === 'function') {
                snapshotContext = buildFreshContext();
            } else if (typeof buildContext === 'function') {
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
        const email = Settings.get('ADMIN_EMAIL');
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

    ExchangeRegistry.getActive().forEach(function (entry) {
        const syncFn = globalThis[entry.functionName];
        try {
            if (typeof syncFn === 'function') {
                syncFn();
            } else {
                console.warn(`${entry.functionName} not found`);
            }
        } catch (e) {
            console.error(`${entry.moduleName} Sync Failed`, e);
        }
    });
}

