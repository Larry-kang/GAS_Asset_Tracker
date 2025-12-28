// ==========================================
// --- Automation Driver (v24.5) ---
// ==========================================

/**
 * [Automation Master]
 * Main driver for the scheduler. Integrates price updates, balance sync, and strategic check.
 */
function runAutomationMaster() {
  const context = "MasterLoop";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('[Master] Starting automated synchronization...');
  LogService.info("Starting automated synchronization loop...", context);

  const startTime = new Date().getTime();

  try {
    // 1. Update Market Data
    LogService.info("Phase 1: Updating Market Data (Prices & Forex)...", context);
    if (typeof syncCurrencyPairs === 'function') syncCurrencyPairs(true);
    updateAllPrices();

    // 2. Sync Exchange Balances
    LogService.info("Phase 2: Synchronizing External Balances...", context);
    const syncTasks = [
      { name: "Binance", fn: typeof getBinanceBalance === 'function' ? getBinanceBalance : null },
      { name: "Okx", fn: typeof getOkxBalance === 'function' ? getOkxBalance : null },
      { name: "BitoPro", fn: typeof getBitoProBalance === 'function' ? getBitoProBalance : null },
      { name: "Pionex", fn: typeof getPionexBalance === 'function' ? getPionexBalance : null }
    ];

    syncTasks.forEach(task => {
      if (task.fn) {
        try {
          task.fn();
          LogService.info(`Sync Success: ${task.name}`, context);
        } catch (e) {
          LogService.error(`Sync Failed: ${task.name} - ${e.toString()}`, context);
        }
      }
    });

    // 3. SAP Strategic Check (Sends Report)
    LogService.info("Phase 3: Executing SAP Strategic Check...", context);
    runDailyInvestmentCheck();

    const duration = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    LogService.strategic(`Master Loop Completed Successfully. Duration: ${duration}s`, context);
    ss.toast(`[OK] Sync completed in ${duration}s.`);

  } catch (e) {
    LogService.error(`CRITICAL FAILURE: ${e.toString()}`, context);
    ss.toast('[ERROR] Sync failed: ' + e.message);

    if (email) MailApp.sendEmail(email, "[CRITICAL] SAP Sync Failed", e.toString());
  }
}

/**
 * [Daily Close Routine]
 * Aggregates all daily-end operations:
 * 1. Force Sync (Updates Prices & Balances)
 * 2. Force Flush (Recalculate Formulas)
 * 3. Take Snapshot (Record Daily Value)
 */
function runDailyCloseRoutine() {
  const context = "DailyClose";
  LogService.info("Starting Daily Close Routine...", context);

  try {
    // Step 1: Force Sync
    runAutomationMaster();

    // Step 2: Force Flush & Wait
    // Spreadsheets sometimes need time to propagate formula calculations after data entry
    SpreadsheetApp.flush();
    Utilities.sleep(5000); // Wait 5 seconds for formulas to update

    // Step 3: Take Snapshot
    LogService.info("Executing Daily Snapshot...", context);
    autoRecordDailyValues();

    LogService.info("Daily Close Routine Completed Successfully.", context);

  } catch (e) {
    LogService.error(`Daily Close Routine Failed: ${e.toString()}`, context);
    const email = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (email) MailApp.sendEmail(email, "【錯誤】每日結算流程失敗", e.toString());
  }
}