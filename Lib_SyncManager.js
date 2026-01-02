/**
 * Lib_SyncManager.js
 * 
 * [Core Shared Library] Asset Sync Manager
 * Purpose: Standardize 'Data Merge', 'Sheet Writing', 'Logging', and 'Error Handling'.
 * 
 * Key Features:
 * 1. updateUnifiedLedger: Unified Asset Ledger (Upsert/Replace by Exchange).
 * 2. log: Routes unified logs to System_Logs via LogService.
 * 3. run: Execution wrapper with error handling.
 */

const SyncManager = {

    // Unified Ledger Configuration
    LEDGER_SHEET_NAME: "Unified Assets",
    LEDGER_HEADERS: ["Exchange", "Currency", "Amount", "Type", "Status", "Meta", "Updated"],

    /**
     * [Core] Execution Wrapper
     * Wraps task execution with error handling and logging.
     * @param {string} moduleName - e.g. "Binance", "Okx"
     * @param {Function} taskFn - Main logic function
     */
    run: function (moduleName, taskFn) {
        try {
            this.log("INFO", `[${moduleName}] Starting Sync...`, moduleName);
            taskFn();
            this.log("INFO", `[${moduleName}] Sync Finished`, moduleName);
        } catch (e) {
            this.log("ERROR", `Execution Crashed: ${e.message}`, moduleName);
            try { SpreadsheetApp.getActiveSpreadsheet().toast(`${moduleName} Error: ${e.message}`); } catch (ignore) { }
            console.error(e);
        }
    },

    /**
     * [Core] Unified Ledger Updater (Atomic Replace by Exchange)
     * 
     * Strategy:
     * 1. Read ALL data from 'Unified Assets'.
     * 2. Filter OUT all rows belonging to 'targetExchange'.
     * 3. APPEND new asset rows from specific exchange.
     * 4. SORT and WRITE back.
     * 
     * @param {SpreadsheetApp.Spreadsheet} ss 
     * @param {string} targetExchange - e.g. "Binance", "Okx"
     * @param {Array<Object>} newAssets - List of standardized asset objects
     * Schema: { ccy, amt, type, status, meta }
     */
    updateUnifiedLedger: function (ss, targetExchange, newAssets) {
        let sheet = ss.getSheetByName(this.LEDGER_SHEET_NAME);

        // 1. Auto-Create Sheet if missing
        if (!sheet) {
            sheet = ss.insertSheet(this.LEDGER_SHEET_NAME);
            // Init Headers
            sheet.getRange(1, 1, 1, this.LEDGER_HEADERS.length).setValues([this.LEDGER_HEADERS]);
            sheet.getRange(1, 1, 1, this.LEDGER_HEADERS.length).setFontWeight('bold').setBackground('#e3f2fd');
            sheet.setFrozenRows(1);
        }

        // 2. Read Existing Data
        const lastRow = sheet.getLastRow();
        let existingData = [];
        if (lastRow > 1) {
            // Read all data (Cols A to G)
            existingData = sheet.getRange(2, 1, lastRow - 1, this.LEDGER_HEADERS.length).getValues();
        }

        // 3. Filter OUT old data for this exchange
        // Col Index 0 is 'Exchange'
        const otherExchangeData = existingData.filter(row => row[0] !== targetExchange);

        this.log("INFO", `[Ledger] Retained ${otherExchangeData.length} rows from other exchanges. Removing old ${targetExchange} data.`, "SyncManager");

        // 4. Transform New Assets to Row Format
        // Schema: Exchange | Currency | Amount | Type | Status | Meta | Updated
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

        const newRows = newAssets.map(asset => {
            return [
                targetExchange,           // Exchange
                asset.ccy,                // Currency
                asset.amt,                // Amount
                asset.type || 'Spot',     // Type
                asset.status || 'Avail',  // Status
                asset.meta || '',         // Meta
                timestamp                 // Updated
            ];
        });

        // 5. Combine & Sort
        const finalData = [...otherExchangeData, ...newRows];

        // Sort logic: Exchange (A-Z) -> Type (Spot/Earn/Loan) -> Currency (A-Z)
        finalData.sort((a, b) => {
            if (a[0] !== b[0]) return a[0].localeCompare(b[0]); // Exchange
            if (a[3] !== b[3]) return b[3].localeCompare(a[3]); // Type (Spot vs Loan desc)
            return a[1].localeCompare(b[1]);                    // Currency
        });

        // 6. Write Back (Atomic Replace)
        // Clear whole sheet content first (except header)
        if (lastRow > 1) {
            sheet.getRange(2, 1, lastRow - 1, this.LEDGER_HEADERS.length).clearContent();
        }

        if (finalData.length > 0) {
            sheet.getRange(2, 1, finalData.length, this.LEDGER_HEADERS.length).setValues(finalData);
            this.log("INFO", `[Ledger] Updated. Total Rows: ${finalData.length} (Added ${newRows.length} from ${targetExchange})`, "SyncManager");
        } else {
            this.log("WARNING", `[Ledger] Sheet is empty after update.`, "SyncManager");
        }
    },

    /**
     * [Core] Unified Logging Interface
     */
    log: function (level, message, context = "SyncManager") {
        console.log(`[${level}] [${context}] ${message}`);
        if (typeof LogService !== 'undefined' && LogService.log) {
            LogService.log(level, message, context);
        }
    }
};

