/**
 * Lib_SyncManager.js
 * 
 * [Core Shared Library] Asset Sync Manager
 * Purpose: Standardize 'Data Merge', 'Sheet Writing', 'Logging', and 'Error Handling'.
 * 
 * Key Features:
 * 1. writeToSheet: Handles Headers, ISO Time, Sorting, Clearing, and Empty State.
 * 2. mergeAssets: Robustly merges multiple Maps or Objects into one Map.
 * 3. log: Routes unified logs to System_Logs via LogService.
 */

const SyncManager = {

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
            // Toast for immediate user feedback
            try { SpreadsheetApp.getActiveSpreadsheet().toast(`${moduleName} Error: ${e.message}`); } catch (ignore) { }
            console.error(e);
        }
    },

    /**
     * [Core] Universal Asset Merger
     * Merges multiple Maps or Objects into a single Map.
     * @param {...(Map|Object)} sources - Data sources
     * @returns {Map<string, number>} Merged Map (Key: Currency, Value: Amount)
     */
    mergeAssets: function (...sources) {
        const combined = new Map();
        let totalSources = 0;
        let totalItems = 0;

        sources.forEach((source, index) => {
            if (!source) return;
            totalSources++;
            let count = 0;

            // Case A: Map
            if (source instanceof Map) {
                source.forEach((val, key) => {
                    const current = combined.get(key) || 0;
                    combined.set(key, current + val);
                    count++;
                });
            }
            // Case B: Object
            else if (typeof source === 'object') {
                Object.keys(source).forEach(key => {
                    const val = source[key];
                    const current = combined.get(key) || 0;
                    combined.set(key, current + val);
                    count++;
                });
            }
            else {
                console.warn(`[SyncManager] Merge source #${index} is neither Map nor Object: ${typeof source}`);
            }
            totalItems += count;
        });

        console.log(`[SyncManager] Merged ${totalSources} sources, total ${totalItems} items into ${combined.size} unique assets.`);
        return combined;
    },

    /**
     * [Core] Standard Sheet Writer
     * @param {SpreadsheetApp.Spreadsheet} ss 
     * @param {string} sheetName - Target Sheet Name
     * @param {string[]} headers - e.g. ['Currency', 'Amount', 'Last Updated']
     * @param {Map<string, number>} dataMap - Asset Data
     */
    writeToSheet: function (ss, sheetName, headers, dataMap) {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) sheet = ss.insertSheet(sheetName);

        // 1. Enforce Headers
        if (headers && headers.length > 0) {
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
        }

        // 2. Prepare Data Rows
        const rows = [];
        dataMap.forEach((val, key) => {
            // Filter dust
            if (val > 0.00000001) {
                rows.push([key, val]);
            }
        });

        // 3. Sort (Descending Amount)
        rows.sort((a, b) => b[1] - a[1]);

        // 4. ISO Timestamp
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

        // 5. Clear Old Data (From Row 2)
        const maxRows = sheet.getMaxRows();
        const maxCols = sheet.getMaxColumns();
        if (maxRows > 1) {
            sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
        }

        // Write New Data
        if (rows.length > 0) {
            // Write Asset Data (Col 1, 2)
            sheet.getRange(2, 1, rows.length, 2).setValues(rows);

            // Write Timestamp (At the end of Row 2)
            if (headers) {
                const timeCol = headers.length;
                sheet.getRange(2, timeCol).setValue(timestamp);
            }

            this.log("INFO", `Written ${rows.length} assets to ${sheetName}`, "SyncManager");
            ss.toast(`${sheetName} Update Success (${rows.length} items)`);
        } else {
            // Empty State
            if (headers) {
                const emptyRow = new Array(headers.length).fill("");
                emptyRow[0] = "No Assets Found (Check Logs)";
                emptyRow[1] = 0;
                emptyRow[headers.length - 1] = timestamp;

                sheet.getRange(2, 1, 1, headers.length).setValues([emptyRow]);
            }
            this.log("WARNING", `${sheetName} No Assets Found`, "SyncManager");
            ss.toast(`${sheetName} No Assets`);
        }
    },

    /**
     * [Core] Unified Logging Interface
     * Routes logs to both Console (Dev) and System_Logs (User Audit)
     */
    log: function (level, message, context = "SyncManager") {
        // 1. Console
        console.log(`[${level}] [${context}] ${message}`);

        // 2. System_Logs (if LogService exists)
        if (typeof LogService !== 'undefined' && LogService.log) {
            LogService.log(level, message, context);
        }
    }
};
