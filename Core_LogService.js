/**
 * SAP Core Log Service
 * Handles unified logging and audit trails.
 */
const LogService = {
    SHEET_NAME: "System_Logs",

    /**
     * Log an event to the System_Logs sheet.
     * @param {string} level - INFO, WARNING, ERROR, STRATEGIC
     * @param {string} message - Description of the event
     * @param {string} context - Source module or additional data
     */
    log: function (level, message, context = "") {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let sheet = ss.getSheetByName(this.SHEET_NAME);

            if (!sheet) {
                sheet = ss.insertSheet(this.SHEET_NAME);
                sheet.appendRow(["Timestamp", "Level", "Module/Context", "Message"]);
                sheet.setFrozenRows(1);
                sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");
            }

            sheet.appendRow([new Date(), level, context, message]);

            // Soft Limit: Keep sheet performant (< 2000 rows)
            // Actual retention is handled by cleanupOldLogs()
            if (sheet.getLastRow() > 2000) {
                sheet.deleteRows(2, 500); // Batch delete old rows
            }

            Logger.log(`[${level}] ${context}: ${message}`);
        } catch (e) {
            Logger.log("Critical failure in LogService: " + e.toString());
        }
    },

    /**
     * Deletes logs older than N days.
     * Recommended to run via Daily Trigger.
     * @param {number} retentionDays - Default 7
     */
    cleanupOldLogs: function (retentionDays = 7) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ss.getSheetByName(this.SHEET_NAME);
            if (!sheet || sheet.getLastRow() <= 1) return;

            const cutoffDate = new Date();
            cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

            // Read Timestamps (Col A)
            // Assume Row 1 is header. Data starts Row 2.
            const lastRow = sheet.getLastRow();
            const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

            let deleteCount = 0;
            for (let i = 0; i < timestamps.length; i++) {
                const rowDate = new Date(timestamps[i][0]);
                if (rowDate < cutoffDate) {
                    deleteCount++;
                } else {
                    // Logs are chronological (usually). 
                    // Once we hit a new log, we can stop checking if we assume order.
                    // But to be safe, if we assume unsorted, we can't break. 
                    // However, logs ARE chronological.
                    break;
                }
            }

            if (deleteCount > 0) {
                sheet.deleteRows(2, deleteCount);
                this.log("INFO", `Cleaned up ${deleteCount} old logs (Retention: ${retentionDays} days).`, "LogService");
            }

        } catch (e) {
            console.error("Log Cleanup Failed", e);
        }
    },

    info: function (msg, ctx) { this.log("INFO", msg, ctx); },
    warn: function (msg, ctx) { this.log("WARNING", msg, ctx); },
    error: function (msg, ctx) { this.log("ERROR", msg, ctx); },
    strategic: function (msg, ctx) { this.log("STRATEGIC", msg, ctx); }
};
