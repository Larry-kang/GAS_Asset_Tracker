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

            // 2. DESC Write: Always insert at Row 2 (under header)
            sheet.insertRowBefore(2);
            sheet.getRange(2, 1, 1, 4).setValues([[new Date(), level, context, message]]);

            // 3. Smart Cleanup: Keep sheet performant (Strict 500 limit managed by cleanup)
            // If rows > 550, delete old rows to bring it back to 500
            const lastRow = sheet.getLastRow();
            if (lastRow > 550) {
                sheet.deleteRows(502, lastRow - 501); // Row 1 is header, Row 501 is the 500th log
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
            // Now logs are DESC (Newest at top Row 2). Oldest are at the bottom.
            const lastRow = sheet.getLastRow();
            const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

            let deleteStartRow = -1;
            // Iterate from bottom to top to find the first log that IS within retention
            for (let i = timestamps.length - 1; i >= 0; i--) {
                const rowDate = new Date(timestamps[i][0]);
                if (rowDate < cutoffDate) {
                    // This row and everything below it should be deleted
                    deleteStartRow = i + 2; // +2 for header and 0-index
                } else {
                    // Found a log that is NOT expired
                    break;
                }
            }

            if (deleteStartRow !== -1) {
                const numToDelete = lastRow - deleteStartRow + 1;
                sheet.deleteRows(deleteStartRow, numToDelete);
                this.log("INFO", `Cleaned up ${numToDelete} old logs (Retention: ${retentionDays} days).`, "LogService");
            }

        } catch (e) {
            console.error("Log Cleanup Failed", e);
        }
    },

    info: function (msg, ctx) { this.log("INFO", msg, ctx); },

    warn: function (msg, ctx) {
        this.log("WARNING", msg, ctx);
        // [Discord] Send Warning
        if (typeof Discord !== 'undefined') {
            Discord.sendAlert("Warning: " + ctx, msg, "WARNING");
        }
    },

    error: function (msg, ctx) {
        this.log("ERROR", msg, ctx);
        // [Discord] Send Error (Critical)
        if (typeof Discord !== 'undefined') {
            Discord.sendAlert("Error: " + ctx, msg, "ERROR");
        }
    },

    strategic: function (msg, ctx) {
        this.log("STRATEGIC", msg, ctx);
        // [Discord] Send Strategic Signal (High Priority)
        if (typeof Discord !== 'undefined') {
            Discord.sendAlert("Strategic Signal: " + ctx, msg, "STRATEGIC");
        }
    }
};

