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

            // Keep only last 1000 logs to prevent sheet bloat
            if (sheet.getLastRow() > 1005) {
                sheet.deleteRows(2, 100);
            }

            Logger.log(`[${level}] ${context}: ${message}`);
        } catch (e) {
            Logger.log("Critical failure in LogService: " + e.toString());
        }
    },

    info: function (msg, ctx) { this.log("INFO", msg, ctx); },
    warn: function (msg, ctx) { this.log("WARNING", msg, ctx); },
    error: function (msg, ctx) { this.log("ERROR", msg, ctx); },
    strategic: function (msg, ctx) { this.log("STRATEGIC", msg, ctx); }
};
