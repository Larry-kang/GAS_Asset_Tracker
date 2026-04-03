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
            SystemLogsRepo.prependLog(ss, {
                timestamp: new Date(),
                level: level,
                context: context,
                message: message
            }, {
                sheetName: this.SHEET_NAME
            });

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
            const numToDelete = SystemLogsRepo.cleanupOldLogs(ss, retentionDays, {
                sheetName: this.SHEET_NAME
            });
            if (numToDelete > 0) {
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

