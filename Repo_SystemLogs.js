/**
 * Repository for the System_Logs sheet.
 */
const SystemLogsRepo = {
    DEFAULT_SHEET_NAME: "System_Logs",
    HEADERS: ["Timestamp", "Level", "Module/Context", "Message"],
    MAX_DATA_ROWS: 500,
    SOFT_DATA_ROW_LIMIT: 550,

    getSheet_: function (ss, options) {
        const opts = options || {};
        return WorkbookContracts.ensureHeaderSheet(
            ss,
            opts.sheetName || this.DEFAULT_SHEET_NAME,
            this.HEADERS,
            {
                headerBackground: "#f3f3f3",
                frozenRows: 1
            }
        );
    },

    getExistingSheet_: function (ss, sheetName) {
        return ss.getSheetByName(sheetName || this.DEFAULT_SHEET_NAME);
    },

    prependLog: function (ss, entry, options) {
        const opts = options || {};
        const sheet = this.getSheet_(ss, opts);
        sheet.insertRowBefore(2);
        sheet.getRange(2, 1, 1, this.HEADERS.length).setValues([[
            entry.timestamp || new Date(),
            entry.level,
            entry.context || "",
            entry.message
        ]]);

        const lastRow = sheet.getLastRow();
        if (lastRow > this.SOFT_DATA_ROW_LIMIT) {
            sheet.deleteRows(this.MAX_DATA_ROWS + 2, lastRow - (this.MAX_DATA_ROWS + 1));
        }

        return sheet;
    },

    cleanupOldLogs: function (ss, retentionDays, options) {
        const opts = options || {};
        const sheet = this.getExistingSheet_(ss, opts.sheetName);
        if (!sheet || sheet.getLastRow() <= 1) return 0;

        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

        const lastRow = sheet.getLastRow();
        const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        let deleteStartRow = -1;

        for (let index = timestamps.length - 1; index >= 0; index--) {
            const rowDate = new Date(timestamps[index][0]);
            if (rowDate < cutoffDate) {
                deleteStartRow = index + 2;
            } else {
                break;
            }
        }

        if (deleteStartRow === -1) return 0;

        const numToDelete = lastRow - deleteStartRow + 1;
        sheet.deleteRows(deleteStartRow, numToDelete);
        return numToDelete;
    }
};
