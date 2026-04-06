/**
 * Read-only repository for the Sync_Status operational view.
 */
const SyncStatusRepo = {
    DEFAULT_SHEET_NAME: "Sync_Status",
    HEADERS: ["Exchange", "Last Attempt", "Last Success", "Status", "Rows", "Message", "Updated"],

    getSheet_: function (ss, options) {
        const opts = options || {};
        const sheetName = opts.sheetName || this.DEFAULT_SHEET_NAME;
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Workbook contract failed: missing required sheet "${sheetName}".`);
        }
        WorkbookContracts.validateSheetAgainstContract(sheet, WorkbookContracts.getContract("SYNC_STATUS"));
        return sheet;
    },

    readAll: function (ss, options) {
        const sheet = this.getSheet_(ss, options);
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return [];

        return sheet.getRange(2, 1, lastRow - 1, this.HEADERS.length).getValues().map(row => ({
            exchange: this.normalizeText_(row[0]),
            lastAttempt: this.parseDate_(row[1]),
            lastSuccess: this.parseDate_(row[2]),
            status: this.normalizeText_(row[3]).toUpperCase(),
            rows: this.parseNumber_(row[4]),
            message: this.normalizeText_(row[5]),
            updatedAt: this.parseDate_(row[6])
        }));
    },

    readByExchange: function (ss, exchange, options) {
        const target = this.normalizeText_(exchange);
        return this.readAll(ss, options).filter(row => row.exchange === target)[0] || null;
    },

    normalizeText_: function (value) {
        return String(value == null ? "" : value).trim();
    },

    parseNumber_: function (value) {
        if (typeof value === "number") {
            return isFinite(value) ? value : 0;
        }
        const parsed = parseFloat(this.normalizeText_(value).replace(/,/g, ""));
        return isFinite(parsed) ? parsed : 0;
    },

    parseDate_: function (value) {
        if (!value) return null;
        if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
            return value;
        }
        const text = this.normalizeText_(value);
        if (!text) return null;
        const parsed = new Date(text);
        return isNaN(parsed.getTime()) ? null : parsed;
    }
};
