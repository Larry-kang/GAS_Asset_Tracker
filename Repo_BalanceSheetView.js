/**
 * Read-only repository for the Balance Sheet strategy input view.
 */
const BalanceSheetViewRepo = {
    DEFAULT_SHEET_NAME: "Balance Sheet",

    getView_: function (ss, options) {
        const opts = options || {};
        const sheetName = opts.sheetName || this.DEFAULT_SHEET_NAME;
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Workbook contract failed: missing required sheet "${sheetName}".`);
        }

        const details = WorkbookContracts.validateSheetAgainstContract(
            sheet,
            WorkbookContracts.getContract("BALANCE_SHEET_VIEW")
        );

        return {
            sheet: sheet,
            details: details
        };
    },

    readPortfolio: function (ss, options) {
        const view = this.getView_(ss, options);
        const values = this.getSheetValues_(view.sheet);
        const startIndex = view.details.dataStartRow - 1;
        const tickerIndex = view.details.tickerColumn - 1;
        const valueIndex = view.details.valueColumn - 1;
        const purposeIndex = view.details.purposeColumn - 1;
        const portfolio = [];

        for (let index = startIndex; index < values.length; index++) {
            const row = values[index];
            const ticker = this.normalizeText_(row[tickerIndex]);
            if (!ticker) continue;

            portfolio.push({
                ticker: ticker,
                value: this.parseNumeric_(row[valueIndex]),
                purpose: this.normalizeText_(row[purposeIndex])
            });
        }

        return portfolio;
    },

    getSheetValues_: function (sheet) {
        if (typeof DataCache !== "undefined" && DataCache && typeof DataCache.getValues === "function") {
            const cached = DataCache.getValues(sheet.getName());
            if (cached) return cached;
        }
        return sheet.getDataRange().getValues();
    },

    normalizeText_: function (value) {
        return String(value == null ? "" : value).trim();
    },

    parseNumeric_: function (value) {
        if (typeof value === "number") {
            return isFinite(value) ? value : 0;
        }

        const normalized = this.normalizeText_(value)
            .replace(/NT\$/gi, "")
            .replace(/\$/g, "")
            .replace(/,/g, "")
            .replace(/%/g, "");

        if (!normalized) return 0;

        const parsed = parseFloat(normalized);
        return isFinite(parsed) ? parsed : 0;
    }
};
