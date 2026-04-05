/**
 * Read-only repository for the Key Market Indicators strategy input view.
 */
const KeyMarketIndicatorsViewRepo = {
    DEFAULT_SHEET_NAME: "Key Market Indicators",
    KEYS_OF_INTEREST: [
        "SAP_Base_ATH",
        "Total_Martingale_Spent",
        "Current_BTC_Price",
        "MAX_MARTINGALE_BUDGET",
        "MONTHLY_DEBT_COST",
        "BTC_MM",
        "Alloc_L1_Target",
        "Alloc_L2_Target",
        "Alloc_L3_Target",
        "Alloc_L4_Target",
        "00713_MM",
        "00662_MM",
        "00713_Price",
        "00662_Price",
        "00713_200DMA_Price",
        "00662_200DMA_Price"
    ],

    getView_: function (ss, options) {
        const opts = options || {};
        const sheetName = opts.sheetName || this.DEFAULT_SHEET_NAME;
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Workbook contract failed: missing required sheet "${sheetName}".`);
        }

        const details = WorkbookContracts.validateSheetAgainstContract(
            sheet,
            WorkbookContracts.getContract("KEY_MARKET_INDICATORS_VIEW")
        );

        return {
            sheet: sheet,
            details: details
        };
    },

    readIndicators: function (ss, options) {
        const view = this.getView_(ss, options);
        const values = this.getSheetValues_(view.sheet);
        const startIndex = view.details.headerRowIndex;
        const keyIndex = view.details.keyColumn - 1;
        const valueIndex = view.details.valueColumn - 1;
        const result = {};

        for (let index = startIndex; index < values.length; index++) {
            const row = values[index];
            const key = this.normalizeText_(row[keyIndex]);
            if (!key || !this.shouldIncludeKey_(key)) continue;
            result[key] = this.parseNumeric_(row[valueIndex]);
        }

        return result;
    },

    shouldIncludeKey_: function (key) {
        return this.KEYS_OF_INTEREST.indexOf(key) >= 0 ||
            key.indexOf("SAP_") > -1 ||
            key.indexOf("_Maint_Alert") > -1 ||
            key.indexOf("_Maint_Critical") > -1;
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
