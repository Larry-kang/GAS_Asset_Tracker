/**
 * Repository for the normalized Unified Assets ledger sheet.
 */
const UnifiedAssetsRepo = {
    DEFAULT_SHEET_NAME: "Unified Assets",
    HEADERS: ["Exchange", "Currency", "Amount", "Type", "Status", "Meta", "Updated"],

    getSheet_: function (ss, options) {
        const opts = options || {};
        return WorkbookContracts.ensureHeaderSheet(
            ss,
            opts.sheetName || this.DEFAULT_SHEET_NAME,
            opts.headers || this.HEADERS,
            {
                headerBackground: "#e3f2fd",
                frozenRows: 1
            }
        );
    },

    readAllRows: function (ss, options) {
        const opts = options || {};
        const headers = opts.headers || this.HEADERS;
        const sheet = this.getSheet_(ss, opts);
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return [];
        return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    },

    replaceExchangeSnapshot: function (ss, targetExchange, newAssets, options) {
        const opts = options || {};
        const headers = opts.headers || this.HEADERS;
        const sheet = this.getSheet_(ss, opts);
        const lastRow = sheet.getLastRow();
        const existingData = lastRow > 1
            ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
            : [];
        const otherExchangeData = existingData.filter(row => row[0] !== targetExchange);
        const timestamp = Utilities.formatDate(
            opts.timestamp ? new Date(opts.timestamp) : new Date(),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd'T'HH:mm:ss"
        );
        const newRows = newAssets.map(asset => [
            targetExchange,
            asset.ccy,
            asset.amt,
            asset.type || "Spot",
            asset.status || "Avail",
            asset.meta || "",
            timestamp
        ]);
        const finalData = otherExchangeData.concat(newRows);
        const existingRowCount = existingData.length;

        finalData.sort(function (a, b) {
            if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
            if (a[3] !== b[3]) return b[3].localeCompare(a[3]);
            return a[1].localeCompare(b[1]);
        });

        if (finalData.length > 0) {
            sheet.getRange(2, 1, finalData.length, headers.length).setValues(finalData);
            SpreadsheetApp.flush();
        }

        if (existingRowCount > finalData.length) {
            sheet.getRange(2 + finalData.length, 1, existingRowCount - finalData.length, headers.length).clearContent();
        }

        return {
            sheet: sheet,
            rowCount: finalData.length,
            insertedRows: newRows.length,
            retainedRows: otherExchangeData.length
        };
    }
};
