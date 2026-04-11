/**
 * Repository for the 價格暫存 sheet.
 */
const PriceCacheRepo = {
    DEFAULT_SHEET_NAME: "價格暫存",
    HEADERS: ["標的", "類型", "價格", "更新時間"],

    getSheet_: function (ss, options) {
        const opts = options || {};
        if (opts.sheetName && opts.sheetName !== this.DEFAULT_SHEET_NAME) {
            const sheet = ss.getSheetByName(opts.sheetName);
            if (!sheet) {
                throw new Error(`Workbook contract failed: missing required sheet "${opts.sheetName}".`);
            }
            WorkbookContracts.assertHeaderRow_(sheet, this.HEADERS, opts.sheetName);
            return sheet;
        }
        return WorkbookContracts.requireContractSheet(ss, "PRICE_CACHE");
    },

    readRows: function (ss, options) {
        const sheet = this.getSheet_(ss, options);
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return [];

        return sheet.getRange(2, 1, lastRow - 1, this.HEADERS.length).getValues().map((row, index) => ({
            rowNumber: index + 2,
            ticker: row[0],
            assetType: row[1],
            price: row[2],
            updatedAt: row[3]
        }));
    },

    writeRows: function (ss, rows, options) {
        const sheet = this.getSheet_(ss, options);
        const dataRows = (rows || []).map(row => [
            row.ticker || "",
            row.assetType || "",
            row.price === undefined ? "" : row.price,
            row.updatedAt || ""
        ]);

        const existingRowCount = Math.max(sheet.getLastRow() - 1, 0);
        if (dataRows.length > 0) {
            sheet.getRange(2, 1, dataRows.length, this.HEADERS.length).setValues(dataRows);
        }

        if (existingRowCount > dataRows.length) {
            sheet.getRange(2 + dataRows.length, 1, existingRowCount - dataRows.length, this.HEADERS.length).clearContent();
        }

        return {
            sheet: sheet,
            rowCount: dataRows.length
        };
    },

    appendRows: function (ss, rows, options) {
        const sheet = this.getSheet_(ss, options);
        const dataRows = (rows || []).map(row => [
            row.ticker || "",
            row.assetType || "",
            row.price === undefined ? "" : row.price,
            row.updatedAt || ""
        ]);

        if (dataRows.length === 0) {
            return {
                sheet: sheet,
                rowCount: 0,
                rows: []
            };
        }

        const startRow = Math.max(sheet.getLastRow() + 1, 2);
        sheet.getRange(startRow, 1, dataRows.length, this.HEADERS.length).setValues(dataRows);

        const appendedRows = (rows || []).map((row, index) => ({
            rowNumber: startRow + index,
            ticker: row.ticker || "",
            assetType: row.assetType || "",
            price: row.price === undefined ? "" : row.price,
            updatedAt: row.updatedAt || ""
        }));

        return {
            sheet: sheet,
            rowCount: dataRows.length,
            rows: appendedRows
        };
    },

    writePriceUpdates: function (ss, updates, options) {
        const sheet = this.getSheet_(ss, options);
        const validUpdates = (updates || []).filter(update => update && update.rowNumber > 1);

        validUpdates.forEach(update => {
            sheet.getRange(update.rowNumber, 3, 1, 2).setValues([[
                update.price === undefined ? "" : update.price,
                update.updatedAt || ""
            ]]);
        });

        return {
            sheet: sheet,
            rowCount: validUpdates.length
        };
    }
};
