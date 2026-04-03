/**
 * Workbook contract helpers for the live Finance spreadsheet.
 * These validators protect the current workbook layout without changing it.
 */
const WorkbookContracts = {
    CONTRACTS: {
        UNIFIED_ASSETS: {
            key: "UNIFIED_ASSETS",
            sheetName: "Unified Assets",
            type: "headerRow",
            headers: ["Exchange", "Currency", "Amount", "Type", "Status", "Meta", "Updated"]
        },
        SYSTEM_LOGS: {
            key: "SYSTEM_LOGS",
            sheetName: "System_Logs",
            type: "headerRow",
            headers: ["Timestamp", "Level", "Module/Context", "Message"]
        },
        PRICE_CACHE: {
            key: "PRICE_CACHE",
            sheetName: "價格暫存",
            type: "headerRow",
            headers: ["標的", "類型", "價格", "更新時間"]
        },
        SETTINGS_MATRIX: {
            key: "SETTINGS_MATRIX",
            sheetName: "參數設定",
            type: "settingsMatrix"
        }
    },

    getContract: function (key) {
        const contract = this.CONTRACTS[key];
        if (!contract) {
            throw new Error(`Unknown workbook contract: ${key}`);
        }
        return contract;
    },

    requireContractSheet: function (ss, key) {
        const contract = this.getContract(key);
        const sheet = ss.getSheetByName(contract.sheetName);
        if (!sheet) {
            throw new Error(`Workbook contract failed: missing required sheet "${contract.sheetName}".`);
        }
        this.validateSheetAgainstContract(sheet, contract);
        return sheet;
    },

    validateCoreSheets: function (ss) {
        return Object.keys(this.CONTRACTS).map(key => {
            const contract = this.getContract(key);
            const sheet = ss.getSheetByName(contract.sheetName);
            if (!sheet) {
                return {
                    key: key,
                    sheetName: contract.sheetName,
                    ok: false,
                    message: `Missing sheet "${contract.sheetName}".`
                };
            }

            try {
                const details = this.validateSheetAgainstContract(sheet, contract);
                return {
                    key: key,
                    sheetName: contract.sheetName,
                    ok: true,
                    message: "OK",
                    details: details
                };
            } catch (e) {
                return {
                    key: key,
                    sheetName: contract.sheetName,
                    ok: false,
                    message: e.message
                };
            }
        });
    },

    validateSheetAgainstContract: function (sheet, contract) {
        if (contract.type === "headerRow") {
            return this.assertHeaderRow_(sheet, contract.headers, contract.sheetName);
        }

        if (contract.type === "settingsMatrix") {
            return this.validateSettingsMatrix_(sheet);
        }

        throw new Error(`Unsupported contract type: ${contract.type}`);
    },

    ensureHeaderSheet: function (ss, sheetName, headers, options) {
        const opts = options || {};
        let sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            this.initializeHeaderSheet_(sheet, headers, opts);
            return sheet;
        }

        const actualHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
        const isBlankHeader = actualHeaders.every(value => !this.normalizeValue_(value));
        if (isBlankHeader) {
            this.initializeHeaderSheet_(sheet, headers, opts);
            return sheet;
        }

        this.assertHeaderRow_(sheet, headers, sheetName);
        return sheet;
    },

    assertHeaderRow_: function (sheet, expectedHeaders, label) {
        const actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0]
            .map(this.normalizeValue_);
        const normalizedExpected = expectedHeaders.map(this.normalizeValue_);
        const mismatches = [];

        normalizedExpected.forEach((expected, index) => {
            if (actualHeaders[index] !== expected) {
                const columnLetter = this.columnToLetter_(index + 1);
                mismatches.push(`${columnLetter} expected "${expected}" got "${actualHeaders[index] || "(blank)"}"`);
            }
        });

        if (mismatches.length > 0) {
            throw new Error(`Workbook contract failed for "${label}": ${mismatches.join("; ")}`);
        }

        return {
            type: "headerRow",
            headers: normalizedExpected
        };
    },

    validateSettingsMatrix_: function (sheet) {
        const lastRow = Math.max(sheet.getLastRow(), 2);
        const lastCol = Math.max(sheet.getLastColumn(), 3);
        const firstRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(this.normalizeValue_);
        const errors = [];
        const toCurrencies = [];
        const lastNonEmptyCol = this.findLastNonEmptyIndex_(firstRow);

        if (!firstRow[1]) {
            errors.push('B1 must contain the first target currency code.');
        }

        for (let index = 1; index <= lastNonEmptyCol; index += 2) {
            const currencyHeader = firstRow[index];
            const timestampHeader = firstRow[index + 1] || "";
            const columnLetter = this.columnToLetter_(index + 1);

            if (!currencyHeader) {
                errors.push(`${columnLetter}1 must contain a target currency code.`);
                continue;
            }

            if (timestampHeader !== "Timestamp") {
                errors.push(`${this.columnToLetter_(index + 2)}1 must be "Timestamp" for ${currencyHeader}.`);
            }

            toCurrencies.push(currencyHeader);
        }

        const fromCurrencies = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
            .flat()
            .map(this.normalizeValue_)
            .filter(Boolean);

        if (fromCurrencies.length === 0) {
            errors.push('Column A must contain at least one source currency starting from row 2.');
        }

        if (errors.length > 0) {
            throw new Error(`Workbook contract failed for "${sheet.getName()}": ${errors.join(" ")}`);
        }

        return {
            type: "settingsMatrix",
            sourceCurrencyCount: fromCurrencies.length,
            targetCurrencyCount: toCurrencies.length,
            targetCurrencies: toCurrencies
        };
    },

    initializeHeaderSheet_: function (sheet, headers, options) {
        const opts = options || {};
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        if (opts.headerBackground) {
            sheet.getRange(1, 1, 1, headers.length).setBackground(opts.headerBackground);
        }
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
        sheet.setFrozenRows(opts.frozenRows || 1);
    },

    findLastNonEmptyIndex_: function (row) {
        for (let index = row.length - 1; index >= 0; index--) {
            if (row[index]) return index;
        }
        return 0;
    },

    normalizeValue_: function (value) {
        return String(value == null ? "" : value).trim();
    },

    columnToLetter_: function (column) {
        let result = "";
        let current = column;
        while (current > 0) {
            const remainder = (current - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            current = Math.floor((current - 1) / 26);
        }
        return result;
    }
};

function validateWorkbookContracts() {
    const results = WorkbookContracts.validateCoreSheets(SpreadsheetApp.getActiveSpreadsheet());
    const failed = results.filter(result => !result.ok);
    if (failed.length > 0) {
        throw new Error(`Workbook contract validation failed: ${failed.map(item => item.message).join(" | ")}`);
    }
    return results;
}
