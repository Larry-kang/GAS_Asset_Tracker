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
        },
        BALANCE_SHEET_VIEW: {
            key: "BALANCE_SHEET_VIEW",
            sheetName: "Balance Sheet",
            type: "balanceSheetView",
            scanRows: 30,
            columns: {
                ticker: { aliases: ["股票代號/銀行", "股票代號 / 銀行", "股票代號", "Ticker", "Asset", "Symbol"], match: "first" },
                value: { aliases: ["現值", "現值(TWD)", "現值（TWD）", "現值 TWD", "市值", "市值(TWD)", "市值（TWD）", "Value", "Market Value", "Market_Value_TWD"], match: "first" },
                purpose: { aliases: ["用途", "Purpose", "Category", "類型", "資產類型", "分類", "策略分類", "配置類型"], match: "last" }
            }
        },
        KEY_MARKET_INDICATORS_VIEW: {
            key: "KEY_MARKET_INDICATORS_VIEW",
            sheetName: "Key Market Indicators",
            type: "keyValueSheet",
            scanRows: 5,
            keyColumn: { aliases: ["Indicator_Name", "指標名稱", "Indicator", "Key"], match: "first" },
            valueColumn: { aliases: ["Value", "值", "數值"], match: "first" },
            requiredKeys: [
                "Current_BTC_Price",
                "SAP_Base_ATH",
                "Total_Martingale_Spent",
                "MAX_MARTINGALE_BUDGET",
                "MONTHLY_DEBT_COST"
            ]
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
        return this.validateSheetsByKeys_(ss, ["UNIFIED_ASSETS", "SYSTEM_LOGS", "PRICE_CACHE", "SETTINGS_MATRIX"]);
    },

    validateStrategyInputSheets: function (ss) {
        return this.validateSheetsByKeys_(ss, ["BALANCE_SHEET_VIEW", "KEY_MARKET_INDICATORS_VIEW"]);
    },

    validateSheetsByKeys_: function (ss, keys) {
        return keys.map(key => {
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

        if (contract.type === "balanceSheetView") {
            return this.validateBalanceSheetView_(sheet, contract);
        }

        if (contract.type === "keyValueSheet") {
            return this.validateKeyValueSheet_(sheet, contract);
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

    validateBalanceSheetView_: function (sheet, contract) {
        const scanRows = Math.max(contract.scanRows || 10, 3);
        const lastRow = Math.max(sheet.getLastRow(), scanRows);
        const lastCol = Math.max(sheet.getLastColumn(), 4);
        const values = sheet.getRange(1, 1, Math.min(lastRow, scanRows), lastCol).getValues();
        const headerMatch = this.findHeaderRowByColumnRules_(values, contract.columns);

        if (!headerMatch) {
            throw new Error(`Workbook contract failed for "${sheet.getName()}": could not locate Balance Sheet header row with ticker/value/purpose columns.`);
        }

        const dataStartRow = headerMatch.headerRowIndex + 1;
        const dataRowCount = Math.max(sheet.getLastRow() - headerMatch.headerRowIndex, 0);
        if (dataRowCount <= 0) {
            throw new Error(`Workbook contract failed for "${sheet.getName()}": no portfolio rows found below header row ${headerMatch.headerRowIndex}.`);
        }

        return {
            type: "balanceSheetView",
            headerRowIndex: headerMatch.headerRowIndex,
            tickerColumn: headerMatch.columns.ticker,
            valueColumn: headerMatch.columns.value,
            purposeColumn: headerMatch.columns.purpose,
            dataStartRow: dataStartRow,
            dataRowCount: dataRowCount
        };
    },

    validateKeyValueSheet_: function (sheet, contract) {
        const scanRows = Math.max(contract.scanRows || 5, 1);
        const lastRow = Math.max(sheet.getLastRow(), 2);
        const lastCol = Math.max(sheet.getLastColumn(), 2);
        const values = sheet.getRange(1, 1, Math.min(lastRow, scanRows), lastCol).getValues();
        const headerMatch = this.findHeaderRowByColumnRules_(values, {
            key: contract.keyColumn,
            value: contract.valueColumn
        });

        if (!headerMatch) {
            throw new Error(`Workbook contract failed for "${sheet.getName()}": could not locate key/value header row.`);
        }

        const fullValues = sheet.getRange(headerMatch.headerRowIndex + 1, 1, Math.max(lastRow - headerMatch.headerRowIndex, 1), lastCol).getValues();
        const rowLookup = {};

        fullValues.forEach((row, index) => {
            const key = this.normalizeValue_(row[headerMatch.columns.key - 1]);
            if (!key || rowLookup[key]) return;
            rowLookup[key] = {
                rowIndex: headerMatch.headerRowIndex + 1 + index,
                rawValue: row[headerMatch.columns.value - 1]
            };
        });

        const missingKeys = [];
        const nonNumericKeys = [];
        (contract.requiredKeys || []).forEach(requiredKey => {
            const entry = rowLookup[requiredKey];
            if (!entry) {
                missingKeys.push(requiredKey);
                return;
            }

            if (!this.isNumericValue_(entry.rawValue)) {
                nonNumericKeys.push(requiredKey);
            }
        });

        if (missingKeys.length > 0) {
            throw new Error(`Workbook contract failed for "${sheet.getName()}": missing required key(s): ${missingKeys.join(", ")}`);
        }

        if (nonNumericKeys.length > 0) {
            throw new Error(`Workbook contract failed for "${sheet.getName()}": required key(s) not numeric: ${nonNumericKeys.join(", ")}`);
        }

        return {
            type: "keyValueSheet",
            headerRowIndex: headerMatch.headerRowIndex,
            keyColumn: headerMatch.columns.key,
            valueColumn: headerMatch.columns.value,
            requiredKeys: contract.requiredKeys || [],
            rowCount: fullValues.length
        };
    },

    findHeaderRowByColumnRules_: function (rows, columnRules) {
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex].map(this.normalizeValue_);
            const match = {};
            let allMatched = true;

            Object.keys(columnRules).forEach(key => {
                if (!allMatched) return;
                const rule = columnRules[key];
                const columnIndex = this.findMatchingColumnIndex_(row, rule.aliases || [], rule.match || "first");
                if (columnIndex < 0) {
                    allMatched = false;
                    return;
                }
                match[key] = columnIndex + 1;
            });

            if (allMatched) {
                return {
                    headerRowIndex: rowIndex + 1,
                    columns: match
                };
            }
        }

        return null;
    },

    findMatchingColumnIndex_: function (row, aliases, matchMode) {
        const normalizedAliases = (aliases || []).map(alias => this.normalizeHeaderToken_(alias)).filter(Boolean);
        const exactIndexes = [];
        const fuzzyIndexes = [];

        row.forEach((value, index) => {
            const normalizedValue = this.normalizeHeaderToken_(value);
            if (!normalizedValue) return;

            if (normalizedAliases.indexOf(normalizedValue) >= 0) {
                exactIndexes.push(index);
                return;
            }

            const hasFuzzyMatch = normalizedAliases.some(alias => {
                if (!alias || alias.length < 2) return false;
                return normalizedValue.indexOf(alias) >= 0 || alias.indexOf(normalizedValue) >= 0;
            });

            if (hasFuzzyMatch) {
                fuzzyIndexes.push(index);
            }
        });

        const indexes = exactIndexes.length > 0 ? exactIndexes : fuzzyIndexes;
        if (indexes.length === 0) return -1;
        return matchMode === "last" ? indexes[indexes.length - 1] : indexes[0];
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

    normalizeHeaderToken_: function (value) {
        return this.normalizeValue_(value)
            .toLowerCase()
            .replace(/[\s\u3000_\-:/()（）\[\]【】]/g, "");
    },

    isNumericValue_: function (value) {
        if (typeof value === "number") {
            return isFinite(value);
        }

        const normalized = this.normalizeValue_(value)
            .replace(/NT\$/gi, "")
            .replace(/\$/g, "")
            .replace(/,/g, "")
            .replace(/%/g, "");

        if (!normalized) return false;
        const num = parseFloat(normalized);
        return isFinite(num);
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

function validateStrategyInputContracts() {
    const results = WorkbookContracts.validateStrategyInputSheets(SpreadsheetApp.getActiveSpreadsheet());
    const failed = results.filter(result => !result.ok);
    if (failed.length > 0) {
        throw new Error(`Strategy input contract validation failed: ${failed.map(item => item.message).join(" | ")}`);
    }
    return results;
}
