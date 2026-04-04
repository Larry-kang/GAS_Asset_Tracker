/**
 * Repository for the 參數設定 / Market Settings currency matrix.
 */
const SettingsMatrixRepo = {
    DEFAULT_SHEET_NAME: "參數設定",
    LEGACY_SHEET_NAMES: ["Market Settings"],
    TIMESTAMP_HEADER: "Timestamp",

    getSheet_: function (ss, options) {
        const opts = options || {};
        const primaryName = opts.sheetName || this.DEFAULT_SHEET_NAME;
        const primarySheet = ss.getSheetByName(primaryName);
        if (primarySheet) {
            WorkbookContracts.validateSheetAgainstContract(primarySheet, WorkbookContracts.getContract("SETTINGS_MATRIX"));
            return primarySheet;
        }

        if (opts.allowLegacyName) {
            for (let index = 0; index < this.LEGACY_SHEET_NAMES.length; index++) {
                const legacyName = this.LEGACY_SHEET_NAMES[index];
                const legacySheet = ss.getSheetByName(legacyName);
                if (!legacySheet) continue;
                WorkbookContracts.validateSheetAgainstContract(legacySheet, WorkbookContracts.getContract("SETTINGS_MATRIX"));
                return legacySheet;
            }
        }

        throw new Error(`Workbook contract failed: missing required sheet "${primaryName}".`);
    },

    readMatrix: function (ss, options) {
        const sheet = this.getSheet_(ss, options);
        const lastRow = Math.max(sheet.getLastRow(), 2);
        const lastCol = Math.max(sheet.getLastColumn(), 3);
        const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        const headerRow = values[0];
        const lastConfiguredIndex = this.findLastNonEmptyIndex_(headerRow);
        const targetColumns = {};
        const targetCurrencies = [];
        const sourceRows = {};
        const sourceCurrencies = [];

        for (let columnIndex = 1; columnIndex <= lastConfiguredIndex; columnIndex += 2) {
            const currency = this.normalizeValue_(headerRow[columnIndex]);
            if (!currency) continue;

            targetColumns[currency] = {
                valueColumn: columnIndex + 1,
                timestampColumn: columnIndex + 2
            };
            targetCurrencies.push(currency);
        }

        for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
            const currency = this.normalizeValue_(values[rowIndex][0]);
            if (!currency) continue;

            sourceRows[currency] = rowIndex + 1;
            sourceCurrencies.push(currency);
        }

        return {
            sheet: sheet,
            values: values,
            sourceRows: sourceRows,
            sourceCurrencies: sourceCurrencies,
            targetColumns: targetColumns,
            targetCurrencies: targetCurrencies,
            lastConfiguredColumn: Math.max(lastConfiguredIndex + 1, 1),
            lastRow: lastRow,
            lastCol: lastCol
        };
    },

    lookupRate: function (ss, fromCurrency, toCurrency, options) {
        const matrix = this.readMatrix(ss, options);
        const rowNumber = matrix.sourceRows[fromCurrency];
        const targetMeta = matrix.targetColumns[toCurrency];

        if (!rowNumber || !targetMeta) {
            return {
                sheet: matrix.sheet,
                hasPair: false,
                value: null
            };
        }

        return {
            sheet: matrix.sheet,
            hasPair: true,
            value: matrix.values[rowNumber - 1][targetMeta.valueColumn - 1],
            timestamp: matrix.values[rowNumber - 1][targetMeta.timestampColumn - 1]
        };
    },

    appendMissingPairs: function (ss, requiredPairs, options) {
        const sheet = this.getSheet_(ss, options);
        const normalizedPairs = this.normalizePairs_(requiredPairs);
        if (normalizedPairs.length === 0) {
            return this.emptyAppendResult_(sheet);
        }

        const matrix = this.readMatrix(ss, options);
        const missingFrom = [];
        const missingTo = [];

        normalizedPairs.forEach(pair => {
            if (!matrix.sourceRows[pair.from] && missingFrom.indexOf(pair.from) === -1) {
                missingFrom.push(pair.from);
                matrix.sourceRows[pair.from] = -1;
            }

            if (!matrix.targetColumns[pair.to] && missingTo.indexOf(pair.to) === -1) {
                missingTo.push(pair.to);
                matrix.targetColumns[pair.to] = { valueColumn: -1, timestampColumn: -1 };
            }
        });

        if (missingFrom.length > 0) {
            const startRow = sheet.getLastRow() + 1;
            const fromRows = missingFrom.map(currency => [currency]);
            sheet.getRange(startRow, 1, fromRows.length, 1).setValues(fromRows);
        }

        if (missingTo.length > 0) {
            const startColumn = matrix.lastConfiguredColumn + 1;
            const headerValues = [];
            missingTo.forEach(currency => {
                headerValues.push(currency, this.TIMESTAMP_HEADER);
            });
            sheet.getRange(1, startColumn, 1, headerValues.length).setValues([headerValues]);
        }

        return {
            sheet: sheet,
            newFromAdded: missingFrom.length > 0,
            newToAdded: missingTo.length > 0,
            addedFromCurrencies: missingFrom,
            addedToCurrencies: missingTo
        };
    },

    normalizePairs_: function (pairs) {
        return (pairs || [])
            .map(pair => ({
                from: this.normalizeValue_(pair && pair.from),
                to: this.normalizeValue_(pair && pair.to)
            }))
            .filter(pair => pair.from && pair.to);
    },

    emptyAppendResult_: function (sheet) {
        return {
            sheet: sheet,
            newFromAdded: false,
            newToAdded: false,
            addedFromCurrencies: [],
            addedToCurrencies: []
        };
    },

    findLastNonEmptyIndex_: function (row) {
        for (let index = row.length - 1; index >= 0; index--) {
            if (this.normalizeValue_(row[index])) return index;
        }
        return 0;
    },

    normalizeValue_: function (value) {
        return String(value == null ? "" : value).trim();
    }
};
