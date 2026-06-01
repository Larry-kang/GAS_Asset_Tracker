/**
 * Optional repository for sheet-driven inventory export contracts.
 * Future fields should be added in the export sheets, not in Apps Script code.
 */
const InventoryExportRepo = {
    SCHEMA_VERSION: "2",

    readExport: function (ss, options) {
        const opts = options || {};
        const summarySheetName = opts.summarySheetName || Config.SHEET_NAMES.INVENTORY_EXPORT_SUMMARY;
        const positionsSheetName = opts.positionsSheetName || Config.SHEET_NAMES.INVENTORY_EXPORT_POSITIONS;

        const summaryValues = this.getOptionalSheetValues_(ss, summarySheetName);
        const positionsValues = this.getOptionalSheetValues_(ss, positionsSheetName);

        return this.buildExportBundleFromTables_(summaryValues, positionsValues, {
            summarySheetName: summarySheetName,
            positionsSheetName: positionsSheetName
        });
    },

    buildExportBundleFromTables_: function (summaryValues, positionsValues, metadata) {
        const meta = metadata || {};
        const summaryRows = this.parseSummaryRows_(summaryValues);
        const positions = this.parsePositionRows_(positionsValues);
        const available = summaryRows.length > 0 || positions.length > 0;
        const summary = {};

        summaryRows.forEach(row => {
            const key = this.normalizeText_(row.key);
            if (!key) return;
            summary[key] = row.value;
        });

        return {
            available: available,
            schemaVersion: this.SCHEMA_VERSION,
            sheets: {
                summary: meta.summarySheetName || null,
                positions: meta.positionsSheetName || null
            },
            summary: summary,
            summaryRows: summaryRows,
            positions: positions
        };
    },

    parseSummaryRows_: function (values) {
        const rows = this.parseTable_(values);
        if (!rows) return [];

        this.assertRequiredHeaders_(rows.headers, ["key", "value"], "summary export");
        return rows.records.filter(row => this.normalizeText_(row.key));
    },

    parsePositionRows_: function (values) {
        const rows = this.parseTable_(values);
        if (!rows) return [];

        this.assertRequiredHeaders_(rows.headers, ["ticker"], "position export");
        return rows.records.filter(row => this.normalizeText_(row.ticker));
    },

    parseTable_: function (values) {
        if (!values || values.length === 0) return null;

        const rawHeaders = values[0].map(value => this.normalizeText_(value));
        const canonicalHeaders = rawHeaders.map(value => this.toCanonicalHeader_(value));
        const records = [];

        for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
            const row = values[rowIndex];
            if (this.isBlankRow_(row)) continue;

            const record = {};
            for (let colIndex = 0; colIndex < canonicalHeaders.length; colIndex++) {
                const header = canonicalHeaders[colIndex];
                if (!header) continue;
                record[header] = this.coerceValue_(row[colIndex]);
            }
            records.push(record);
        }

        return {
            headers: canonicalHeaders,
            records: records
        };
    },

    assertRequiredHeaders_: function (headers, requiredHeaders, label) {
        const missing = (requiredHeaders || []).filter(header => headers.indexOf(header) < 0);
        if (missing.length > 0) {
            throw new Error(`Inventory export ${label} missing required header(s): ${missing.join(", ")}`);
        }
    },

    getOptionalSheetValues_: function (ss, sheetName) {
        if (!sheetName) return null;
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return null;
        return sheet.getDataRange().getValues();
    },

    toCanonicalHeader_: function (value) {
        const normalized = this.normalizeText_(value)
            .replace(/\s+/g, " ")
            .trim();
        const lower = normalized.toLowerCase();

        if (lower === "key" || lower === "value" || lower === "ticker") {
            return lower;
        }

        return normalized;
    },

    normalizeText_: function (value) {
        return String(value == null ? "" : value).trim();
    },

    coerceValue_: function (value) {
        if (typeof value === "number") {
            return isFinite(value) ? value : null;
        }

        if (typeof value === "boolean") return value;
        if (value == null) return "";

        const normalized = this.normalizeText_(value);
        if (!normalized) return "";

        if (/^(true|false)$/i.test(normalized)) {
            return normalized.toLowerCase() === "true";
        }

        const numeric = normalized.replace(/,/g, "");
        if (/^-?\d+(\.\d+)?$/.test(numeric)) {
            const parsed = parseFloat(numeric);
            return isFinite(parsed) ? parsed : normalized;
        }

        return normalized;
    },

    isBlankRow_: function (row) {
        if (!Array.isArray(row)) return true;
        return row.every(value => this.normalizeText_(value) === "");
    }
};

function getInventoryExportBundle_() {
    return InventoryExportRepo.readExport(SpreadsheetApp.getActiveSpreadsheet());
}
