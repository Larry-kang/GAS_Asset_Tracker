/**
 * Lib_SyncManager.js
 * 
 * [Core Shared Library] Asset Sync Manager
 * Purpose: Standardize 'Data Merge', 'Sheet Writing', 'Logging', and 'Error Handling'.
 * 
 * Key Features:
 * 1. updateUnifiedLedger: Unified Asset Ledger (Upsert/Replace by Exchange).
 * 2. log: Routes unified logs to System_Logs via LogService.
 * 3. run: Execution wrapper with error handling.
 */

const SyncManager = {

    // Unified Ledger Configuration
    LEDGER_SHEET_NAME: "Unified Assets",
    LEDGER_HEADERS: ["Exchange", "Currency", "Amount", "Type", "Status", "Meta", "Updated"],
    STATUS_SHEET_NAME: "Sync_Status",
    STATUS_HEADERS: ["Exchange", "Last Attempt", "Last Success", "Status", "Rows", "Message", "Updated"],

    /**
     * [Core] Execution Wrapper
     * Wraps task execution with error handling and logging.
     * @param {string} moduleName - e.g. "Binance", "Okx"
     * @param {Function} taskFn - Main logic function
     */
    run: function (moduleName, taskFn) {
        try {
            this.log("INFO", `[${moduleName}] Starting Sync...`, moduleName);
            const taskResult = taskFn();
            if (taskResult === false) {
                this.log("WARNING", `[${moduleName}] Sync Finished Without Commit`, moduleName);
                return false;
            }
            this.log("INFO", `[${moduleName}] Sync Finished`, moduleName);
            return taskResult;
        } catch (e) {
            this.log("ERROR", `Execution Crashed: ${e.message}`, moduleName);
            try { SpreadsheetApp.getActiveSpreadsheet().toast(`${moduleName} Error: ${e.message}`); } catch (ignore) { }
            console.error(e);
        }
    },

    /**
     * Builds a standard exchange sync result envelope.
     * @param {string} exchange
     * @returns {Object}
     */
    createResult: function (exchange) {
        return {
            exchange: exchange,
            status: "running",
            assets: [],
            requiredChecks: [],
            optionalChecks: [],
            errors: [],
            warnings: [],
            rowCount: 0,
            startedAt: new Date().toISOString(),
            finishedAt: ""
        };
    },

    /**
     * Registers the outcome of one data source fetch.
     * @param {Object} result
     * @param {Object} check
     */
    registerSourceCheck: function (result, check) {
        const normalized = {
            name: check.name || "Unknown",
            required: check.required !== false,
            success: !!check.success,
            rows: check.rows || 0,
            message: check.message || ""
        };

        const targetList = normalized.required ? result.requiredChecks : result.optionalChecks;
        targetList.push(normalized);

        if (!normalized.success) {
            const msg = `${normalized.name}: ${normalized.message || "Failed"}`;
            if (normalized.required) result.errors.push(msg);
            else result.warnings.push(msg);
        }
    },

    /**
     * Finalizes the result state before commit/status logging.
     * @param {Object} result
     * @returns {Object}
     */
    finalizeResult: function (result) {
        const requiredFailed = result.requiredChecks.some(check => !check.success);
        result.rowCount = result.assets.length;
        result.status = requiredFailed ? "failed" : "complete";
        result.finishedAt = new Date().toISOString();
        return result;
    },

    /**
     * Attempts to acquire a script lock for ledger commit.
     * Extracted into a helper so staging tests can override it.
     * @param {number} timeoutMs
     * @returns {GoogleAppsScript.Lock.Lock|null}
     */
    tryAcquireCommitLock_: function (timeoutMs = 5000) {
        const lock = LockService.getScriptLock();
        if (lock.tryLock(timeoutMs)) {
            return lock;
        }
        return null;
    },

    /**
     * Formats timestamps for Sync_Status comparison/storage.
     * @param {string|Date} value
     * @returns {string}
     */
    formatStatusTimestamp_: function (value) {
        const date = value ? new Date(value) : new Date();
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSS");
    },

    /**
     * Normalizes status timestamp values read from Sheets.
     * Sheet cells may come back as Date objects even when strings were written.
     * @param {string|Date|*} value
     * @returns {string}
     */
    normalizeStatusTimestampValue_: function (value) {
        if (!value) return "";

        if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
            return this.formatStatusTimestamp_(value);
        }

        const text = String(value).trim();
        if (!text) return "";
        if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}$/.test(text)) return text;

        const parsed = new Date(text);
        return isNaN(parsed.getTime()) ? text : this.formatStatusTimestamp_(parsed);
    },

    /**
     * Returns the current Sync_Status row for one exchange, if present.
     * @param {SpreadsheetApp.Spreadsheet} ss
     * @param {string} exchange
     * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet|null, rowIndex: number, row: Array|null}}
     */
    getSyncStatusRecord_: function (ss, exchange) {
        const sheet = ss.getSheetByName(this.STATUS_SHEET_NAME);
        if (!sheet) return { sheet: null, rowIndex: -1, row: null };

        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return { sheet: sheet, rowIndex: -1, row: null };

        const existingData = sheet.getRange(2, 1, lastRow - 1, this.STATUS_HEADERS.length).getValues();
        const rowIndex = existingData.findIndex(row => row[0] === exchange);
        return {
            sheet: sheet,
            rowIndex: rowIndex,
            row: rowIndex >= 0 ? existingData[rowIndex] : null
        };
    },

    /**
     * Determines whether the incoming attempt started before the currently recorded one.
     * @param {string} incomingAttemptAt
     * @param {string} recordedAttemptAt
     * @returns {boolean}
     */
    isStaleAttempt_: function (incomingAttemptAt, recordedAttemptAt) {
        return !!(recordedAttemptAt && incomingAttemptAt && incomingAttemptAt < recordedAttemptAt);
    },

    /**
     * Commit exchange assets only when the sync result is complete.
     * Failed syncs keep the previous ledger snapshot intact.
     * @param {SpreadsheetApp.Spreadsheet} ss
     * @param {string} moduleName
     * @param {Object} result
     * @returns {boolean}
     */
    commitExchangeResult: function (ss, moduleName, result) {
        this.finalizeResult(result);
        const attemptAt = this.formatStatusTimestamp_(result.startedAt);

        if (result.status !== "complete") {
            this.recordSyncStatus(ss, result, { lockTimeoutMs: 10000 });
            this.log("ERROR", `[${result.exchange}] Commit aborted. ${this.buildStatusMessage_(result)}`, moduleName);
            return false;
        }

        let lock = null;
        try {
            lock = this.tryAcquireCommitLock_();
            if (!lock) {
                result.status = "locked";
                result.errors.push("Unable to acquire script lock for ledger commit.");
                this.recordSyncStatus(ss, result, { lockTimeoutMs: 10000 });
                this.log("ERROR", `[${result.exchange}] Commit skipped: script lock unavailable.`, moduleName);
                return false;
            }

            const currentStatus = this.getSyncStatusRecord_(ss, result.exchange);
            const recordedAttemptAt = currentStatus.row
                ? this.normalizeStatusTimestampValue_(currentStatus.row[1])
                : "";
            if (this.isStaleAttempt_(attemptAt, recordedAttemptAt)) {
                result.status = "stale";
                result.warnings.push(`Stale attempt skipped. Incoming attempt ${attemptAt} older than recorded ${recordedAttemptAt}.`);
                this.log("WARNING", `[${result.exchange}] Commit skipped: stale attempt ${attemptAt} older than recorded ${recordedAttemptAt}.`, moduleName);
                return false;
            }

            this.updateUnifiedLedger(ss, result.exchange, result.assets);
            this.recordSyncStatus(ss, result, { lockAlreadyHeld: true });
            this.log("INFO", `[${result.exchange}] Commit completed. Rows: ${result.rowCount}`, moduleName);
            return true;
        } catch (e) {
            result.status = "failed";
            result.errors.push(`Ledger commit failed: ${e.message}`);
            this.recordSyncStatus(ss, result, { lockAlreadyHeld: !!lock, lockTimeoutMs: 10000 });
            this.log("ERROR", `[${result.exchange}] Ledger commit failed: ${e.message}`, moduleName);
            throw e;
        } finally {
            try { SpreadsheetApp.flush(); } catch (ignore) { }
            if (lock) {
                try { lock.releaseLock(); } catch (ignore) { }
            }
        }
    },

    /**
     * [Core] Unified Ledger Updater (Atomic Replace by Exchange)
     * 
     * Strategy:
     * 1. Read ALL data from 'Unified Assets'.
     * 2. Filter OUT all rows belonging to 'targetExchange'.
     * 3. APPEND new asset rows from specific exchange.
     * 4. SORT and WRITE back.
     * 
     * @param {SpreadsheetApp.Spreadsheet} ss 
     * @param {string} targetExchange - e.g. "Binance", "Okx"
     * @param {Array<Object>} newAssets - List of standardized asset objects
     * Schema: { ccy, amt, type, status, meta }
     */
    updateUnifiedLedger: function (ss, targetExchange, newAssets) {
        const summary = UnifiedAssetsRepo.replaceExchangeSnapshot(ss, targetExchange, newAssets, {
            sheetName: this.LEDGER_SHEET_NAME,
            headers: this.LEDGER_HEADERS
        });

        this.log("INFO", `[Ledger] Retained ${summary.retainedRows} rows from other exchanges. Removing old ${targetExchange} data.`, "SyncManager");

        if (summary.rowCount > 0) {
            this.log("INFO", `[Ledger] Updated. Total Rows: ${summary.rowCount} (Added ${summary.insertedRows} from ${targetExchange})`, "SyncManager");
        } else {
            this.log("WARNING", `[Ledger] Sheet is empty after update.`, "SyncManager");
        }
    },

    /**
     * Records the latest sync attempt status by exchange.
     * @param {SpreadsheetApp.Spreadsheet} ss
     * @param {Object} result
     * @param {Object=} options
     */
    recordSyncStatus: function (ss, result, options) {
        const opts = options || {};
        const lockAlreadyHeld = !!opts.lockAlreadyHeld;
        const lockTimeoutMs = opts.lockTimeoutMs || 5000;
        let lock = null;

        try {
            if (!lockAlreadyHeld) {
                lock = this.tryAcquireCommitLock_(lockTimeoutMs);
                if (!lock) {
                    this.log("WARNING", `[${result.exchange}] Sync_Status update skipped: unable to acquire script lock.`, "SyncManager");
                    return false;
                }
            }

            let sheet = WorkbookContracts.ensureHeaderSheet(
                ss,
                this.STATUS_SHEET_NAME,
                this.STATUS_HEADERS,
                {
                    headerBackground: '#fff3e0',
                    frozenRows: 1
                }
            );

            const lastRow = sheet.getLastRow();
            let existingData = [];
            if (lastRow > 1) {
                existingData = sheet.getRange(2, 1, lastRow - 1, this.STATUS_HEADERS.length).getValues();
            }

            const rowIndex = existingData.findIndex(row => row[0] === result.exchange);
            const attemptAt = this.formatStatusTimestamp_(result.startedAt || new Date());
            const updatedAt = this.formatStatusTimestamp_(result.finishedAt || new Date());
            const existingLastAttempt = rowIndex >= 0
                ? this.normalizeStatusTimestampValue_(existingData[rowIndex][1])
                : "";
            if (this.isStaleAttempt_(attemptAt, existingLastAttempt)) {
                this.log("WARNING", `[${result.exchange}] Sync_Status update skipped: stale attempt ${attemptAt} older than recorded ${existingLastAttempt}.`, "SyncManager");
                return false;
            }

            const previousLastSuccess = rowIndex >= 0
                ? this.normalizeStatusTimestampValue_(existingData[rowIndex][2])
                : "";
            const lastSuccess = result.status === "complete" ? updatedAt : previousLastSuccess;

            const rowData = [[
                result.exchange,
                attemptAt,
                lastSuccess || "",
                String(result.status || "unknown").toUpperCase(),
                result.rowCount || 0,
                this.buildStatusMessage_(result),
                updatedAt
            ]];

            if (rowIndex >= 0) {
                sheet.getRange(rowIndex + 2, 1, 1, this.STATUS_HEADERS.length).setValues(rowData);
            } else {
                sheet.getRange(lastRow + 1, 1, 1, this.STATUS_HEADERS.length).setValues(rowData);
            }
            SpreadsheetApp.flush();
            return true;
        } finally {
            if (lock) {
                try { lock.releaseLock(); } catch (ignore) { }
            }
        }
    },

    /**
     * Builds a concise message for Sync_Status and logs.
     * @param {Object} result
     * @returns {string}
     */
    buildStatusMessage_: function (result) {
        if (result.errors && result.errors.length > 0) {
            return result.errors.join(" | ").slice(0, 500);
        }
        if (result.warnings && result.warnings.length > 0) {
            return `Warnings: ${result.warnings.join(" | ").slice(0, 450)}`;
        }
        return `OK (${result.rowCount || 0} rows)`;
    },

    /**
     * [Core] Unified Logging Interface
     */
    log: function (level, message, context = "SyncManager") {
        console.log(`[${level}] [${context}] ${message}`);
        if (typeof LogService !== 'undefined' && LogService.log) {
            LogService.log(level, message, context);
        }
    }
};

