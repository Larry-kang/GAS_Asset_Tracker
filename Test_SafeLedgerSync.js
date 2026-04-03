/**
 * Guarded staging verification helpers for Safe Unified Ledger Sync refactor.
 *
 * Safety rules:
 * 1. Tests only run when staging Script Properties explicitly enable them.
 * 2. Tests only run against the approved staging spreadsheet ID.
 * 3. Tests only read/write `_TEST_*` sheets and never touch production tabs.
 */

const SafeLedgerTestHarness = {
  ENABLED_PROPERTY: 'SAFE_LEDGER_TEST_ENABLED',
  SPREADSHEET_ID_PROPERTY: 'SAFE_LEDGER_TEST_SPREADSHEET_ID',
  SHEET_PREFIX_PROPERTY: 'SAFE_LEDGER_TEST_PREFIX',
  DEFAULT_SHEET_PREFIX: '_TEST_',

  getConfig_: function () {
    const props = PropertiesService.getScriptProperties();
    const prefix = props.getProperty(this.SHEET_PREFIX_PROPERTY) || this.DEFAULT_SHEET_PREFIX;

    return {
      enabled: String(props.getProperty(this.ENABLED_PROPERTY) || '').toLowerCase() === 'true',
      spreadsheetId: props.getProperty(this.SPREADSHEET_ID_PROPERTY) || '',
      prefix: prefix
    };
  },

  assertReady_: function () {
    const config = this.getConfig_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!config.enabled) {
      throw new Error(
        `SafeLedgerTest blocked: set Script Property ${this.ENABLED_PROPERTY}=true in staging first.`
      );
    }

    if (!config.spreadsheetId) {
      throw new Error(
        `SafeLedgerTest blocked: set Script Property ${this.SPREADSHEET_ID_PROPERTY} to the staging spreadsheet ID.`
      );
    }

    if (ss.getId() !== config.spreadsheetId) {
      throw new Error(
        `SafeLedgerTest blocked: active spreadsheet ${ss.getId()} does not match approved staging spreadsheet ${config.spreadsheetId}.`
      );
    }

    return {
      ss: ss,
      prefix: config.prefix,
      sheetNames: {
        ledger: `${config.prefix}Unified Assets`,
        status: `${config.prefix}Sync_Status`,
        logs: `${config.prefix}System_Logs`
      }
    };
  },

  withSandbox_: function (callback) {
    const context = this.assertReady_();
    const original = {
      ledgerSheet: SyncManager.LEDGER_SHEET_NAME,
      statusSheet: SyncManager.STATUS_SHEET_NAME,
      logSheet: (typeof LogService !== 'undefined' && LogService) ? LogService.SHEET_NAME : null
    };

    SyncManager.LEDGER_SHEET_NAME = context.sheetNames.ledger;
    SyncManager.STATUS_SHEET_NAME = context.sheetNames.status;
    if (typeof LogService !== 'undefined' && LogService) {
      LogService.SHEET_NAME = context.sheetNames.logs;
    }

    try {
      return callback(context);
    } finally {
      SyncManager.LEDGER_SHEET_NAME = original.ledgerSheet;
      SyncManager.STATUS_SHEET_NAME = original.statusSheet;
      if (typeof LogService !== 'undefined' && LogService && original.logSheet) {
        LogService.SHEET_NAME = original.logSheet;
      }
    }
  },

  resolveSheetName_: function (logicalSheetName) {
    const context = this.assertReady_();
    const map = {
      'Unified Assets': context.sheetNames.ledger,
      'Sync_Status': context.sheetNames.status,
      'System_Logs': context.sheetNames.logs
    };
    return map[logicalSheetName] || logicalSheetName;
  },

  getSheetValues_: function (sheetName) {
    const context = this.assertReady_();
    const sheet = context.ss.getSheetByName(sheetName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return [];

    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  },

  captureProductionState_: function () {
    return {
      ledger: this.getSheetValues_('Unified Assets'),
      status: this.getSheetValues_('Sync_Status'),
      logs: this.getSheetValues_('System_Logs')
    };
  },

  assertProductionStateUnchanged_: function (beforeState) {
    const afterState = this.captureProductionState_();
    const labels = {
      ledger: 'Unified Assets',
      status: 'Sync_Status',
      logs: 'System_Logs'
    };

    Object.keys(labels).forEach(key => {
      assertSafeLedgerTest_(
        JSON.stringify(afterState[key]) === JSON.stringify(beforeState[key]),
        `Expected production sheet ${labels[key]} to remain unchanged during guarded test run.`
      );
    });

    return afterState;
  }
};

function setupSafeLedgerSyncTestEnvironment() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    const ss = context.ss;
    const sheetNames = [context.sheetNames.ledger, context.sheetNames.status, context.sheetNames.logs];

    sheetNames.forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) sheet = ss.insertSheet(name);
      sheet.clearContents();
    });

    ss.getSheetByName(context.sheetNames.ledger)
      .getRange(1, 1, 1, SyncManager.LEDGER_HEADERS.length)
      .setValues([SyncManager.LEDGER_HEADERS]);

    ss.getSheetByName(context.sheetNames.status)
      .getRange(1, 1, 1, SyncManager.STATUS_HEADERS.length)
      .setValues([SyncManager.STATUS_HEADERS]);

    ss.getSheetByName(context.sheetNames.logs)
      .getRange(1, 1, 1, 4)
      .setValues([['Timestamp', 'Level', 'Module/Context', 'Message']]);

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Safe ledger sync staging sheets initialized.',
      sheets: context.sheetNames
    };
  });
}

function runSafeLedgerSyncTestSuite() {
  const productionStateBefore = SafeLedgerTestHarness.captureProductionState_();
  const results = [];
  const testFns = [
    'testSyncManagerCommitSuccess',
    'testSyncManagerCommitRequiredFailure',
    'testSyncManagerOptionalFailure',
    'testSyncManagerLockGuard',
    'testSyncStatusWrite',
    'testStaleAttemptGuard',
    'testStandaloneStatusWriteUsesLock'
  ];

  testFns.forEach(name => {
    try {
      results.push({
        name: name,
        status: 'passed',
        result: this[name]()
      });
    } catch (e) {
      results.push({
        name: name,
        status: 'failed',
        error: e.message
      });
    }
  });

  try {
    SafeLedgerTestHarness.assertProductionStateUnchanged_(productionStateBefore);
    results.push({
      name: 'testProductionSheetsUntouched',
      status: 'passed',
      result: {
        status: 'success',
        message: 'Production-named sheets remained unchanged during guarded suite.'
      }
    });
  } catch (e) {
    results.push({
      name: 'testProductionSheetsUntouched',
      status: 'failed',
      error: e.message
    });
  }

  const failedResults = results.filter(r => r.status !== 'passed');
  if (failedResults.length > 0) {
    const summary = failedResults
      .map(r => `${r.name}: ${r.error}`)
      .join(' | ')
      .slice(0, 1000);
    throw new Error(`SafeLedgerTestSuite failed (${failedResults.length}/${results.length}): ${summary}`);
  }

  return {
    status: 'success',
    results: results
  };
}

function testSyncManagerCommitSuccess() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;

    SyncManager.updateUnifiedLedger(ss, 'Binance', [
      { ccy: 'BTC', amt: 1, type: 'Spot', status: 'Available', meta: 'seed' }
    ]);

    SyncManager.updateUnifiedLedger(ss, 'Bitget', [
      { ccy: 'USDT', amt: 100, type: 'Spot', status: 'Available', meta: 'old' }
    ]);

    const result = SyncManager.createResult('Bitget');
    result.assets.push(
      { ccy: 'BTC', amt: 0.5, type: 'Spot', status: 'Available', meta: 'new spot' },
      { ccy: 'USDT', amt: -50, type: 'Loan', status: 'Debt', meta: 'new debt' }
    );
    SyncManager.registerSourceCheck(result, { name: 'Spot', required: true, success: true, rows: 1 });
    SyncManager.registerSourceCheck(result, { name: 'Loan', required: true, success: true, rows: 1 });

    const committed = SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', result);
    assertSafeLedgerTest_(committed === true, 'Expected commit success.');

    const rows = getSafeLedgerRows_('Unified Assets');
    const bitgetRows = rows.filter(row => row[0] === 'Bitget');
    const binanceRows = rows.filter(row => row[0] === 'Binance');

    assertSafeLedgerTest_(bitgetRows.length === 2, 'Expected 2 Bitget rows after commit.');
    assertSafeLedgerTest_(binanceRows.length === 1, 'Expected Binance seed row to be preserved.');

    return {
      status: 'success',
      message: 'Commit success test passed.',
      rows: bitgetRows
    };
  });
}

function testSyncManagerCommitRequiredFailure() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;

    SyncManager.updateUnifiedLedger(ss, 'Bitget', [
      { ccy: 'BTC', amt: 0.25, type: 'Loan', status: 'Collateral', meta: 'previous success' }
    ]);

    const initialResult = SyncManager.createResult('Bitget');
    initialResult.assets.push({ ccy: 'BTC', amt: 0.25, type: 'Loan', status: 'Collateral', meta: 'baseline' });
    SyncManager.registerSourceCheck(initialResult, { name: 'Loan', required: true, success: true, rows: 1 });
    SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', initialResult);

    const failedResult = SyncManager.createResult('Bitget');
    failedResult.assets.push({ ccy: 'USDT', amt: 999, type: 'Spot', status: 'Available', meta: 'should not write' });
    SyncManager.registerSourceCheck(failedResult, { name: 'Spot', required: true, success: false, message: 'Injected required failure' });

    const committed = SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', failedResult);
    assertSafeLedgerTest_(committed === false, 'Expected commit abort on required failure.');

    const rows = getSafeLedgerRows_('Unified Assets').filter(row => row[0] === 'Bitget');
    assertSafeLedgerTest_(rows.length === 1, 'Expected previous Bitget snapshot to be preserved.');
    assertSafeLedgerTest_(rows[0][1] === 'BTC', 'Expected preserved Bitget currency to remain BTC.');

    return {
      status: 'success',
      message: 'Required failure preservation test passed.',
      rows: rows
    };
  });
}

function testSyncManagerOptionalFailure() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;

    const result = SyncManager.createResult('OKX');
    result.assets.push({ ccy: 'USDT', amt: 200, type: 'Funding', status: 'Available', meta: 'ok' });
    SyncManager.registerSourceCheck(result, { name: 'Account', required: true, success: true, rows: 1 });
    SyncManager.registerSourceCheck(result, { name: 'Structured', required: false, success: false, message: 'Injected optional failure' });

    const committed = SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', result);
    assertSafeLedgerTest_(committed === true, 'Expected commit to continue on optional failure.');

    const statusRows = getSafeLedgerRows_('Sync_Status').filter(row => row[0] === 'OKX');
    assertSafeLedgerTest_(statusRows.length === 1, 'Expected one Sync_Status row for OKX.');
    assertSafeLedgerTest_(String(statusRows[0][3]).toUpperCase() === 'COMPLETE', 'Expected COMPLETE status for optional failure case.');

    return {
      status: 'success',
      message: 'Optional failure warning test passed.',
      syncStatus: statusRows[0]
    };
  });
}

function testSyncManagerLockGuard() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;
    const originalLockFn = SyncManager.tryAcquireCommitLock_;

    try {
      SyncManager.tryAcquireCommitLock_ = function () { return null; };

      const result = SyncManager.createResult('Binance');
      result.assets.push({ ccy: 'BTC', amt: 1, type: 'Spot', status: 'Available', meta: 'lock test' });
      SyncManager.registerSourceCheck(result, { name: 'Spot', required: true, success: true, rows: 1 });

      const committed = SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', result);
      assertSafeLedgerTest_(committed === false, 'Expected commit to stop when lock is unavailable.');

      const rows = getSafeLedgerRows_('Unified Assets').filter(row => row[0] === 'Binance');
      assertSafeLedgerTest_(rows.length === 0, 'Expected no Binance rows to be written when lock acquisition fails.');

      return {
        status: 'success',
        message: 'Lock guard test passed.'
      };
    } finally {
      SyncManager.tryAcquireCommitLock_ = originalLockFn;
    }
  });
}

function testSyncStatusWrite() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;

    const success = SyncManager.createResult('BitoPro');
    success.assets.push({ ccy: 'TWD', amt: 1000, type: 'Spot', status: 'Available', meta: 'seed' });
    SyncManager.registerSourceCheck(success, { name: 'Balance', required: true, success: true, rows: 1 });
    SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', success);

    const failure = SyncManager.createResult('BitoPro');
    SyncManager.registerSourceCheck(failure, { name: 'Balance', required: true, success: false, message: 'Injected failure after success' });
    SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', failure);

    const statusRows = getSafeLedgerRows_('Sync_Status').filter(row => row[0] === 'BitoPro');
    assertSafeLedgerTest_(statusRows.length === 1, 'Expected one Sync_Status row for BitoPro.');
    assertSafeLedgerTest_(statusRows[0][2], 'Expected Last Success to remain populated after failure.');
    assertSafeLedgerTest_(String(statusRows[0][3]).toUpperCase() === 'FAILED', 'Expected FAILED status after second attempt.');

    const ledgerRows = getSafeLedgerRows_('Unified Assets').filter(row => row[0] === 'BitoPro');
    assertSafeLedgerTest_(ledgerRows.length === 1, 'Expected successful BitoPro snapshot to remain after failure.');

    return {
      status: 'success',
      message: 'Sync status persistence test passed.',
      syncStatus: statusRows[0]
    };
  });
}

function testStaleAttemptGuard() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;

    const newerResult = SyncManager.createResult('Bitget');
    newerResult.startedAt = '2026-04-03T08:00:10.000Z';
    newerResult.assets.push({ ccy: 'USDT', amt: 100, type: 'Spot', status: 'Available', meta: 'newer attempt' });
    SyncManager.registerSourceCheck(newerResult, { name: 'Spot', required: true, success: true, rows: 1 });
    const newerCommitted = SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', newerResult);
    assertSafeLedgerTest_(newerCommitted === true, 'Expected newer attempt to commit successfully.');

    const olderResult = SyncManager.createResult('Bitget');
    olderResult.startedAt = '2026-04-03T08:00:00.000Z';
    olderResult.assets.push({ ccy: 'BTC', amt: 1, type: 'Spot', status: 'Available', meta: 'older attempt' });
    SyncManager.registerSourceCheck(olderResult, { name: 'Spot', required: true, success: true, rows: 1 });
    const olderCommitted = SyncManager.commitExchangeResult(ss, 'Test_SafeLedgerSync', olderResult);
    assertSafeLedgerTest_(olderCommitted === false, 'Expected stale attempt to be skipped.');

    const ledgerRows = getSafeLedgerRows_('Unified Assets').filter(row => row[0] === 'Bitget');
    assertSafeLedgerTest_(ledgerRows.length === 1, 'Expected only newer Bitget snapshot to remain.');
    assertSafeLedgerTest_(ledgerRows[0][1] === 'USDT', 'Expected newer Bitget snapshot to remain in ledger.');

    const statusRows = getSafeLedgerRows_('Sync_Status').filter(row => row[0] === 'Bitget');
    const expectedAttemptAt = SyncManager.formatStatusTimestamp_(newerResult.startedAt);
    assertSafeLedgerTest_(statusRows.length === 1, 'Expected one Sync_Status row for Bitget.');
    assertSafeLedgerTest_(
      SyncManager.normalizeStatusTimestampValue_(statusRows[0][1]) === expectedAttemptAt,
      'Expected newer attempt timestamp to remain recorded.'
    );

    return {
      status: 'success',
      message: 'Stale attempt guard test passed.',
      syncStatus: statusRows[0]
    };
  });
}

function testStandaloneStatusWriteUsesLock() {
  return SafeLedgerTestHarness.withSandbox_(function (context) {
    setupSafeLedgerSyncTestEnvironment();
    const ss = context.ss;
    const originalLockFn = SyncManager.tryAcquireCommitLock_;
    let acquired = 0;
    let released = 0;

    try {
      SyncManager.tryAcquireCommitLock_ = function () {
        acquired++;
        return {
          releaseLock: function () {
            released++;
          }
        };
      };

      const newer = SyncManager.createResult('Bitget');
      newer.startedAt = '2026-04-03T08:00:10.000Z';
      newer.finishedAt = '2026-04-03T08:00:12.000Z';
      newer.status = 'complete';
      newer.rowCount = 1;
      newer.assets.push({ ccy: 'USDT', amt: 100, type: 'Spot', status: 'Available', meta: 'newer status' });
      const newerWrote = SyncManager.recordSyncStatus(ss, newer);
      assertSafeLedgerTest_(newerWrote === true, 'Expected standalone status write to succeed.');

      const olderFailure = SyncManager.createResult('Bitget');
      olderFailure.startedAt = '2026-04-03T08:00:00.000Z';
      olderFailure.finishedAt = '2026-04-03T08:00:01.000Z';
      olderFailure.status = 'failed';
      olderFailure.errors.push('Older failed attempt');
      const olderWrote = SyncManager.recordSyncStatus(ss, olderFailure);
      assertSafeLedgerTest_(olderWrote === false, 'Expected stale standalone status write to be skipped.');

      assertSafeLedgerTest_(acquired === 2, 'Expected standalone status writes to acquire the shared lock.');
      assertSafeLedgerTest_(released === 2, 'Expected standalone status writes to release the shared lock.');

      const statusRows = getSafeLedgerRows_('Sync_Status').filter(row => row[0] === 'Bitget');
      const expectedAttemptAt = SyncManager.formatStatusTimestamp_(newer.startedAt);
      assertSafeLedgerTest_(statusRows.length === 1, 'Expected one Sync_Status row after standalone writes.');
      assertSafeLedgerTest_(
        SyncManager.normalizeStatusTimestampValue_(statusRows[0][1]) === expectedAttemptAt,
        'Expected newer standalone attempt timestamp to remain recorded.'
      );
      assertSafeLedgerTest_(String(statusRows[0][3]).toUpperCase() === 'COMPLETE', 'Expected COMPLETE status to remain after stale failed write.');

      return {
        status: 'success',
        message: 'Standalone status write lock test passed.',
        syncStatus: statusRows[0]
      };
    } finally {
      SyncManager.tryAcquireCommitLock_ = originalLockFn;
    }
  });
}

function getSafeLedgerRows_(logicalSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resolvedName = SafeLedgerTestHarness.resolveSheetName_(logicalSheetName);
  const sheet = ss.getSheetByName(resolvedName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function assertSafeLedgerTest_(condition, message) {
  if (!condition) {
    throw new Error(`SafeLedgerTest failed: ${message}`);
  }
}
