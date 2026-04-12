/**
 * Guarded staging verification helpers for FX rate sync.
 *
 * Safety rules:
 * 1. Tests only run when staging Script Properties explicitly enable them.
 * 2. Tests only run against the approved staging spreadsheet ID.
 * 3. Tests only read/write `_TEST_*` sheets and never touch production tabs.
 */

const FxRateTestHarness = {
  ENABLED_PROPERTY: "FX_RATE_TEST_ENABLED",
  SPREADSHEET_ID_PROPERTY: "FX_RATE_TEST_SPREADSHEET_ID",
  SHEET_PREFIX_PROPERTY: "FX_RATE_TEST_PREFIX",
  DEFAULT_SHEET_PREFIX: "_TEST_",

  getConfig_: function () {
    const props = PropertiesService.getScriptProperties();
    const prefix = props.getProperty(this.SHEET_PREFIX_PROPERTY) || this.DEFAULT_SHEET_PREFIX;

    return {
      enabled: String(props.getProperty(this.ENABLED_PROPERTY) || "").toLowerCase() === "true",
      spreadsheetId: props.getProperty(this.SPREADSHEET_ID_PROPERTY) || "",
      prefix: prefix
    };
  },

  assertReady_: function () {
    const config = this.getConfig_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!config.enabled) {
      throw new Error(`FxRateTest blocked: set Script Property ${this.ENABLED_PROPERTY}=true in staging first.`);
    }

    if (!config.spreadsheetId) {
      throw new Error(`FxRateTest blocked: set Script Property ${this.SPREADSHEET_ID_PROPERTY} to the staging spreadsheet ID.`);
    }

    if (ss.getId() !== config.spreadsheetId) {
      throw new Error(`FxRateTest blocked: active spreadsheet ${ss.getId()} does not match approved staging spreadsheet ${config.spreadsheetId}.`);
    }

    return {
      ss: ss,
      sheetNames: {
        settings: `${config.prefix}參數設定`,
        temp: `${config.prefix}Temp`,
        formulas: `${config.prefix}FX_Formulas`
      }
    };
  },

  runSuite_: function () {
    const context = this.assertReady_();
    const tests = [
      this.testCollectRequiredCurrencyPairs_.bind(this, context),
      this.testUpdateWritesIdentityDirectAndInverse_.bind(this, context),
      this.testFailedFetchKeepsOldValueAndFatalWhenBlank_.bind(this, context),
      this.testOptionalPairFailureDoesNotFatal_.bind(this, context),
      this.testSafeWriteSkipsChangedRows_.bind(this, context),
      this.testFixedUsdTwdSemanticGuard_.bind(this, context)
    ];
    const failures = [];

    tests.forEach(testFn => {
      try {
        testFn();
      } catch (e) {
        failures.push(e.message || String(e));
      }
    });

    if (failures.length > 0) {
      throw new Error(`FxRateTest failed (${failures.length}): ${failures.join(" | ")}`);
    }

    return {
      ok: true,
      passed: tests.length,
      message: `FxRateTest passed ${tests.length} tests.`
    };
  },

  testCollectRequiredCurrencyPairs_: function (context) {
    this.seedSettingsSheet_(context, [
      ["", "TWD", "Timestamp"],
      ["TWD", 1, new Date("2026-04-12T00:00:00Z")]
    ]);

    const formulaSheet = this.resetSheet_(context.ss, context.sheetNames.formulas);
    formulaSheet.getRange(1, 1, 3, 3).setValues([
      ['=MY_GOOGLEFINANCE("USD","TWD")', "", ""],
      ["EUR", "TWD", '=MY_GOOGLEFINANCE(A2,B2)'],
      ["", "", ""]
    ]);

    const result = collectRequiredCurrencyPairs_(context.ss, {
      settingsSheetName: context.sheetNames.settings,
      sheetNames: [context.sheetNames.formulas]
    });
    const lookup = this.buildPairLookup_(result.pairs);

    this.assertTrue_(lookup["USD/TWD"], "formula literal pair USD/TWD should be collected");
    this.assertTrue_(lookup["EUR/TWD"], "formula cell-reference pair EUR/TWD should be collected");
  },

  testUpdateWritesIdentityDirectAndInverse_: function (context) {
    this.seedSettingsSheet_(context, [
      ["", "TWD", "Timestamp"],
      ["TWD", "", ""],
      ["USD", 31, new Date("2026-01-01T00:00:00Z")],
      ["EUR", "", ""]
    ]);
    this.resetSheet_(context.ss, context.sheetNames.temp);

    const result = updateAllFxRates({
      spreadsheet: context.ss,
      settingsSheetName: context.sheetNames.settings,
      tempSheetName: context.sheetNames.temp,
      now: new Date("2026-04-12T00:00:00Z"),
      requiredPairs: [
        { from: "TWD", to: "TWD" },
        { from: "USD", to: "TWD" },
        { from: "EUR", to: "TWD" }
      ],
      rateProvider: function (request) {
        if (request.from === "USD" && request.to === "TWD") return { rate: 32, source: "test:direct" };
        if (request.from === "TWD" && request.to === "EUR") return { rate: 0.025, source: "test:inverse" };
        return null;
      }
    });

    this.assertTrue_(result.status === "COMPLETE", `FX update should complete, got ${result.status}: ${result.message}`);

    const rows = context.ss.getSheetByName(context.sheetNames.settings).getRange(2, 1, 3, 3).getValues();
    this.assertNear_(1, rows[0][1], "TWD/TWD identity rate should be 1");
    this.assertNear_(32, rows[1][1], "USD/TWD direct rate should be written");
    this.assertNear_(40, rows[2][1], "EUR/TWD inverse rate should be written");
  },

  testFailedFetchKeepsOldValueAndFatalWhenBlank_: function (context) {
    this.seedSettingsSheet_(context, [
      ["", "TWD", "Timestamp"],
      ["USD", 31, new Date("2026-01-01T00:00:00Z")],
      ["JPY", "", ""]
    ]);
    this.resetSheet_(context.ss, context.sheetNames.temp);

    const result = updateAllFxRates({
      spreadsheet: context.ss,
      settingsSheetName: context.sheetNames.settings,
      tempSheetName: context.sheetNames.temp,
      now: new Date("2026-04-12T00:00:00Z"),
      requiredPairs: [
        { from: "USD", to: "TWD" },
        { from: "JPY", to: "TWD" }
      ],
      rateProvider: function () {
        return null;
      }
    });

    this.assertTrue_(result.fatal, "blank JPY/TWD with failed fetch should be fatal");
    this.assertTrue_(result.missingPairs.indexOf("JPY/TWD") >= 0, "JPY/TWD should be reported as missing usable rate");

    const rows = context.ss.getSheetByName(context.sheetNames.settings).getRange(2, 1, 2, 3).getValues();
    this.assertNear_(31, rows[0][1], "USD/TWD old value should be retained");
    this.assertTrue_(rows[1][1] === "", "JPY/TWD blank value should remain blank");
  },

  testSafeWriteSkipsChangedRows_: function (context) {
    this.seedSettingsSheet_(context, [
      ["", "TWD", "Timestamp"],
      ["USD", 31, new Date("2026-01-01T00:00:00Z")]
    ]);

    const summary = createFxRateUpdateSummary_();
    const update = {
      from: "USD",
      to: "TWD",
      rowNumber: 2,
      valueColumn: 2,
      timestampColumn: 3,
      expectedValue: 30,
      expectedTimestamp: new Date("2026-01-01T00:00:00Z"),
      rate: 32,
      updatedAt: new Date("2026-04-12T00:00:00Z")
    };

    const safeUpdates = filterSafeFxRateUpdatesBeforeWrite_(
      context.ss,
      [update],
      { sheetName: context.sheetNames.settings },
      new Date("2026-04-12T00:00:00Z"),
      summary
    );

    this.assertEqual_(0, safeUpdates.length, "changed row should be skipped before write");
    this.assertEqual_(1, summary.skippedChanged, "skippedChanged should increment");
  },

  testOptionalPairFailureDoesNotFatal_: function (context) {
    this.seedSettingsSheet_(context, [
      ["", "TWD", "Timestamp"],
      ["USD", 31, new Date("2026-01-01T00:00:00Z")],
      ["EUR", "", ""]
    ]);
    this.resetSheet_(context.ss, context.sheetNames.temp);

    const result = updateAllFxRates({
      spreadsheet: context.ss,
      settingsSheetName: context.sheetNames.settings,
      tempSheetName: context.sheetNames.temp,
      now: new Date("2026-04-12T00:00:00Z"),
      requiredPairs: [
        { from: "USD", to: "TWD" }
      ],
      rateProvider: function (request) {
        if (request.from === "USD" && request.to === "TWD") return { rate: 32, source: "test:direct" };
        return null;
      }
    });

    this.assertTrue_(!result.fatal, "optional EUR/TWD failure should not be fatal");
    this.assertTrue_(result.status === "WARNING", `optional failure should produce WARNING, got ${result.status}`);
    this.assertTrue_(result.optionalUnusablePairs.indexOf("EUR/TWD") >= 0, "EUR/TWD should be reported as optional unusable");
    this.assertTrue_(result.missingPairs.length === 0, "optional failure should not be reported as missing required rate");
  },

  testFixedUsdTwdSemanticGuard_: function (context) {
    const sheet = this.seedSettingsSheet_(context, [
      ["", "TWD", "Timestamp"],
      ["TWD", 1, new Date("2026-04-12T00:00:00Z")],
      ["JPY", "", ""],
      ["USD", 32, new Date("2026-04-12T00:00:00Z")]
    ]);
    const now = new Date("2026-04-12T01:00:00Z");
    const staleMs = 24 * 60 * 60 * 1000;

    let result = inspectFixedUsdTwdReference_(sheet, now, staleMs);
    this.assertTrue_(result.ok, "valid fixed USD/TWD cells should pass semantic guard");

    sheet.getRange("B1").setValue("JPY");
    result = inspectFixedUsdTwdReference_(sheet, now, staleMs);
    this.assertTrue_(!result.ok && result.message.indexOf("B1 expected TWD") >= 0, "B1 semantic mismatch should fail");

    sheet.getRange("B1").setValue("TWD");
    sheet.getRange("A4").setValue("EUR");
    result = inspectFixedUsdTwdReference_(sheet, now, staleMs);
    this.assertTrue_(!result.ok && result.message.indexOf("A4 expected USD") >= 0, "A4 semantic mismatch should fail");

    sheet.getRange("A4").setValue("USD");
    sheet.getRange("B4").setValue("");
    result = inspectFixedUsdTwdReference_(sheet, now, staleMs);
    this.assertTrue_(!result.ok && result.message.indexOf("B4 expected positive USD/TWD rate") >= 0, "B4 blank rate should fail");
  },

  seedSettingsSheet_: function (context, values) {
    const sheet = this.resetSheet_(context.ss, context.sheetNames.settings);
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    return sheet;
  },

  resetSheet_: function (ss, sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.clear();
    return sheet;
  },

  buildPairLookup_: function (pairs) {
    const lookup = {};
    (pairs || []).forEach(pair => {
      lookup[`${pair.from}/${pair.to}`] = true;
    });
    return lookup;
  },

  assertEqual_: function (expected, actual, message) {
    if (expected !== actual) {
      throw new Error(`${message}: expected ${expected}, got ${actual}`);
    }
  },

  assertNear_: function (expected, actual, message) {
    const actualNumber = parseFloat(actual);
    if (!isFinite(actualNumber) || Math.abs(expected - actualNumber) > 1e-9) {
      throw new Error(`${message}: expected ${expected}, got ${actual}`);
    }
  },

  assertTrue_: function (condition, message) {
    if (!condition) {
      throw new Error(message);
    }
  }
};

function runFxRateTestSuite() {
  return FxRateTestHarness.runSuite_();
}
