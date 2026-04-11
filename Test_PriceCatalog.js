/**
 * Guarded staging verification helpers for the price catalog refactor.
 *
 * Safety rules:
 * 1. Tests only run when staging Script Properties explicitly enable them.
 * 2. Tests only run against the approved staging spreadsheet ID.
 * 3. Tests only read/write `_TEST_*` sheets and never touch production tabs.
 */

const PriceCatalogTestHarness = {
  ENABLED_PROPERTY: "PRICE_CATALOG_TEST_ENABLED",
  SPREADSHEET_ID_PROPERTY: "PRICE_CATALOG_TEST_SPREADSHEET_ID",
  SHEET_PREFIX_PROPERTY: "PRICE_CATALOG_TEST_PREFIX",
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
      throw new Error(`PriceCatalogTest blocked: set Script Property ${this.ENABLED_PROPERTY}=true in staging first.`);
    }

    if (!config.spreadsheetId) {
      throw new Error(`PriceCatalogTest blocked: set Script Property ${this.SPREADSHEET_ID_PROPERTY} to the staging spreadsheet ID.`);
    }

    if (ss.getId() !== config.spreadsheetId) {
      throw new Error(`PriceCatalogTest blocked: active spreadsheet ${ss.getId()} does not match approved staging spreadsheet ${config.spreadsheetId}.`);
    }

    return {
      ss: ss,
      sheetNames: {
        assetAllocation: `${config.prefix}資產配置`,
        priceCache: `${config.prefix}價格暫存`
      }
    };
  },

  runSuite_: function () {
    const context = this.assertReady_();
    const tests = [
      this.testDiscoveryAddsMissingTickers_.bind(this, context),
      this.testManualLockAndInvalidDateRules_.bind(this)
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
      throw new Error(`PriceCatalogTest failed (${failures.length}): ${failures.join(" | ")}`);
    }

    return {
      ok: true,
      passed: tests.length,
      message: `PriceCatalogTest passed ${tests.length} tests.`
    };
  },

  testDiscoveryAddsMissingTickers_: function (context) {
    this.seedDiscoverySheets_(context);

    const rows = PriceCacheRepo.readRows(context.ss, { sheetName: context.sheetNames.priceCache });
    const result = syncPriceCatalogFromAssetAllocation_(context.ss, rows, {
      assetAllocationSheetName: context.sheetNames.assetAllocation,
      priceCacheSheetName: context.sheetNames.priceCache
    });

    this.assertEqual_(2, result.addedCount, "discovery should add 0056 and 00878 only");
    const afterRows = PriceCacheRepo.readRows(context.ss, { sheetName: context.sheetNames.priceCache });
    const lookup = this.buildTickerLookup_(afterRows);

    this.assertTrue_(lookup["0056"] && lookup["0056"].assetType === "股票", "0056 should be added as 股票");
    this.assertTrue_(lookup["00878"] && lookup["00878"].assetType === "股票", "00878 should be added as 股票");
    this.assertTrue_(lookup.BITO && lookup.BITO.price === 0.1496, "existing BITO manual row should not be overwritten");
    this.assertTrue_(!lookup.CASH_TWD, "cash rows should not be added to price catalog");
  },

  testManualLockAndInvalidDateRules_: function () {
    const now = new Date("2026-04-11T00:00:00Z");
    const future = new Date("2026-12-02T00:00:00Z");

    this.assertTrue_(isPriceCacheRowLocked_({ updatedAt: future }, now), "future updatedAt should be treated as manual lock");
    this.assertTrue_(!isPriceCacheRowDueForUpdate_({ updatedAt: future }, now, 60000), "manual lock should not be due");
    this.assertTrue_(isPriceCacheRowDueForUpdate_({ updatedAt: "not-a-date" }, now, 60000), "invalid date should be due");
    this.assertTrue_(isPriceCacheRowDueForUpdate_({ updatedAt: "" }, now, 60000), "blank date should be due");
  },

  seedDiscoverySheets_: function (context) {
    const priceSheet = this.resetSheet_(context.ss, context.sheetNames.priceCache);
    priceSheet.getRange(1, 1, 3, 4).setValues([
      PriceCacheRepo.HEADERS,
      ["BTC", "數位幣", 66000, new Date("2026-04-11T00:00:00Z")],
      ["BITO", "數位幣", 0.1496, new Date("2026-12-02T00:00:00Z")]
    ]);

    const assetSheet = this.resetSheet_(context.ss, context.sheetNames.assetAllocation);
    assetSheet.getRange(1, 1, 7, 7).setValues([
      ["", "", "", "", "", "", ""],
      ["", "", "", "", "", "", ""],
      ["交易所", "類型", "股票代號/銀行", "股數/幣數", "幣別", "現價", "類型"],
      ["永豐", "股票", "0056", 1, "TWD", "", "Credit Base"],
      ["永豐", "股票", "00878", 1, "TWD", "", "Credit Base"],
      ["現金", "現金", "CASH_TWD", 1, "TWD", "", "Tactical Liquidity"],
      ["BitoPro", "數位幣", "BITO", 1, "USD", "", "Digital Reserve"]
    ]);
  },

  resetSheet_: function (ss, sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.clear();
    return sheet;
  },

  buildTickerLookup_: function (rows) {
    const lookup = {};
    (rows || []).forEach(row => {
      const ticker = String(row.ticker || "").trim().toUpperCase();
      if (ticker) lookup[ticker] = row;
    });
    return lookup;
  },

  assertEqual_: function (expected, actual, message) {
    if (expected !== actual) {
      throw new Error(`${message}: expected ${expected}, got ${actual}`);
    }
  },

  assertTrue_: function (condition, message) {
    if (!condition) {
      throw new Error(message);
    }
  }
};

function runPriceCatalogTestSuite() {
  return PriceCatalogTestHarness.runSuite_();
}
