if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadStrategicEngineContext() {
    const repoRoot = path.resolve(__dirname, '..');
    const source = fs.readFileSync(path.join(repoRoot, 'Core_StrategicEngine.js'), 'utf8');
    const sandbox = {
      JSON,
      Date,
      Math,
      console,
      Settings: { get() { return ''; } },
      LogService: { info() {}, warn() {}, error() {} },
      SpreadsheetApp: { getActiveSpreadsheet() { return {}; } },
      DataCache: { clear() {}, getValues() { return null; } },
      Config: { BTC_MARTINGALE: { ENABLED: false, LEVELS: [] } }
    };
    const context = vm.createContext(sandbox);
    vm.runInContext(source, context, { filename: 'Core_StrategicEngine.js' });
    return context;
  }

  function buildFixture(includeCapitalBase = false) {
    const summary = {
      StockTargetExposureRatio: 1,
      TaiwanTargetExposureWeight: 0.2,
      NasdaqTargetExposureWeight: 0.8,
      CashBufferTWD: 100000,
      StockDebtStatus: 'REPAYMENT_PENDING',
      OriginalCoreTicker: '00662',
      OriginalCoreMinPreSplitEquivalentQty: 1000,
      CorporateAction00662Status: 'PENDING_OFFICIAL'
    };
    if (includeCapitalBase) summary.StockCapitalBaseTWD = 692831;

    return {
      available: true,
      summary,
      positions: [
        {
          ticker: '00713', quantity: 1000, valueTwd: 60550, region: 'TW',
          exposureMultiplier: 1, effectiveExposureTwd: 60550,
          strategyRole: 'LegacyCore', settlementStatus: 'TRADE_PENDING',
          pledgeStatus: 'REPAYMENT_PENDING'
        },
        {
          ticker: '00662', quantity: 1000, valueTwd: 120250, region: 'NASDAQ',
          exposureMultiplier: 1, effectiveExposureTwd: 120250,
          strategyRole: 'OriginalCore', settlementStatus: 'TRADE_PENDING',
          pledgeStatus: 'NOT_PLEDGED'
        },
        {
          ticker: '00670L', quantity: 950, valueTwd: 186580, region: 'NASDAQ',
          exposureMultiplier: 2, effectiveExposureTwd: 373160,
          strategyRole: 'ExposureCore', settlementStatus: 'TRADE_PENDING',
          pledgeStatus: 'NOT_PLEDGED'
        },
        {
          ticker: '00685L', quantity: 6000, valueTwd: 68820, region: 'TW',
          exposureMultiplier: 2, effectiveExposureTwd: 137640,
          strategyRole: 'ExposureCore', settlementStatus: 'TRADE_PENDING',
          pledgeStatus: 'NOT_PLEDGED'
        },
        {
          ticker: 'IBIT', quantity: 683.56, valueTwd: 813425, region: 'US',
          exposureMultiplier: 1, effectiveExposureTwd: 813425,
          strategyRole: 'BTC_ETF_CORE', settlementStatus: 'SETTLED',
          pledgeStatus: 'NOT_PLEDGED'
        }
      ]
    };
  }

  test('stock exposure strategy uses effective exposure and excludes non-stock strategy regions', () => {
    const context = loadStrategicEngineContext();
    context.fixture = buildFixture(true);
    const strategy = vm.runInContext('buildStockExposureStrategy_(fixture)', context);

    assert.equal(strategy.positions.length, 4);
    assert.equal(strategy.marketValueTwd, 436200);
    assert.equal(strategy.grossExposureTwd, 691600);
    assert.equal(strategy.taiwanExposureTwd, 198190);
    assert.equal(strategy.nasdaqExposureTwd, 493410);
    assert.equal(strategy.debtStatus, 'REPAYMENT_PENDING');
    assert.equal(strategy.settlementStatus, 'TRADE_PENDING');
    assert.equal(strategy.isPending, true);
    assert.ok(Math.abs(strategy.exposureRatio - (691600 / 692831)) < 1e-12);
  });

  test('stock exposure snapshot does not invent a ratio without StockCapitalBaseTWD', () => {
    const context = loadStrategicEngineContext();
    context.fixture = buildFixture(false);
    const snapshot = vm.runInContext(
      'buildStockExposureSnapshot_(buildStockExposureStrategy_(fixture))',
      context
    );

    assert.match(snapshot, /名目曝險: 691,600 TWD/);
    assert.match(snapshot, /N\/A（Strategy_Config 缺少 StockCapitalBaseTWD）/);
    assert.match(snapshot, /台灣 28\.66%/);
    assert.match(snapshot, /NASDAQ 71\.34%/);
    assert.match(snapshot, /TRADE_PENDING/);
    assert.match(snapshot, /REPAYMENT_PENDING/);
  });

  test('pending stock restructuring replaces the legacy leverage advice', () => {
    const context = loadStrategicEngineContext();
    context.fixture = buildFixture(true);
    const alert = vm.runInContext(
      'buildStockExposureAlert_(buildStockExposureStrategy_(fixture))',
      context
    );

    assert.match(alert.level, /股票重整待完成/);
    assert.match(alert.action, /暫停新增股票操作與股票質押/);
    assert.doesNotMatch(alert.action, /67%|00713 佔比過低|Target Loan/);
  });
}
