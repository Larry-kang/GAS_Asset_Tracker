if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadEngineContext() {
    const repoRoot = path.resolve(__dirname, '..');
    const sandbox = {
      JSON,
      Date,
      Math,
      parseFloat,
      isFinite,
      console,
      Set,
      Logger: { log() {} },
      LogService: { info() {}, warn() {}, error() {} },
      Settings: { get() { return ''; } }
    };
    const context = vm.createContext(sandbox);
    const files = [
      'Config.js',
      'Strategy_BullStateMachine.js',
      'Strategy_StockExposure.js',
      'Strategy_RebalanceEngine.js',
      'Strategy_ReportFormatter.js',
      'Repo_FreshnessAuditor.js',
      'Core_StrategicEngine.js'
    ];
    files.forEach(file => {
      const filePath = path.join(repoRoot, file);
      if (fs.existsSync(filePath)) {
        vm.runInContext(fs.readFileSync(filePath, 'utf8'), context, { filename: file });
      }
    });
    return context;
  }

  test('getRebalanceTargets detects overweight and underweight groups', () => {
    const context = loadEngineContext();
    // Total = 1,000,000. L1 = 800,000 (80%), L2 = 100,000 (10%), L3 = 100,000 (10%)
    // L1 defaultTarget = 0.70, L2 defaultTarget = 0.20, L3 defaultTarget = 0.10
    const assetGroups = [
      { id: 'L1', name: 'Digital Reserve', defaultTarget: 0.70, value: 800000, tickers: ['BTC', 'IBIT'] },
      { id: 'L2', name: 'Equity Exposure', defaultTarget: 0.20, value: 100000, tickers: ['00713', 'QQQ'] },
      { id: 'L3', name: 'Tactical Liquidity', defaultTarget: 0.10, value: 100000, tickers: ['CASH_TWD'] }
    ];

    const targets = context.getRebalanceTargets(assetGroups, 1000000, {});
    assert.ok(targets.length >= 2);

    const trimL1 = targets.find(t => t.id === 'L1');
    assert.ok(trimL1);
    assert.equal(trimL1.action, 'TRIM');
    assert.equal(trimL1.currentWeight, 0.80);
    assert.equal(trimL1.targetWeight, 0.70);

    const addL2 = targets.find(t => t.id === 'L2');
    assert.ok(addL2);
    assert.equal(addL2.action, 'ADD');
    assert.equal(addL2.currentWeight, 0.10);
    assert.equal(addL2.targetWeight, 0.20);
  });

  test('rankRebalanceProducts_ prioritizes configured tickers properly', () => {
    const context = loadEngineContext();
    const target = { id: 'L1', action: 'ADD', tickers: ['BTC', 'IBIT', 'BTC_Spot'] };
    const portfolioSummary = { BTC: 500000, IBIT: 200000 };
    const assetGroups = [
      { id: 'L1', name: 'Layer 1: Digital Reserve (Attack)', defaultTarget: 0.70, tickers: ['IBIT', 'BTC_Spot', 'BTC'] }
    ];

    const hints = context.buildRebalanceProductHints_(target, portfolioSummary, assetGroups);
    assert.ok(hints.primary.length > 0);
    assert.equal(hints.primary[0], 'BTC');
  });
}
