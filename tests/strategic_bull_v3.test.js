if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadStrategyContext() {
    const repoRoot = path.resolve(__dirname, '..');
    const sandbox = {
      JSON,
      Date,
      Math,
      parseFloat,
      isFinite,
      console,
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

  test('calculateBtcDrawdownFromATH_ calculates drawdown accurately', () => {
    const context = loadStrategyContext();
    const dd = context.calculateBtcDrawdownFromATH_(60000, 100000);
    assert.equal(dd, -0.40);

    assert.equal(context.calculateBtcDrawdownFromATH_(0, 100000), null);
    assert.equal(context.calculateBtcDrawdownFromATH_(60000, 0), null);
    assert.equal(context.calculateBtcDrawdownFromATH_(null, 100000), null);
  });

  test('getBtcRegime_ triggers DEFCON 1 when LTV is >= 55% but < 60%', () => {
    const context = loadStrategyContext();
    const regime = context.getBtcRegime_(0.5, -0.2, 0.56, 12);
    assert.equal(regime.regime, 'DEFCON_1');
    assert.equal(regime.action, 'STOP_BUYING_REPAY_DEBT');
    assert.equal(regime.restockMode, 'DEFCON');
    assert.equal(regime.severity, 'CRITICAL');
  });

  test('getBtcRegime_ triggers HARD_CAP_BREACH when LTV >= 60%', () => {
    const context = loadStrategyContext();
    const regime = context.getBtcRegime_(0.5, -0.2, 0.62, 12);
    assert.equal(regime.regime, 'HARD_CAP_BREACH');
    assert.equal(regime.action, 'FORCE_DELEVERAGE');
    assert.equal(regime.severity, 'CRITICAL');
  });

  test('applyBtcRestockGate_ clamps restock when runway < 6 months or LTV >= 45%', () => {
    const context = loadStrategyContext();
    const baseRegime = {
      regime: 'ACCUMULATION_STAGE_1',
      restockMode: 'NORMAL',
      restockAllowed: true,
      severity: 'INFO'
    };

    // LTV >= 45% -> DCA_ONLY
    const gatedByLtv = context.applyBtcRestockGate_(baseRegime, 0.46, 12);
    assert.equal(gatedByLtv.restockAllowed, false);
    assert.equal(gatedByLtv.restockMode, 'DCA_ONLY');

    // Runway < 6 -> DCA_ONLY
    const gatedByRunway = context.applyBtcRestockGate_(baseRegime, 0.30, 4);
    assert.equal(gatedByRunway.restockAllowed, false);
    assert.equal(gatedByRunway.restockMode, 'DCA_ONLY');

    // Safe conditions -> restockAllowed = true
    const allowed = context.applyBtcRestockGate_(baseRegime, 0.30, 12);
    assert.equal(allowed.restockAllowed, true);
  });
}
