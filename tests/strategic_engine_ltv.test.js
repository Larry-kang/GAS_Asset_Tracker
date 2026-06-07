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

  test('Version B accumulation regimes expose the new LTV bands', () => {
    const context = loadStrategicEngineContext();
    const panic = vm.runInContext('getBtcRegime_(0.58, -0.55, 0.34, 12)', context);
    const deepValue = vm.runInContext('getBtcRegime_(0.70, -0.40, 0.34, 12)', context);
    const accumulate = vm.runInContext('getBtcRegime_(0.77, -0.30, 0.34, 12)', context);
    const neutral = vm.runInContext('getBtcRegime_(1.12, -0.15, 0.34, 12)', context);

    assert.equal(panic.targetLtvMin, 0.50);
    assert.equal(panic.targetLtvMax, 0.55);
    assert.equal(deepValue.targetLtvMin, 0.45);
    assert.equal(deepValue.targetLtvMax, 0.50);
    assert.equal(accumulate.targetLtvMin, 0.40);
    assert.equal(accumulate.targetLtvMax, 0.45);
    assert.equal(neutral.targetLtvMin, 0.30);
    assert.equal(neutral.targetLtvMax, 0.35);
  });

  test('Version B guardrails keep 50 percent as no-new-borrow, 55 percent as DEFCON, and 60 percent as hard stop', () => {
    const context = loadStrategicEngineContext();
    const blockedAt50 = vm.runInContext('getBtcRegime_(0.70, -0.40, 0.50, 12)', context);
    const blockedAt55 = vm.runInContext('getBtcRegime_(0.58, -0.55, 0.55, 12)', context);
    const blockedAt60 = vm.runInContext('getBtcRegime_(0.58, -0.55, 0.60, 12)', context);

    assert.equal(blockedAt50.regime, 'STRETCH_LOCK');
    assert.equal(blockedAt50.restockAllowed, false);
    assert.equal(blockedAt55.regime, 'DEFCON_1');
    assert.equal(blockedAt55.restockAllowed, false);
    assert.equal(blockedAt60.regime, 'HARD_CAP_BREACH');
    assert.match(blockedAt60.guardrailAction, /強制去槓桿|停止所有買入/);
  });

  test('Version B guardrail formatting reflects the updated neutral band', () => {
    const context = loadStrategicEngineContext();
    const neutral = vm.runInContext('getBtcRegime_(1.12, -0.15, 0.22, 12)', context);
    const range = vm.runInContext('formatBtcLtvRange_(getBtcRegime_(1.12, -0.15, 0.22, 12))', context);

    assert.equal(neutral.regime, 'NEUTRAL');
    assert.equal(range, '30% - 35%');
  });
}
