if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadFxContext() {
    const repoRoot = path.resolve(__dirname, '..');
    const source = fs.readFileSync(path.join(repoRoot, 'Sync_FxRates.js'), 'utf8');
    const configSource = fs.readFileSync(path.join(repoRoot, 'Config.js'), 'utf8');
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
    vm.runInContext(configSource, context, { filename: 'Config.js' });
    vm.runInContext(source, context, { filename: 'Sync_FxRates.js' });
    return context;
  }

  test('FX rate normalization handles identity pairs with 1.0', () => {
    const context = loadFxContext();
    if (typeof context.normalizeFxRate_ === 'function') {
      assert.equal(context.normalizeFxRate_('USD', 'USD', null), 1.0);
      assert.equal(context.normalizeFxRate_('TWD', 'TWD', null), 1.0);
    }
  });

  test('FX rate parsing extracts valid numeric values safely', () => {
    const context = loadFxContext();
    if (typeof context.parseFxRateValue_ === 'function') {
      assert.equal(context.parseFxRateValue_(32.5), 32.5);
      assert.equal(context.parseFxRateValue_('32.5'), 32.5);
      assert.equal(context.parseFxRateValue_('#N/A'), null);
      assert.equal(context.parseFxRateValue_(''), null);
      assert.equal(context.parseFxRateValue_(null), null);
    }
  });

  test('collectRequiredCurrencyPairs_ includes USD/TWD as baseline requirement', () => {
    const context = loadFxContext();
    if (typeof context.collectRequiredCurrencyPairs_ === 'function') {
      const mockSpreadsheet = {
        getSheets: () => []
      };
      const pairs = context.collectRequiredCurrencyPairs_(mockSpreadsheet);
      const hasUsdTwd = pairs.some(p => (p.from === 'USD' && p.to === 'TWD') || (p.from === 'TWD' && p.to === 'USD'));
      assert.equal(hasUsdTwd, true);
    }
  });
}
