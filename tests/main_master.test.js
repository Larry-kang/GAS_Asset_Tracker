if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadMasterContext(overrides = {}) {
    const repoRoot = path.resolve(__dirname, '..');
    const configSource = fs.readFileSync(path.join(repoRoot, 'Config.js'), 'utf8');
    const masterSource = fs.readFileSync(path.join(repoRoot, 'Core_MainMaster.js'), 'utf8');
    const logs = { info: [], warn: [], error: [] };

    const sandbox = {
      JSON,
      Date,
      Math,
      console: {
        log: (...args) => logs.info.push(args.join(' ')),
        warn: (...args) => logs.warn.push(args.join(' ')),
        error: (...args) => logs.error.push(args.join(' '))
      },
      Utilities: { sleep: () => {} },
      ExchangeRegistry: {
        getActive: () => [
          { moduleName: 'Binance', functionName: 'getBinanceBalance' }
        ]
      },
      LogService: {
        cleanupOldLogs: () => {}
      },
      ...overrides
    };

    const context = vm.createContext(sandbox);
    vm.runInContext(configSource, context, { filename: 'Config.js' });
    vm.runInContext(masterSource, context, { filename: 'Core_MainMaster.js' });
    context._logs = logs;
    return context;
  }

  test('runAutomationMaster executes all pipeline stages in order', () => {
    const executionOrder = [];
    const context = loadMasterContext({
      syncCurrencyPairs: () => { executionOrder.push('syncCurrencyPairs'); return { status: 'COMPLETE' }; },
      updateAllFxRates: () => { executionOrder.push('updateAllFxRates'); return { status: 'COMPLETE' }; },
      updateAllPrices: () => { executionOrder.push('updateAllPrices'); return { status: 'COMPLETE' }; },
      getBinanceBalance: () => { executionOrder.push('getBinanceBalance'); return true; },
      runStrategicMonitor: () => { executionOrder.push('runStrategicMonitor'); return true; }
    });

    const result = context.runAutomationMaster();
    assert.equal(result.ok, true);
    assert.equal(result.status, 'COMPLETE');
    assert.deepEqual(executionOrder, [
      'syncCurrencyPairs',
      'updateAllFxRates',
      'updateAllPrices',
      'getBinanceBalance',
      'runStrategicMonitor'
    ]);
  });

  test('runAutomationMaster stops when FX update produces fatal error', () => {
    const executed = [];
    const context = loadMasterContext({
      syncCurrencyPairs: () => { executed.push('syncCurrencyPairs'); return { status: 'FAILED', fatal: true, message: 'Rate Limit' }; },
      updateAllPrices: () => { executed.push('updateAllPrices'); },
      runStrategicMonitor: () => { executed.push('runStrategicMonitor'); }
    });

    const result = context.runAutomationMaster();
    assert.equal(result.ok, false);
    assert.equal(result.fatal, true);
    assert.match(result.message, /Currency pair sync failed/);
    assert.deepEqual(executed, ['syncCurrencyPairs']);
  });

  test('runAutomationMaster gracefully detects timeout and aborts subsequent tasks with warning', () => {
    let nowOffset = 0;
    const executed = [];
    const context = loadMasterContext({
      Date: {
        now: () => 1000000 + nowOffset
      },
      syncCurrencyPairs: () => { executed.push('syncCurrencyPairs'); return { status: 'COMPLETE' }; },
      updateAllFxRates: () => {
        executed.push('updateAllFxRates');
        nowOffset = 300 * 1000; // Fast-forward time by 300 seconds (> 240s threshold)
        return { status: 'COMPLETE' };
      },
      updateAllPrices: () => { executed.push('updateAllPrices'); return { status: 'COMPLETE' }; },
      getBinanceBalance: () => { executed.push('getBinanceBalance'); return true; },
      runStrategicMonitor: () => { executed.push('runStrategicMonitor'); return true; }
    });

    const result = context.runAutomationMaster({ timeoutSeconds: 240 });
    assert.equal(result.ok, false);
    assert.equal(result.status, 'TIMEOUT_GUARD');
    assert.equal(result.fatal, false);
    assert.deepEqual(executed, ['syncCurrencyPairs', 'updateAllFxRates']);
  });
}
