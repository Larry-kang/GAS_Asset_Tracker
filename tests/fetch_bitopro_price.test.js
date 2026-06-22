if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadScripts(files, sandbox) {
    const context = vm.createContext(sandbox);
    files.forEach((file) => {
      const source = fs.readFileSync(file, 'utf8');
      vm.runInContext(source, context, { filename: file });
    });
    return context;
  }

  function createScriptCache(initialValues) {
    const store = new Map(Object.entries(initialValues || {}));
    return {
      get(key) {
        return store.has(key) ? store.get(key) : null;
      },
      put(key, value) {
        store.set(key, String(value));
      }
    };
  }

  function createBitoSandbox(overrides) {
    const repoRoot = path.resolve(__dirname, '..');
    const sandbox = Object.assign({
      JSON,
      Date,
      Math,
      console,
      Logger: { log() {} },
      LogService: { info() {}, warn() {}, error() {} },
      ScriptCache: createScriptCache(),
      SpreadsheetApp: { getActiveSpreadsheet() { return {}; } },
      SettingsMatrixRepo: { lookupRate() { return null; } },
      getCryptoPrice() {
        throw new Error('generic fallback should not be used in this test');
      }
    }, overrides || {});

    const context = loadScripts([
      path.join(repoRoot, 'Fetch_BitoProPrice.js'),
      path.join(repoRoot, 'Fetch_CryptoPrice.js')
    ], sandbox);

    return { context, sandbox };
  }

  test('fetchBitoProTickerLastPrice_ parses array-based ticker payloads from the official API format', () => {
    const fetchCalls = [];
    const { context } = createBitoSandbox({
      UrlFetchApp: {
        fetch(url) {
          fetchCalls.push(url);
          return {
            getResponseCode() {
              return 200;
            },
            getContentText() {
              return JSON.stringify({
                data: [
                  { pair: 'btc_twd', lastPrice: '3000000' },
                  { pair: 'bito_usdt', lastPrice: '0.1685' }
                ]
              });
            }
          };
        }
      }
    });

    const price = vm.runInContext("fetchBitoProTickerLastPrice_('bito_usdt', true)", context);

    assert.equal(price, 0.1685);
    assert.equal(fetchCalls.length, 1);
  });

  test('fetchCryptoPrice uses the shared crypto cache for BITO before hitting the BitoPro API', () => {
    let networkCalls = 0;
    const { context } = createBitoSandbox({
      ScriptCache: createScriptCache({
        PRICE_CRYPTO_BITO_USD: '0.1496'
      }),
      UrlFetchApp: {
        fetch() {
          networkCalls += 1;
          throw new Error('network should not be called when shared cache is warm');
        }
      }
    });

    const price = vm.runInContext("fetchCryptoPrice('BITO')", context);

    assert.equal(price, 0.1496);
    assert.equal(networkCalls, 0);
  });

  test('fetchCryptoPrice returns null for BITO when BitoPro is unavailable instead of falling through to generic same-symbol sources', () => {
    let genericFallbackCalls = 0;
    const { context } = createBitoSandbox({
      UrlFetchApp: {
        fetch() {
          return {
            getResponseCode() {
              return 503;
            },
            getContentText() {
              return '{}';
            }
          };
        }
      },
      getCryptoPrice() {
        genericFallbackCalls += 1;
        return 0.072639;
      }
    });

    const price = vm.runInContext("fetchCryptoPrice('BITO', 'USD', true)", context);

    assert.equal(price, null);
    assert.equal(genericFallbackCalls, 0);
  });

  test('fetchBitoProTickerLastPrice_ backs off briefly after an HTTP 503 so repeated runs do not hammer the same failing pair', () => {
    let networkCalls = 0;
    const { context } = createBitoSandbox({
      UrlFetchApp: {
        fetch() {
          networkCalls += 1;
          return {
            getResponseCode() {
              return 503;
            },
            getContentText() {
              return '{}';
            }
          };
        }
      }
    });

    const first = vm.runInContext("fetchBitoProTickerLastPrice_('bito_usdt', false)", context);
    const second = vm.runInContext("fetchBitoProTickerLastPrice_('bito_usdt', false)", context);

    assert.equal(first, null);
    assert.equal(second, null);
    assert.equal(networkCalls, 1);
  });
}
