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

  function evaluateInContext(expression, context) {
    return vm.runInContext(expression, context);
  }

  function createWebhookSandbox() {
    return {
      JSON,
      Date,
      console,
      ContentService: {
        createTextOutput(content) {
          return {
            content,
            getContent() {
              return content;
            }
          };
        }
      },
      Settings: {
        get() {
          return '';
        },
        set() {}
      },
      LogService: {
        info() {},
        warn() {},
        error() {}
      },
      buildContext() {
        return {
          portfolioSummary: {
            BTC_Spot: 123456,
            IBIT: 654321
          },
          rawPortfolio: [],
          indicators: {},
          market: {}
        };
      },
      buildFreshContext() {
        return {
          portfolioSummary: {
            BTC_Spot: 123456,
            IBIT: 654321
          },
          rawPortfolio: [],
          indicators: {},
          market: {}
        };
      },
      getInventoryExportBundle_() {
        return {
          available: true,
          schemaVersion: '2',
          summary: {
            BTC_Spot_Quantity: 0.5,
            IBIT_Quantity: 10,
            IBIT_BTC_Per_Share: 0.00024
          },
          positions: [
            {
              ticker: 'BTC_Spot',
              account: 'OKX',
              quantity: 0.5,
              valueTwd: 123456
            }
          ]
        };
      }
    };
  }

  test('handleGetInventory keeps legacy context and appends inventoryExport when available', () => {
    const repoRoot = path.resolve(__dirname, '..');
    const sandbox = createWebhookSandbox();
    const context = loadScripts([
      path.join(repoRoot, 'Event_Webhook.js')
    ], sandbox);

    const response = context.handleGetInventory({});
    const payload = JSON.parse(response.getContent());

    assert.equal(payload.status, 'success');
    assert.equal(payload.data.portfolioSummary.BTC_Spot, 123456);
    assert.equal(payload.data.inventoryExport.schemaVersion, '2');
    assert.equal(payload.data.inventoryExport.summary.BTC_Spot_Quantity, 0.5);
    assert.equal(payload.data.inventoryExport.positions[0].account, 'OKX');
  });

  test('InventoryExportRepo parses row-based summary and header-driven positions without hardcoding extra fields', () => {
    const repoRoot = path.resolve(__dirname, '..');
    const context = loadScripts([
      path.join(repoRoot, 'Repo_InventoryExport.js')
    ], {
      JSON,
      Date,
      console
    });
    const repo = evaluateInContext('InventoryExportRepo', context);

    const bundle = repo.buildExportBundleFromTables_(
      [
        ['key', 'value', 'source', 'asOf'],
        ['BTC_Spot_Quantity', '0.5', 'formula', '2026-06-01T00:00:00Z'],
        ['IBIT_Quantity', 10, 'manual', '2026-06-01T00:00:00Z'],
        ['IBIT_BTC_Per_Share', '0.00024', 'blackrock', '2026-06-01T00:00:00Z']
      ],
      [
        ['ticker', 'account', 'quantity', 'valueTwd', 'note'],
        ['BTC_Spot', 'OKX', '0.5', '123456', 'cold wallet'],
        ['IBIT', 'IBKR', 10, 654321, 'tax lot a']
      ]
    );

    assert.equal(bundle.available, true);
    assert.equal(bundle.summary.BTC_Spot_Quantity, 0.5);
    assert.equal(bundle.summary.IBIT_Quantity, 10);
    assert.equal(bundle.summaryRows[0].source, 'formula');
    assert.equal(bundle.positions[0].quantity, 0.5);
    assert.equal(bundle.positions[0].note, 'cold wallet');
    assert.equal(bundle.positions[1].account, 'IBKR');
  });

  test('OKX DCA summary export helpers normalize headers and build stable rows', () => {
    const repoRoot = path.resolve(__dirname, '..');
    const context = loadScripts([
      path.join(repoRoot, 'Event_Webhook.js')
    ], {
      JSON,
      Date,
      console
    });

    const headerIndexes = evaluateInContext(
      "mapSummaryExportHeaderIndexes_(['key', 'value', ' source', 'asOf', 'note'])",
      context
    );
    assert.equal(headerIndexes.key, 0);
    assert.equal(headerIndexes.value, 1);
    assert.equal(headerIndexes.source, 2);
    assert.equal(headerIndexes.asOf, 3);
    assert.equal(headerIndexes.note, 4);

    const entries = evaluateInContext(`
      buildOkxDcaSummaryExportEntries_({
        method: 'spot_fills_incremental',
        timestamp: '2026-07-01T16:13:58.757Z',
        rawMessage: 'method=spot_fills_incremental; mode=incremental; newRows=0',
        preview: { buyCount: 1535 },
        summary: {
          totalBoughtBtc: 0.05075447,
          totalInvestedUsdt: 3536.4779348869974,
          derivedAvgPrice: 69678.15711378698,
          lastFillBillId: '3705053633186324480',
          lastFillTime: 1782921602838,
          syncMode: 'incremental',
          incrementalBuyCount: 10
        }
      })
    `, context);

    assert.equal(entries.length, 8);
    assert.equal(entries[0].key, 'OKX_BTC_DCA_BuyCount');
    assert.equal(entries[0].value, 1535);
    assert.equal(entries[1].key, 'OKX_BTC_DCA_IncrementalBuyCount');
    assert.equal(entries[1].value, 10);
    assert.equal(entries[2].key, 'OKX_BTC_DCA_TotalBought_BTC');
    assert.equal(entries[2].value, 0.05075447);
    assert.equal(entries[7].key, 'OKX_BTC_DCA_SyncMode');
    assert.equal(entries[7].value, 'incremental');
  });
}
