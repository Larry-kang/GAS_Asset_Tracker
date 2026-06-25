if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadSnapshotContext() {
    const repoRoot = path.resolve(__dirname, '..');
    const source = fs.readFileSync(path.join(repoRoot, 'Record_DailySnapshot.js'), 'utf8');
    const sandbox = {
      JSON,
      Date,
      Math,
      console,
      LogService: { info() {}, warn() {}, error() {} }
    };
    const context = vm.createContext(sandbox);
    vm.runInContext(source, context, { filename: 'Record_DailySnapshot.js' });
    return context;
  }

  test('buildSnapshotSummaryFromAllocationRows_ matches chart parity rules', () => {
    const context = loadSnapshotContext();
    const rows = [
      ['TPE:', '股票', '00713', 10000, 'TWD', 48.42, 60.85, 484200, 608500],
      ['TPE:', '股票', '00662', 3000, 'TWD', 100.38, 122, 301140, 366000],
      ['TPE:', '股票', '00670L', 38, 'TWD', 131.18, 210, 4984.84, 7980],
      ['TPE:', '股票', 'YMAG', 26, 'USD', 14.672, 11.27, 12146.26, 9329.90331],
      ['TPE:', '股票', '0056', 142, 'TWD', 40.53, 53.2, 5755.26, 7554.4],
      ['TPE:', '股票', '00878', 238, 'TWD', 24.43, 33.75, 5814.34, 8032.5],
      ['NYSEARCA:', '股票', 'IBIT', 683.56061, 'USD', 53.37, 33.87, 1161593.33, 737177.5559836184],
      ['Binance', '數位幣', 'BTC', 0.10020935, 'USD', 95518.31, 61636, 304771.78, 196662.95958349234],
      ['OKX', '數位穩定幣', 'USDT', -7579.18906566, 'USD', 1, 0.998602, -241325.16, -240987.79685826294],
      ['聯邦銀行', '現金', 'TWD', 160457, 'TWD', '', 160457, 160457, 160457],
      ['永豐銀行', '現金', 'TWD', 12861.9898644, 'TWD', '', 12861.9898644, 12861.9898644, 12861.9898644],
      ['貸款來源', '貸款', 'Credit_Loan', 0, 'TWD', '', '', 0, -1020200],
      ['質押來源', '質押', 'Pledge_Loan', 0, 'TWD', '', '', 0, -309511.3698630137],
      ['卡費來源', '信用卡費', 'Credit_Card_Fees', 0, 'TWD', '', '', 0, -24796]
    ];

    const summary = vm.runInContext(
      `buildSnapshotSummaryFromAllocationRows_(${JSON.stringify(rows)}, 1511.369863013699)`,
      context
    );

    assert.ok(Math.abs(summary.cashValue - 173318.9898644) < 1e-9);
    assert.ok(Math.abs(summary.stockValue - 1007396.80331) < 1e-9);
    assert.ok(Math.abs(summary.cryptoValue - 933840.5155671107) < 1e-9);
    assert.ok(Math.abs(summary.cryptoliabilityValue - (-240987.79685826294)) < 1e-9);
    assert.ok(Math.abs(summary.liabilityValue - (-1356018.7397260275)) < 1e-9);
    assert.ok(Math.abs(summary.totalAssetValue - 2114556.308741511) < 1e-9);
    assert.ok(Math.abs(summary.netWorthValue - 517549.7721572205) < 1e-9);
  });
}
