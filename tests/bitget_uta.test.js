const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadBitget() {
  const file = path.resolve(__dirname, '..', 'Sync_Bitget.js');
  const context = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
  return context;
}

test('Bitget account reader uses UTA v3 assets', () => {
  const context = loadBitget();
  const calls = [];
  context.fetchBitgetApi_ = function (baseUrl, endpoint, params) {
    calls.push({ baseUrl, endpoint, params });
    return {
      code: '00000',
      data: {
        assets: [{ coin: 'USDT', equity: '12.5', balance: '12.5' }]
      }
    };
  };

  const result = context.fetchBitgetSpotAssets_('https://api.bitget.com', 'key', 'secret', 'pass');

  assert.equal(result.success, true);
  assert.equal(result.accountMode, 'UTA');
  assert.equal(result.data[0].equity, '12.5');
  assert.equal(calls.length, 1);
  assert.equal(calls[0].endpoint, '/api/v3/account/assets');
});

test('Bitget Earn reader uses account aggregate instead of Classic savings orders', () => {
  const context = loadBitget();
  let calledEndpoint = '';
  context.fetchBitgetApi_ = function (baseUrl, endpoint) {
    calledEndpoint = endpoint;
    return {
      code: '00000',
      data: [{ coin: 'USDT', amount: '400' }]
    };
  };

  const result = context.fetchBitgetEarnAssets_('https://api.bitget.com', 'key', 'secret', 'pass');

  assert.equal(result.success, true);
  assert.equal(result.data[0].amount, '400');
  assert.equal(calledEndpoint, '/api/v2/earn/account/assets');
  assert.doesNotMatch(calledEndpoint, /savings\/assets/);
});

test('Bitget Earn reader identifies UTA unsupported response', () => {
  const context = loadBitget();
  context.fetchBitgetApi_ = function () {
    return {
      code: '40085',
      msg: 'You are in Unified Account mode, and the Classic Account API is not supported at this time'
    };
  };

  const result = context.fetchBitgetEarnAssets_('https://api.bitget.com', 'key', 'secret', 'pass');

  assert.equal(result.success, false);
  assert.equal(result.unsupportedInUta, true);
  assert.equal(result.code, '40085');
});

test('Bitget Earn carry-forward preserves the original source timestamp', () => {
  const context = loadBitget();
  context.UnifiedAssetsRepo = {
    readAllRows() {
      return [
        ['Bitget', 'USDT', 100, 'Earn', 'Flexible', 'apy=5', '2026-07-01T12:00:00'],
        ['Bitget', 'BTC', 0.1, 'Unified', 'Equity', '', '2026-07-28T12:00:00'],
        ['OKX', 'USDT', 50, 'Earn', 'Flexible', '', '2026-07-28T12:00:00']
      ];
    }
  };

  const first = context.carryForwardBitgetEarnAssets_({});
  assert.equal(first.length, 1);
  assert.equal(first[0].ccy, 'USDT');
  assert.equal(first[0].status, 'Stale');
  assert.match(first[0].meta, /sourceUpdated=2026-07-01T12:00:00/);
  assert.match(first[0].meta, /warning=UTA Earn API unavailable/);

  context.UnifiedAssetsRepo.readAllRows = function () {
    return [['Bitget', 'USDT', 100, 'Earn', 'Stale', first[0].meta, '2026-07-29T12:00:00']];
  };
  const second = context.carryForwardBitgetEarnAssets_({});
  assert.match(second[0].meta, /sourceUpdated=2026-07-01T12:00:00/);
  assert.doesNotMatch(second[0].meta, /sourceUpdated=2026-07-29T12:00:00/);
});

test('Bitget funding reader uses UTA v3 funding assets', () => {
  const context = loadBitget();
  let calledEndpoint = '';
  context.fetchBitgetApi_ = function (baseUrl, endpoint) {
    calledEndpoint = endpoint;
    return {
      code: '00000',
      data: [{ coin: 'USDT', available: '1', frozen: '0', balance: '1' }]
    };
  };

  const result = context.fetchBitgetFundingAssets_('https://api.bitget.com', 'key', 'secret', 'pass');

  assert.equal(result.success, true);
  assert.equal(calledEndpoint, '/api/v3/account/funding-assets');
});
