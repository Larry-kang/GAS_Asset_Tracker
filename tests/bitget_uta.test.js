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

