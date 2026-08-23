if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadWebhookContext(settingsStore = {}) {
    const repoRoot = path.resolve(__dirname, '..');
    const source = fs.readFileSync(path.join(repoRoot, 'Event_Webhook.js'), 'utf8');
    const configSource = fs.readFileSync(path.join(repoRoot, 'Config.js'), 'utf8');
    const store = { ...settingsStore };

    const sandbox = {
      JSON,
      Date,
      Math,
      console,
      Logger: { log() {} },
      LogService: { info() {}, warn() {}, error() {} },
      Settings: {
        get: (k) => store[k] || '',
        set: (k, v) => { store[k] = v; }
      },
      ContentService: {
        createTextOutput: (content) => ({
          getContent: () => content,
          parsed: () => JSON.parse(content)
        })
      }
    };

    const context = vm.createContext(sandbox);
    vm.runInContext(configSource, context, { filename: 'Config.js' });
    vm.runInContext(source, context, { filename: 'Event_Webhook.js' });
    context._store = store;
    return context;
  }

  test('doPost returns error when postData is missing', () => {
    const context = loadWebhookContext();
    const output = context.doPost({});
    assert.equal(output.parsed().status, 'error');
    assert.match(output.parsed().msg, /No Post Data/);
  });

  test('doPost rejects unauthorized request when PROXY_PASSWORD is set', () => {
    const context = loadWebhookContext({ PROXY_PASSWORD: 'correct_secret' });
    const e = {
      postData: {
        contents: JSON.stringify({ action: 'get_inventory', password: 'wrong_secret' })
      }
    };
    const output = context.doPost(e);
    assert.equal(output.parsed().status, 'error');
    assert.match(output.parsed().msg, /Auth Failed/);
  });

  test('doPost handles update_tunnel_url action correctly', () => {
    const context = loadWebhookContext({ PROXY_PASSWORD: 'secret' });
    const e = {
      postData: {
        contents: JSON.stringify({ action: 'update_tunnel_url', password: 'secret', url: 'https://new-tunnel.trycloudflare.com' })
      }
    };
    const output = context.doPost(e);
    assert.equal(output.parsed().status, 'success');
    assert.equal(context._store['TUNNEL_URL'], 'https://new-tunnel.trycloudflare.com');
  });

  test('doPost returns unknown action for unrecognized commands', () => {
    const context = loadWebhookContext({ PROXY_PASSWORD: 'secret' });
    const e = {
      postData: {
        contents: JSON.stringify({ action: 'non_existent_command', password: 'secret' })
      }
    };
    const output = context.doPost(e);
    assert.equal(output.parsed().status, 'error');
    assert.match(output.parsed().msg, /Unknown Action/);
  });
}
