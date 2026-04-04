const ExchangeRegistry = {
  getActive: function () {
    return [
      {
        id: 'BINANCE',
        displayName: '幣安',
        moduleName: 'Binance',
        functionName: 'getBinanceBalance',
        requiredSettings: ['BINANCE_API_KEY', 'BINANCE_API_SECRET', 'TUNNEL_URL', 'PROXY_PASSWORD']
      },
      {
        id: 'OKX',
        displayName: 'OKX',
        moduleName: 'OKX',
        functionName: 'getOkxBalance',
        requiredSettings: ['OKX_API_KEY', 'OKX_API_SECRET', 'OKX_API_PASSPHRASE']
      },
      {
        id: 'BITGET',
        displayName: 'Bitget',
        moduleName: 'Bitget',
        functionName: 'getBitgetBalance',
        requiredSettings: ['BITGET_API_KEY', 'BITGET_API_SECRET', 'BITGET_API_PASSPHRASE']
      },
      {
        id: 'BYBIT',
        displayName: 'Bybit',
        moduleName: 'Bybit',
        functionName: 'getBybitBalance',
        requiredSettings: ['BYBIT_API_KEY', 'BYBIT_API_SECRET']
      },
      {
        id: 'PIONEX',
        displayName: '派網',
        moduleName: 'Pionex',
        functionName: 'getPionexBalance',
        requiredSettings: ['PIONEX_API_KEY', 'PIONEX_API_SECRET']
      },
      {
        id: 'BITOPRO',
        displayName: '幣託',
        moduleName: 'BitoPro',
        functionName: 'getBitoProBalance',
        requiredSettings: ['BITOPRO_API_KEY', 'BITOPRO_API_SECRET']
      }
    ];
  },

  findByFunctionName: function (functionName) {
    return this.getActive().filter(function (entry) {
      return entry.functionName === functionName;
    })[0] || null;
  },

  getCredentialStatus: function (entry) {
    const missing = (entry.requiredSettings || []).filter(function (key) {
      return !Settings.get(key);
    });

    return {
      ok: missing.length === 0,
      missing: missing
    };
  }
};
