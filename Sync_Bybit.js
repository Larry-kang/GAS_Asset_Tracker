// =======================================================
// --- Bybit Service Layer (v1.0 - Unified Ledger)
// --- Phase 1 MVP: Wallet, Funding, Positions, Earn
// =======================================================

function getBybitBalance() {
  const MODULE_NAME = "Sync_Bybit";

  return SyncManager.run(MODULE_NAME, () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = SyncManager.createResult("Bybit");
    const creds = Credentials.get('BYBIT');
    const { apiKey, apiSecret, bridgeV2Url, bridgeV2Password } = creds;
    const resolvedBridgeUrl = normalizeBybitBridgeUrl_(bridgeV2Url);
    const useBridge = !!resolvedBridgeUrl;
    const baseUrl = useBridge ? resolvedBridgeUrl : 'https://api.bybit.com';
    const proxyPassword = useBridge ? bridgeV2Password : '';

    if (!Credentials.isValid(creds)) {
      SyncManager.registerSourceCheck(result, {
        name: 'Credentials',
        required: true,
        success: false,
        message: 'Missing BYBIT_API_KEY / BYBIT_API_SECRET'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }

    if (useBridge && !proxyPassword) {
      SyncManager.registerSourceCheck(result, {
        name: 'Bridge V2',
        required: true,
        success: false,
        message: 'Missing BRIDGE_V2_PASSWORD'
      });
      return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
    }

    SyncManager.registerSourceCheck(result, {
      name: 'Transport',
      required: false,
      success: true,
      rows: 1,
      message: useBridge ? 'Using local-bridge V2' : 'Using direct Bybit API'
    });

    const accountInfoRes = fetchBybitAccountInfo_(baseUrl, apiKey, apiSecret, proxyPassword);
    const accountInfo = accountInfoRes.success ? (accountInfoRes.data || {}) : null;
    SyncManager.registerSourceCheck(result, {
      name: 'Account Info',
      required: false,
      success: accountInfoRes.success,
      rows: accountInfoRes.success ? 1 : 0,
      message: accountInfoRes.success ? '' : (accountInfoRes.status || 'Unknown error')
    });

    const instrumentRes = fetchBybitDerivativeInstrumentMap_(baseUrl);
    const instrumentMap = instrumentRes.data || { linear: createBybitInstrumentBucket_(), inverse: createBybitInstrumentBucket_() };
    SyncManager.registerSourceCheck(result, {
      name: 'Derivative Instruments',
      required: false,
      success: instrumentRes.success,
      rows: instrumentRes.success
        ? (Object.keys(instrumentMap.linear.bySymbol).length + Object.keys(instrumentMap.inverse.bySymbol).length)
        : 0,
      message: instrumentRes.success ? '' : (instrumentRes.status || 'Unknown error')
    });

    const earnProductRes = fetchBybitEarnProductMap_(baseUrl);
    const earnProductMap = earnProductRes.data || {};
    SyncManager.registerSourceCheck(result, {
      name: 'Earn Products',
      required: false,
      success: earnProductRes.success,
      rows: earnProductRes.success ? Object.keys(earnProductMap).length : 0,
      message: earnProductRes.success ? '' : (earnProductRes.status || 'Unknown error')
    });

    // A. Unified Wallet
    const walletRes = fetchBybitWalletBalance_(baseUrl, apiKey, apiSecret, proxyPassword);
    let walletRows = 0;
    if (walletRes.success && walletRes.data) {
      (walletRes.data.coins || []).forEach(item => {
        const coin = bybitUpper_(item.coin);
        const walletBalance = parseBybitNumber_(item.walletBalance);
        const locked = parseBybitNumber_(item.locked);
        const available = Math.max(walletBalance - locked, 0);
        const borrowAmount = Math.max(parseBybitNumber_(item.borrowAmount), parseBybitNumber_(item.spotBorrow));
        const accruedInterest = parseBybitNumber_(item.accruedInterest);
        const totalDebt = borrowAmount + accruedInterest;

        if (!coin) return;

        if (available > 0) {
          result.assets.push({
            ccy: coin,
            amt: available,
            type: 'Spot',
            status: 'Available',
            meta: buildBybitWalletMeta_(item)
          });
          walletRows++;
        }

        if (locked > 0) {
          result.assets.push({
            ccy: coin,
            amt: locked,
            type: 'Spot',
            status: 'Frozen',
            meta: buildBybitWalletMeta_(item)
          });
          walletRows++;
        }

        if (totalDebt > 0) {
          result.assets.push({
            ccy: coin,
            amt: -Math.abs(totalDebt),
            type: 'Loan',
            status: 'Debt',
            meta: buildBybitDebtMeta_(item)
          });
          walletRows++;
        }
      });
      SyncManager.registerSourceCheck(result, {
        name: 'Unified Wallet',
        required: true,
        success: true,
        rows: walletRows
      });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Unified Wallet',
        required: true,
        success: false,
        message: walletRes.status || 'Unknown error'
      });
    }

    // B. Funding Account
    const fundingRes = fetchBybitFundingBalances_(baseUrl, apiKey, apiSecret, proxyPassword);
    let fundingRows = 0;
    if (fundingRes.success && fundingRes.data) {
      fundingRes.data.forEach(item => {
        const coin = bybitUpper_(item.coin);
        const walletBalance = parseBybitNumber_(item.walletBalance);
        const transferBalance = parseBybitNumber_(item.transferBalance);
        const hasTransferField = item.transferBalance !== null
          && item.transferBalance !== undefined
          && String(item.transferBalance).trim() !== '';
        const available = hasTransferField ? transferBalance : walletBalance;
        const frozen = Math.max(walletBalance - available, 0);

        if (!coin) return;

        if (available > 0) {
          result.assets.push({
            ccy: coin,
            amt: available,
            type: 'Funding',
            status: 'Available',
            meta: buildBybitFundingMeta_(item)
          });
          fundingRows++;
        }
        if (frozen > 0) {
          result.assets.push({
            ccy: coin,
            amt: frozen,
            type: 'Funding',
            status: 'Frozen',
            meta: buildBybitFundingMeta_(item)
          });
          fundingRows++;
        }
      });
      SyncManager.registerSourceCheck(result, {
        name: 'Funding',
        required: false,
        success: true,
        rows: fundingRows
      });
    } else {
      SyncManager.registerSourceCheck(result, {
        name: 'Funding',
        required: false,
        success: false,
        message: fundingRes.status || 'Unknown error'
      });
    }

    // C. Positions
    const linearPositionsRes = fetchBybitPositionsForCategory_(baseUrl, apiKey, apiSecret, proxyPassword, 'linear', instrumentMap.linear);
    const inversePositionsRes = fetchBybitPositionsForCategory_(baseUrl, apiKey, apiSecret, proxyPassword, 'inverse', instrumentMap.inverse);
    let positionRows = 0;
    if (linearPositionsRes.success && inversePositionsRes.success) {
      linearPositionsRes.data.concat(inversePositionsRes.data).forEach(item => {
        result.assets.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Positions',
          status: 'Position',
          meta: buildBybitPositionMeta_(item)
        });
        positionRows++;
      });
      SyncManager.registerSourceCheck(result, {
        name: 'Positions',
        required: false,
        success: true,
        rows: positionRows
      });
    } else {
      const positionErrors = [];
      if (!linearPositionsRes.success) positionErrors.push(`linear: ${linearPositionsRes.status || 'Unknown error'}`);
      if (!inversePositionsRes.success) positionErrors.push(`inverse: ${inversePositionsRes.status || 'Unknown error'}`);
      SyncManager.registerSourceCheck(result, {
        name: 'Positions',
        required: false,
        success: false,
        message: positionErrors.join(' | ')
      });
    }

    // D. Earn Positions
    const flexibleEarnRes = fetchBybitEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword, 'FlexibleSaving');
    const onChainEarnRes = fetchBybitEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword, 'OnChain');
    let earnRows = 0;
    if (flexibleEarnRes.success && onChainEarnRes.success) {
      flexibleEarnRes.data.concat(onChainEarnRes.data).forEach(item => {
        result.assets.push({
          ccy: item.ccy,
          amt: item.amt,
          type: 'Earn',
          status: item.status,
          meta: buildBybitEarnMeta_(item, earnProductMap)
        });
        earnRows++;
      });
      SyncManager.registerSourceCheck(result, {
        name: 'Earn',
        required: false,
        success: true,
        rows: earnRows
      });
    } else {
      const earnErrors = [];
      if (!flexibleEarnRes.success) earnErrors.push(`FlexibleSaving: ${flexibleEarnRes.status || 'Unknown error'}`);
      if (!onChainEarnRes.success) earnErrors.push(`OnChain: ${onChainEarnRes.status || 'Unknown error'}`);
      SyncManager.registerSourceCheck(result, {
        name: 'Earn',
        required: false,
        success: false,
        message: earnErrors.join(' | ')
      });
    }

    if (accountInfo && !isBybitUnifiedAccount_(accountInfo)) {
      result.warnings.push(`Account mode may not be UTA. unifiedMarginStatus=${accountInfo.unifiedMarginStatus || 'unknown'}`);
    }

    return SyncManager.commitExchangeResult(ss, MODULE_NAME, result);
  });
}

function fetchBybitAccountInfo_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBybitPrivateApi_(baseUrl, '/v5/account/info', {}, apiKey, apiSecret, proxyPassword);
  if (!res.success) return res;
  return { success: true, data: res.data || {} };
}

function fetchBybitWalletBalance_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBybitPrivateApi_(baseUrl, '/v5/account/wallet-balance', {
    accountType: 'UNIFIED'
  }, apiKey, apiSecret, proxyPassword);
  if (!res.success) return res;

  const firstAccount = Array.isArray(res.data.list) ? res.data.list[0] : null;
  return {
    success: true,
    data: {
      summary: firstAccount || {},
      coins: firstAccount && Array.isArray(firstAccount.coin) ? firstAccount.coin : []
    }
  };
}

function fetchBybitFundingBalances_(baseUrl, apiKey, apiSecret, proxyPassword) {
  const res = fetchBybitPrivateApi_(baseUrl, '/v5/asset/transfer/query-account-coins-balance', {
    accountType: 'FUND'
  }, apiKey, apiSecret, proxyPassword);
  if (!res.success) return res;

  return {
    success: true,
    data: res.data && Array.isArray(res.data.balance) ? res.data.balance : []
  };
}

function fetchBybitPositionsForCategory_(baseUrl, apiKey, apiSecret, proxyPassword, category, instrumentBucket) {
  const bucket = instrumentBucket || createBybitInstrumentBucket_();
  const rawPositions = [];

  if (category === 'linear') {
    const settleCoins = bucket.settleCoins.length > 0 ? bucket.settleCoins : ['USDT', 'USDC'];
    settleCoins.forEach(settleCoin => {
      const res = fetchBybitPositionPageSet_(baseUrl, apiKey, apiSecret, proxyPassword, {
        category: 'linear',
        settleCoin: settleCoin
      });
      if (!res.success) {
        rawPositions.push({ __error: `settleCoin=${settleCoin}: ${res.status || 'Unknown error'}` });
        return;
      }
      rawPositions.push.apply(rawPositions, res.data || []);
    });
  } else {
    const res = fetchBybitPositionPageSet_(baseUrl, apiKey, apiSecret, proxyPassword, {
      category: 'inverse'
    });
    if (!res.success) return res;
    rawPositions.push.apply(rawPositions, res.data || []);
  }

  const errors = rawPositions.filter(item => item && item.__error).map(item => item.__error);
  if (errors.length > 0) {
    return { success: false, status: errors.join(' | ') };
  }

  return {
    success: true,
    data: rawPositions
      .filter(Boolean)
      .filter(item => {
        const size = parseBybitNumber_(item.size);
        return size > 0 && String(item.side || '') !== '';
      })
      .map(item => normalizeBybitPosition_(item, bucket.bySymbol))
  };
}

function fetchBybitPositionPageSet_(baseUrl, apiKey, apiSecret, proxyPassword, params) {
  const rows = [];
  let cursor = '';

  while (true) {
    const query = Object.assign({}, params, { limit: 200 });
    if (cursor) query.cursor = cursor;

    const res = fetchBybitPrivateApi_(baseUrl, '/v5/position/list', query, apiKey, apiSecret, proxyPassword);
    if (!res.success) return res;

    const data = res.data || {};
    const list = Array.isArray(data.list) ? data.list : [];
    rows.push.apply(rows, list);

    cursor = String(data.nextPageCursor || '').trim();
    if (!cursor) break;
  }

  return { success: true, data: rows };
}

function fetchBybitEarnPositions_(baseUrl, apiKey, apiSecret, proxyPassword, category) {
  const res = fetchBybitPrivateApi_(baseUrl, '/v5/earn/position', { category: category }, apiKey, apiSecret, proxyPassword);
  if (!res.success) return res;

  const list = res.data && Array.isArray(res.data.list) ? res.data.list : [];
  return {
    success: true,
    data: list.map(item => ({
      ccy: bybitUpper_(item.coin),
      amt: parseBybitNumber_(item.amount),
      status: category === 'FlexibleSaving' ? 'Flexible' : 'OnChain',
      category: category,
      productId: String(item.productId || ''),
      totalPnl: item.totalPnl,
      claimableYield: item.claimableYield,
      positionStatus: item.status,
      orderId: item.orderId,
      estimateRedeemTime: item.estimateRedeemTime,
      estimateStakeTime: item.estimateStakeTime,
      estimateInterestCalculationTime: item.estimateInterestCalculationTime,
      settlementTime: item.settlementTime,
      autoReinvest: item.autoReinvest
    }))
  };
}

function fetchBybitEarnProductMap_(baseUrl) {
  const categories = ['FlexibleSaving', 'OnChain'];
  const map = {};

  for (let i = 0; i < categories.length; i++) {
    const category = categories[i];
    const res = fetchBybitPublicApi_(baseUrl, '/v5/earn/product', { category: category });
    if (!res.success) return res;

    const list = res.data && Array.isArray(res.data.list) ? res.data.list : [];
    list.forEach(item => {
      const key = buildBybitEarnProductKey_(category, item.productId || item.coin);
      map[key] = item;
    });
  }

  return { success: true, data: map };
}

function fetchBybitDerivativeInstrumentMap_(baseUrl) {
  const linear = fetchBybitInstrumentBucket_(baseUrl, 'linear');
  if (!linear.success) return linear;

  const inverse = fetchBybitInstrumentBucket_(baseUrl, 'inverse');
  if (!inverse.success) return inverse;

  return {
    success: true,
    data: {
      linear: linear.data,
      inverse: inverse.data
    }
  };
}

function fetchBybitInstrumentBucket_(baseUrl, category) {
  const bucket = createBybitInstrumentBucket_();
  let cursor = '';

  while (true) {
    const query = {
      category: category,
      limit: 1000
    };
    if (cursor) query.cursor = cursor;

    const res = fetchBybitPublicApi_(baseUrl, '/v5/market/instruments-info', query);
    if (!res.success) return res;

    const data = res.data || {};
    const list = Array.isArray(data.list) ? data.list : [];
    list.forEach(item => {
      const symbol = String(item.symbol || '').toUpperCase();
      const settleCoin = bybitUpper_(item.settleCoin);
      if (symbol) bucket.bySymbol[symbol] = item;
      if (settleCoin && bucket.settleCoins.indexOf(settleCoin) === -1) {
        bucket.settleCoins.push(settleCoin);
      }
    });

    cursor = String(data.nextPageCursor || '').trim();
    if (!cursor) break;
  }

  return { success: true, data: bucket };
}

function fetchBybitPrivateApi_(baseUrl, endpoint, params, apiKey, apiSecret, proxyPassword) {
  const recvWindow = '10000';
  const timestamp = String(Date.now());
  const queryString = buildBybitQueryString_(params || {});
  const signPayload = `${timestamp}${apiKey}${recvWindow}${queryString}`;
  const signature = hexEncodeBytes_(Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    signPayload,
    apiSecret
  ));

  const url = queryString ? `${baseUrl}${endpoint}?${queryString}` : `${baseUrl}${endpoint}`;
  const options = {
    method: 'get',
    headers: {
      'X-BAPI-API-KEY': apiKey,
      'X-BAPI-TIMESTAMP': timestamp,
      'X-BAPI-RECV-WINDOW': recvWindow,
      'X-BAPI-SIGN': signature
    },
    muteHttpExceptions: true
  };

  if (proxyPassword) {
    options.headers['x-proxy-auth'] = proxyPassword;
  }

  return fetchBybitApiResponse_(url, options);
}

function fetchBybitPublicApi_(baseUrl, endpoint, params) {
  const queryString = buildBybitQueryString_(params || {});
  const url = queryString ? `${baseUrl}${endpoint}?${queryString}` : `${baseUrl}${endpoint}`;
  const options = {
    method: 'get',
    muteHttpExceptions: true
  };

  return fetchBybitApiResponse_(url, options);
}

function fetchBybitApiResponse_(url, options) {
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const contentText = response.getContentText();

    if (responseCode === 403) {
      return {
        success: false,
        status: `HTTP 403: Bybit API rejected this IP. Official docs note US/Mainland China IPs are restricted. ${truncateBybitMessage_(contentText)}`
      };
    }
    if (responseCode >= 400) {
      return {
        success: false,
        status: `HTTP ${responseCode}: ${truncateBybitMessage_(contentText)}`
      };
    }

    const json = safeParseBybitJson_(contentText);
    if (!json) {
      return {
        success: false,
        status: `Invalid JSON response: ${truncateBybitMessage_(contentText)}`
      };
    }
    if (!isBybitApiSuccess_(json)) {
      return {
        success: false,
        status: bybitStatus_(json)
      };
    }

    return {
      success: true,
      data: json.result || {}
    };
  } catch (e) {
    return {
      success: false,
      status: e.message
    };
  }
}

function buildBybitQueryString_(params) {
  return Object.keys(params || {})
    .filter(key => params[key] !== null && params[key] !== undefined && String(params[key]) !== '')
    .sort()
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
    .join('&');
}

function isBybitApiSuccess_(json) {
  if (!json) return false;
  return Number(json.retCode) === 0;
}

function bybitStatus_(json) {
  if (!json) return 'Unknown error';
  return `Code: ${json.retCode}, Msg: ${json.retMsg || 'Unknown error'}`;
}

function buildBybitWalletMeta_(item) {
  const parts = [];
  pushBybitNumericMetaIfNonZero_(parts, 'usdValue', item.usdValue);
  pushBybitNumericMetaIfNonZero_(parts, 'equity', item.equity);
  pushBybitNumericMetaIfNonZero_(parts, 'unrealisedPnl', item.unrealisedPnl);
  pushBybitNumericMetaIfNonZero_(parts, 'totalOrderIM', item.totalOrderIM);
  pushBybitNumericMetaIfNonZero_(parts, 'totalPositionIM', item.totalPositionIM);
  pushBybitNumericMetaIfNonZero_(parts, 'totalPositionMM', item.totalPositionMM);
  if (parseBybitBoolean_(item.marginCollateral)) parts.push('marginCollateral=true');
  if (parseBybitBoolean_(item.collateralSwitch)) parts.push('collateralSwitch=true');
  return parts.join('; ');
}

function buildBybitDebtMeta_(item) {
  const parts = [];
  pushBybitNumericMetaIfNonZero_(parts, 'borrowAmount', item.borrowAmount);
  pushBybitNumericMetaIfNonZero_(parts, 'spotBorrow', item.spotBorrow);
  pushBybitNumericMetaIfNonZero_(parts, 'accruedInterest', item.accruedInterest);
  pushBybitNumericMetaIfNonZero_(parts, 'usdValue', item.usdValue);
  return parts.join('; ');
}

function buildBybitFundingMeta_(item) {
  const parts = [];
  pushBybitNumericMetaIfNonZero_(parts, 'walletBalance', item.walletBalance);
  pushBybitNumericMetaIfNonZero_(parts, 'transferBalance', item.transferBalance);
  pushBybitNumericMetaIfNonZero_(parts, 'bonus', item.bonus);
  return parts.join('; ');
}

function normalizeBybitPosition_(item, instrumentLookup) {
  const instrument = instrumentLookup[String(item.symbol || '').toUpperCase()] || null;
  return {
    ccy: inferBybitPositionCurrency_(item.symbol, instrument),
    amt: parseBybitNumber_(item.size),
    symbol: item.symbol,
    side: item.side,
    category: item.category,
    positionIdx: item.positionIdx,
    avgPrice: item.avgPrice,
    markPrice: item.markPrice,
    positionValue: item.positionValue,
    leverage: item.leverage,
    liqPrice: item.liqPrice,
    positionIM: item.positionIM,
    positionMM: item.positionMM,
    unrealisedPnl: item.unrealisedPnl,
    curRealisedPnl: item.curRealisedPnl,
    cumRealisedPnl: item.cumRealisedPnl,
    positionStatus: item.positionStatus,
    adlRankIndicator: item.adlRankIndicator,
    updatedTime: item.updatedTime,
    settleCoin: instrument ? instrument.settleCoin : ''
  };
}

function buildBybitPositionMeta_(item) {
  const parts = [];
  pushBybitMetaIfPresent_(parts, 'category', item.category);
  pushBybitMetaIfPresent_(parts, 'symbol', item.symbol);
  pushBybitMetaIfPresent_(parts, 'side', item.side);
  pushBybitMetaIfPresent_(parts, 'positionStatus', item.positionStatus);
  pushBybitMetaIfPresent_(parts, 'settleCoin', item.settleCoin);
  pushBybitNumericMetaIfNonZero_(parts, 'avgPrice', item.avgPrice);
  pushBybitNumericMetaIfNonZero_(parts, 'markPrice', item.markPrice);
  pushBybitNumericMetaIfNonZero_(parts, 'positionValue', item.positionValue);
  pushBybitNumericMetaIfNonZero_(parts, 'unrealisedPnl', item.unrealisedPnl);
  pushBybitNumericMetaIfNonZero_(parts, 'leverage', item.leverage);
  pushBybitNumericMetaIfNonZero_(parts, 'liqPrice', item.liqPrice);
  pushBybitNumericMetaIfNonZero_(parts, 'positionIM', item.positionIM);
  pushBybitNumericMetaIfNonZero_(parts, 'positionMM', item.positionMM);
  pushBybitNumericMetaIfNonZero_(parts, 'adlRank', item.adlRankIndicator);
  pushBybitMetaIfPresent_(parts, 'updatedTime', formatBybitTimestamp_(item.updatedTime));
  return parts.join('; ');
}

function buildBybitEarnMeta_(item, earnProductMap) {
  const product = earnProductMap[buildBybitEarnProductKey_(item.category, item.productId || item.ccy)] || null;
  const parts = [];
  pushBybitMetaIfPresent_(parts, 'productId', item.productId);
  pushBybitMetaIfPresent_(parts, 'estimateApr', product ? product.estimateApr : '');
  pushBybitNumericMetaIfNonZero_(parts, 'totalPnl', item.totalPnl);
  pushBybitNumericMetaIfNonZero_(parts, 'claimableYield', item.claimableYield);
  pushBybitMetaIfPresent_(parts, 'status', item.positionStatus);
  pushBybitMetaIfPresent_(parts, 'autoReinvest', item.autoReinvest);
  pushBybitMetaIfPresent_(parts, 'duration', product ? product.duration : '');
  pushBybitMetaIfPresent_(parts, 'rewardDistributionType', product ? product.rewardDistributionType : '');
  pushBybitMetaIfPresent_(parts, 'settlementTime', formatBybitTimestamp_(item.settlementTime));
  return parts.join('; ');
}

function buildBybitEarnProductKey_(category, productIdOrCoin) {
  return `${String(category || '')}|${String(productIdOrCoin || '').toUpperCase()}`;
}

function inferBybitPositionCurrency_(symbol, instrument) {
  if (instrument && instrument.baseCoin) return bybitUpper_(instrument.baseCoin);

  const rawSymbol = String(symbol || '').toUpperCase();
  const commonSuffixes = ['USDT', 'USDC', 'USD', 'BTC', 'ETH'];
  for (let i = 0; i < commonSuffixes.length; i++) {
    const suffix = commonSuffixes[i];
    if (rawSymbol.endsWith(suffix) && rawSymbol.length > suffix.length) {
      return rawSymbol.slice(0, rawSymbol.length - suffix.length);
    }
  }
  return rawSymbol;
}

function createBybitInstrumentBucket_() {
  return {
    bySymbol: {},
    settleCoins: []
  };
}

function isBybitUnifiedAccount_(accountInfo) {
  const status = Number(accountInfo && accountInfo.unifiedMarginStatus);
  return isFinite(status) && status > 1;
}

function parseBybitNumber_(value) {
  const num = parseFloat(value);
  return isFinite(num) ? num : 0;
}

function parseBybitBoolean_(value) {
  if (value === true || value === 'true' || value === 'TRUE') return true;
  return false;
}

function bybitUpper_(value) {
  return String(value || '').trim().toUpperCase();
}

function pushBybitNumericMetaIfNonZero_(parts, label, value) {
  const num = parseFloat(value);
  if (!isFinite(num) || num === 0) return;
  parts.push(`${label}=${value}`);
}

function pushBybitMetaIfPresent_(parts, label, value) {
  const text = String(value || '').trim();
  if (!text) return;
  parts.push(`${label}=${text}`);
}

function formatBybitTimestamp_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';

  const date = new Date(Number(raw));
  if (isNaN(date.getTime())) return raw;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function truncateBybitMessage_(text) {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  return clean.length > 220 ? clean.slice(0, 217) + '...' : clean;
}

function normalizeBybitBridgeUrl_(value) {
  const text = String(value || '').trim();
  if (!/^https?:\/\//i.test(text)) return '';
  const normalized = text.replace(/\/+$/, '');
  if (/\/bybit$/i.test(normalized)) return normalized;
  return normalized + '/bybit';
}

function safeParseBybitJson_(text) {
  try {
    return text ? JSON.parse(text) : {};
  } catch (e) {
    return null;
  }
}

function hexEncodeBytes_(bytes) {
  return bytes.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}
