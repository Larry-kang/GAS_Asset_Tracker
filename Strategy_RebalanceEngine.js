/**
 * Strategy_RebalanceEngine.js
 * Sovereign Asset Protocol - Layer Portfolio Aggregation & Rebalance Targets
 */


function calculateAutoPledgeRatios(rawPortfolio, indicatorsRaw) {
  const labelMap = {};
  rawPortfolio.forEach(item => {
    let label = item.purpose ? item.purpose.trim() : "";

    // [V24.11 Refined] Normalization: Binance_Pledge -> Binance, Stock_Pledge -> Stock
    // This allows exact matching with indicators and rules.
    if (label.indexOf("_Pledge") > -1) {
      label = label.split("_")[0];
    }

    if (!label || label.toLowerCase() === "none") return;
    if (!labelMap[label]) { labelMap[label] = { assets: 0, debt: 0 }; }
    if (item.value > 0) { labelMap[label].assets += item.value; }
    else { labelMap[label].debt += Math.abs(item.value); }
  });

  const groups = [];
  Object.keys(labelMap).forEach(name => {
    const data = labelMap[name];
    if (data.debt > 0) {
      const ratio = data.assets / data.debt;
      const lowerName = name.toLowerCase();

      // [V24.11 Refined] Dynamic Threshold Mapping
      // Matches indicators like: Stock_Pledge_Maint_Alert, Binance_Pledge_Maint_Alert
      const alertKey = name + "_Pledge_Maint_Alert";
      const criticalKey = name + "_Pledge_Maint_Critical";

      const isCrypto = lowerName !== "stock";
      const defaultAlert = isCrypto ? Config.STRATEGIC.CRYPTO_LOAN_RATIO_ALERT : Config.STRATEGIC.PLEDGE_RATIO_ALERT;
      const defaultCritical = isCrypto ? Config.STRATEGIC.CRYPTO_LOAN_RATIO_CRITICAL : Config.STRATEGIC.PLEDGE_RATIO_CRITICAL;

      const alertThreshold = indicatorsRaw[alertKey] || defaultAlert;
      const criticalThreshold = indicatorsRaw[criticalKey] || defaultCritical;

      groups.push({
        name: name,
        ratio: ratio,
        collateralValue: data.assets,
        loanAmount: data.debt,
        alert: alertThreshold,
        critical: criticalThreshold
      });
    }
  });
  return groups;
}

function isActiveCryptoPledgeGroup_(group) {
  const name = String((group && group.name) || '').toLowerCase();
  if (!name) return false;
  if (name.indexOf('stock') >= 0) return false;
  return (parseFloat(group.loanAmount) || 0) > 0;
}

function aggregatePortfolio(rawPortfolio) {
  const summary = {};
  rawPortfolio.forEach(item => {
    if (!summary[item.ticker]) summary[item.ticker] = 0;
    summary[item.ticker] += item.value;
  });
  return summary;
}

function calculateGroupValue(summary, group) {
  let value = 0;
  group.tickers.forEach(t => value += (summary[t] || 0));
  return value;
}

function getRebalanceTargets(portfolio, assets, market) {
  const targets = [];
  const groups = Array.isArray(portfolio) ? portfolio : [];
  if (assets <= 0 || groups.length === 0) return targets;

  const threshold = Config.STRATEGIC.REBALANCE_ABS || 0.03;
  const coreStates = groups
    .filter(group => !group.isMisc)
    .map(group => buildRebalanceGroupState_(group, assets));
  const overweightStates = coreStates
    .filter(state => state.weightGap < -threshold)
    .sort((a, b) => Math.abs(b.deltaValue) - Math.abs(a.deltaValue));
  const underweightStates = coreStates
    .filter(state => state.weightGap > threshold)
    .sort((a, b) => Math.abs(b.deltaValue) - Math.abs(a.deltaValue));

  groups.forEach(group => {
    const state = buildRebalanceGroupState_(group, assets);

    if (group.isMisc) {
      if (state.currentValue > 0 && state.currentWeight > state.targetWeight) {
        targets.push({
          id: group.id,
          name: group.name,
          action: 'CLEAR',
          currentValue: state.currentValue,
          targetValue: state.targetValue,
          deltaValue: state.deltaValue,
          currentWeight: state.currentWeight,
          targetWeight: state.targetWeight,
          weightGap: state.weightGap,
          priority: 1,
          rationale: buildRebalanceRationale_(group, 'CLEAR', state),
          suggestedFundingSource: suggestRebalanceFundingSource_(group, 'CLEAR', market, overweightStates, underweightStates),
          executionHint: buildRebalanceExecutionHint_(group, 'CLEAR', market, overweightStates, underweightStates),
          tickers: group.tickers || []
        });
      }
      return;
    }

    if (Math.abs(state.weightGap) < threshold) return;

    const action = state.deltaValue > 0 ? 'ADD' : 'TRIM';
    if (action === 'ADD' && group.id === 'L1' && shouldSuppressAggressiveBtcBuying_(market && market.btcRegime)) {
      return;
    }

    targets.push({
      id: group.id,
      name: group.name,
      action: action,
      currentValue: state.currentValue,
      targetValue: state.targetValue,
      deltaValue: state.deltaValue,
      currentWeight: state.currentWeight,
      targetWeight: state.targetWeight,
      weightGap: state.weightGap,
      priority: getRebalancePriority_(group, action, market),
      rationale: buildRebalanceRationale_(group, action, state),
      suggestedFundingSource: suggestRebalanceFundingSource_(group, action, market, overweightStates, underweightStates),
      executionHint: buildRebalanceExecutionHint_(group, action, market, overweightStates, underweightStates),
      tickers: group.tickers || []
    });
  });

  return targets.sort(function (a, b) {
    if (a.priority !== b.priority) return a.priority - b.priority;
    return Math.abs(b.deltaValue) - Math.abs(a.deltaValue);
  });
}

function buildRebalanceGroupState_(group, assets) {
  const currentValue = parseFloat(group && group.value) || 0;
  const targetWeight = parseFloat(group && (group.target !== undefined ? group.target : group.defaultTarget)) || 0;
  const currentWeight = assets > 0 ? (currentValue / assets) : 0;
  const targetValue = assets * targetWeight;
  const deltaValue = targetValue - currentValue;
  const weightGap = targetWeight - currentWeight;

  return {
    id: group.id,
    name: group.name,
    currentValue: currentValue,
    targetWeight: targetWeight,
    currentWeight: currentWeight,
    targetValue: targetValue,
    deltaValue: deltaValue,
    weightGap: weightGap
  };
}

function getRebalanceShortName_(group) {
  return String((group && group.name) || (group && group.id) || '')
    .replace(/^Layer\s+\d+:\s*/i, '')
    .trim();
}

function getRebalancePriority_(group, action, market) {
  if (group && group.isMisc) return 1;
  if (action === 'TRIM' && group && group.id === 'L3') return 2;
  if (action === 'TRIM' && group && group.id === 'L2') return 3;
  if (action === 'TRIM') return 4;
  if (action === 'ADD' && group && group.id === 'L1') return 5;
  if (action === 'ADD' && group && group.id === 'L2') return 6;
  if (action === 'ADD' && group && group.id === 'L3') {
    return market && market.surplus < 0 ? 5 : 7;
  }
  if (action === 'ADD') return 8;
  return 9;
}

function buildRebalanceRationale_(group, action, state) {
  const gapPct = Math.abs(state.weightGap * 100).toFixed(1);
  const shortName = getRebalanceShortName_(group);
  if (action === 'ADD') {
    return `${shortName} 低於目標配置 ${gapPct}%`;
  }
  if (action === 'TRIM') {
    return `${shortName} 高於目標配置 ${gapPct}%`;
  }
  if (action === 'CLEAR') {
    if (state.targetWeight > 0) {
      return `${shortName} 超出容許配置 ${gapPct}%`;
    }
    return `${shortName} 需要清理`;
  }
  return `${shortName} 需要清理`;
}

function suggestRebalanceFundingSource_(group, action, market, overweightStates, underweightStates) {
  const preferredDestination = getPreferredRebalanceDestination_(group, underweightStates);
  const preferredSource = getPreferredRebalanceSource_(group, overweightStates);

  if (action === 'CLEAR') {
    if (preferredDestination) {
      return `優先回補 ${preferredDestination}`;
    }
    return '優先降低負債或回補低配層';
  }

  if (action === 'TRIM') {
    if (preferredDestination) {
      return `優先轉入 ${preferredDestination}`;
    }
    return '轉入低配層或增加流動性';
  }

  if (preferredSource) {
    return `先釋放 ${preferredSource}`;
  }

  if (market && market.surplus > 0) {
    return '新增盈餘 / 現金流';
  }

  return 'L3 / 閒置現金';
}

function buildRebalanceExecutionHint_(group, action, market, overweightStates, underweightStates) {
  const groupId = String((group && group.id) || '');
  const preferredDestination = getPreferredRebalanceDestination_(group, underweightStates);
  const preferredSource = getPreferredRebalanceSource_(group, overweightStates);

  if (action === 'CLEAR') {
    if (group && group.tickers && group.tickers.length > 0) {
      const destinationText = preferredDestination ? `；回收資金優先轉入 ${preferredDestination}` : '；回收資金優先降低負債或補低配層';
      return `優先處理: ${group.tickers.join(', ')}${destinationText}`;
    }
    return '優先清理雜項資產';
  }

  if (groupId === 'L1') {
    return action === 'ADD'
      ? (preferredSource
        ? `建議先釋放 ${preferredSource}，再補強 BTC / IBIT 核心持倉`
        : '以新增資金或減碼防禦倉補強 BTC / IBIT 核心持倉')
      : '可逐步把過重的 BTC / IBIT 轉入 L2 或 L3';
  }

  if (groupId === 'L2') {
    return action === 'ADD'
      ? (preferredSource
        ? `先釋放 ${preferredSource}，再依股票名目曝險與區域缺口補強`
        : '依股票名目曝險與台灣/NASDAQ 區域缺口補強')
      : (preferredDestination
        ? `視評價與比重調節，將資金優先回流至 ${preferredDestination}`
        : '視評價與比重調節，將資金回流至 L1 或 L3');
  }

  if (groupId === 'L3') {
    return action === 'ADD'
      ? '提高法幣、穩定幣或 BOXX 流動性'
      : (preferredDestination
        ? `釋放部分流動性，優先回補 ${preferredDestination}`
        : '釋放部分流動性回補低配核心層');
  }

  return action === 'ADD'
    ? '補強目前低配層'
    : '調節目前高配層';
}

function getPreferredRebalanceSource_(group, overweightStates) {
  const groupId = String((group && group.id) || '');
  const orderedIds = ['L4', 'L3', 'L2'];
  const states = Array.isArray(overweightStates) ? overweightStates : [];

  for (let i = 0; i < orderedIds.length; i++) {
    const id = orderedIds[i];
    if (id === groupId) continue;
    const match = states.find(function (state) {
      return state.id === id;
    });
    if (match) {
      return getRebalanceShortName_(match);
    }
  }

  return null;
}

function getPreferredRebalanceDestination_(group, underweightStates) {
  const groupId = String((group && group.id) || '');
  const orderedIds = ['L1', 'L2', 'L3'];
  const states = Array.isArray(underweightStates) ? underweightStates : [];

  for (let i = 0; i < orderedIds.length; i++) {
    const id = orderedIds[i];
    if (id === groupId) continue;
    const match = states.find(function (state) {
      return state.id === id;
    });
    if (match) {
      return getRebalanceShortName_(match);
    }
  }

  return null;
}

function buildRebalanceSequenceSummary_(targets) {
  const sequence = [];
  const list = Array.isArray(targets) ? targets : [];

  const clearTarget = list.find(function (target) { return target.action === 'CLEAR'; });
  if (clearTarget) sequence.push(`先清理 ${getRebalanceShortName_(clearTarget)}`);

  list
    .filter(function (target) { return target.action === 'TRIM'; })
    .slice(0, 2)
    .forEach(function (target) {
      sequence.push(`再調節 ${getRebalanceShortName_(target)}`);
    });

  const addTarget = list.find(function (target) { return target.action === 'ADD'; });
  if (addTarget) sequence.push(`最後補強 ${getRebalanceShortName_(addTarget)}`);

  return sequence;
}

function enrichRebalanceTargets_(targets, portfolioSummary, assetGroups) {
  const list = Array.isArray(targets) ? targets : [];
  return list.map(function (target) {
    const productHints = buildRebalanceProductHints_(target, portfolioSummary, assetGroups);
    return Object.assign({}, target, { productHints: productHints });
  });
}

function buildRebalanceProductHints_(target, portfolioSummary, assetGroups) {
  const ranked = rankRebalanceProducts_(target, portfolioSummary, assetGroups);
  const primary = ranked.slice(0, 3);
  const secondary = ranked.slice(3);

  return {
    primary: primary,
    secondary: secondary,
    rationale: buildRebalanceProductRationale_(target, ranked)
  };
}

function rankRebalanceProducts_(target, portfolioSummary, assetGroups) {
  const summary = portfolioSummary || {};
  const group = resolveRebalanceTargetGroup_(target, assetGroups);
  const tickers = getTargetTickers_(target, group);

  if (target && target.action === 'CLEAR') {
    return tickers
      .map(function (ticker) {
        return {
          ticker: ticker,
          value: parseFloat(summary[ticker]) || 0
        };
      })
      .filter(function (item) { return item.value > 0; })
      .sort(function (a, b) {
        if (b.value !== a.value) return b.value - a.value;
        return a.ticker.localeCompare(b.ticker);
      })
      .map(function (item) { return item.ticker; });
  }

  const positiveTickerSet = new Set(
    tickers.filter(function (ticker) {
      return (parseFloat(summary[ticker]) || 0) > 0;
    })
  );
  const configuredPriority = getConfiguredProductPriority_(target);
  const ordered = [];
  const seen = new Set();

  configuredPriority.forEach(function (ticker) {
    if (!tickers.includes(ticker)) return;
    if (target && target.action === 'TRIM' && !positiveTickerSet.has(ticker)) return;
    if (!seen.has(ticker)) {
      ordered.push(ticker);
      seen.add(ticker);
    }
  });

  const remaining = tickers
    .filter(function (ticker) {
      if (seen.has(ticker)) return false;
      if (target && target.action === 'TRIM') return positiveTickerSet.has(ticker);
      return true;
    })
    .sort(function (a, b) {
      const aValue = parseFloat(summary[a]) || 0;
      const bValue = parseFloat(summary[b]) || 0;
      if (bValue !== aValue) return bValue - aValue;
      return a.localeCompare(b);
    });

  remaining.forEach(function (ticker) {
    if (!seen.has(ticker)) {
      ordered.push(ticker);
      seen.add(ticker);
    }
  });

  return ordered;
}

function getConfiguredProductPriority_(target) {
  const strategic = (Config && Config.STRATEGIC) || {};
  const priorityMap = strategic.PRODUCT_PRIORITY || {};
  const key = `${String((target && target.id) || '').toUpperCase()}_${String((target && target.action) || '').toUpperCase()}`;
  return Array.isArray(priorityMap[key]) ? priorityMap[key].slice() : [];
}

function resolveRebalanceTargetGroup_(target, assetGroups) {
  const groups = Array.isArray(assetGroups) ? assetGroups : [];
  return groups.find(function (group) {
    return group && target && group.id === target.id;
  }) || null;
}

function getTargetTickers_(target, group) {
  const source = (target && Array.isArray(target.tickers) && target.tickers.length > 0)
    ? target.tickers
    : (group && Array.isArray(group.tickers) ? group.tickers : []);
  return source.filter(function (ticker, index, list) {
    return ticker && list.indexOf(ticker) === index;
  });
}

function buildRebalanceProductRationale_(target, rankedTickers) {
  const tickers = Array.isArray(rankedTickers) ? rankedTickers : [];
  if (tickers.length === 0) return '';

  if (target && target.action === 'CLEAR') {
    return '依部位大小排序，優先處理最能釋放資金的雜項資產';
  }

  if (target && target.id === 'L1' && target.action === 'ADD') {
    return '優先補強核心 BTC 曝險，其次才是代理曝險部位';
  }

  if (target && target.id === 'L2' && target.action === 'TRIM') {
    return '先減較偏成長或次核心的部位，再保留主要信用基底';
  }

  if (target && target.id === 'L3' && target.action === 'TRIM') {
    return '先釋放非最核心的流動性替代部位，最後才動用法幣現金';
  }

  if (target && target.id === 'L2' && target.action === 'ADD') {
    return '先補策略核心防禦倉，再增加成長型信用基底';
  }

  if (target && target.id === 'L3' && target.action === 'ADD') {
    return '先補最直接可用的現金流動性，再補其他流動性工具';
  }

  return '依目前策略優先序提供商品級建議';
}

function formatRebalanceProductHint_(target) {
  const hints = target && target.productHints;
  const primary = hints && Array.isArray(hints.primary) ? hints.primary.filter(Boolean) : [];
  if (primary.length === 0) return '';
  return primary.join(', ');
}

function buildRebalanceAlert_(target) {
  const shortName = getRebalanceShortName_(target);
  const currentPct = (target.currentWeight * 100).toFixed(1);
  const targetPct = (target.targetWeight * 100).toFixed(1);
  const deltaAbs = Math.round(Math.abs(target.deltaValue)).toLocaleString();
  const productHint = formatRebalanceProductHint_(target);

  if (target.action === 'CLEAR') {
    const allowedPct = (target.targetWeight * 100).toFixed(1);
    let actionText = `${target.executionHint || '優先清理雜項資產'}，並將資金回補低配層或降低槓桿。`;
    if (target.suggestedFundingSource) {
      actionText += `\n資金優先方向：${target.suggestedFundingSource}`;
    }
    if (productHint) {
      actionText += `\n優先標的：${productHint}`;
    }
    return {
      level: "[配置] 再平衡建議 (清理雜項)",
      message: `${shortName} 目前佔比 ${currentPct}%${target.targetWeight > 0 ? `，已超出容許 ${allowedPct}%` : '，已被歸類為待清理資產'}。`,
      action: actionText
    };
  }

  const actionLabel = target.action === 'ADD' ? '補強' : '減碼';
  let actionText = `建議${actionLabel} ${deltaAbs} TWD。`;
  if (target.suggestedFundingSource) {
    actionText += `\n資金流向/來源：${target.suggestedFundingSource}`;
  }
  if (target.executionHint) {
    actionText += `\n執行提示：${target.executionHint}`;
  }
  if (productHint) {
    actionText += `\n優先標的：${productHint}`;
  }

  return {
    level: target.action === 'ADD' ? "[配置] 再平衡建議 (補強)" : "[配置] 再平衡建議 (調節)",
    message: `${shortName} 目前 ${currentPct}% / 目標 ${targetPct}%`,
    action: actionText
  };
}
