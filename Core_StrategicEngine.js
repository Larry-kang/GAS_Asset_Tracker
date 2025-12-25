/**
 * @OnlyCurrentDoc
 * Sovereign Asset Protocol (SAP) v24.5 - Command Center Edition
 * Version: 24.5 (Bitcoin Standard / Treasury Edition)
 */

const CONFIG = {
  SHEET_NAMES: {
    BALANCE_SHEET: "Balance Sheet",
    INDICATORS: "Key Market Indicators"
  },
  get EMAIL_RECIPIENT() {
    const email = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (!email) {
      Logger.log("[WARNING] ScriptProperty 'ADMIN_EMAIL' not set.");
    }
    return email || "";
  },
  THRESHOLDS: {
    REBALANCE_ABS: 0.03,
    PLEDGE_RATIO_SAFE: 2.5,
    PLEDGE_RATIO_ALERT: 2.1,
    PLEDGE_RATIO_CRITICAL: 1.8,
    CRYPTO_LOAN_RATIO_SAFE: 2.0,
    CRYPTO_LOAN_RATIO_ALERT: 1.5,
    CRYPTO_LOAN_RATIO_CRITICAL: 1.3,
    get TREASURY_RESERVE_TWD() {
      const val = PropertiesService.getScriptProperties().getProperty('TREASURY_RESERVE_TWD');
      return val ? parseFloat(val) : 100000;
    }
  },
  BTC_MARTINGALE: {
    ENABLED: true,
    BASE_AMOUNT: 10000,
    LEVELS: [
      { drop: -0.30, multiplier: 1, name: "Level 1 (Probe)" },
      { drop: -0.40, multiplier: 2, name: "Level 2 (Add)" },
      { drop: -0.50, multiplier: 4, name: "Level 3 (Half-Price Sniper)" },
      { drop: -0.60, multiplier: 8, name: "Level 4 (Abyss Catcher)" },
      { drop: -0.70, multiplier: 16, name: "Level 5 (Capitulation Play)" }
    ]
  },
  ASSET_GROUPS: [
    { name: "Layer 1: Digital Reserve (Attack)", target: 0.80, tickers: ["IBIT", "BTC_Spot", "BTC"] },
    { name: "Layer 2: Credit Base (Defend)", target: 0.20, tickers: ["00713", "00662", "QQQ"] },
    { name: "Layer 3: Tactical Liquidity", target: 0.00, tickers: ["BOXX", "CASH_TWD"] }
  ],
  NOISE_ASSETS: ["ETH", "BNB", "TQQQ"]
};

const RULES = [
  {
    name: "Maintenance Ratio Monitor",
    phase: "All",
    condition: function (context) { return context.indicators.isValid && context.indicators.maintenanceRatio < CONFIG.THRESHOLDS.PLEDGE_RATIO_SAFE; },
    getAction: function (context) {
      const ratio = context.indicators.maintenanceRatio;
      if (ratio <= CONFIG.THRESHOLDS.PLEDGE_RATIO_CRITICAL) {
        return {
          level: "[嚴重] 斷頭追繳警報 (1.8)",
          message: "維持率崩跌至 " + ratio.toFixed(2),
          action: "執行焦土防禦: 強制清算所有雜訊資產 (ETH/BNB/TQQQ) 以償還債務。禁止買入。"
        };
      } else if (ratio <= CONFIG.THRESHOLDS.PLEDGE_RATIO_ALERT) {
        return {
          level: "[警告] 警戒區 (2.1)",
          message: "維持率降至 " + ratio.toFixed(2),
          action: "停止 BTC 新增買入。保留現金以應對潛在回調。"
        };
      }
      return null;
    }
  },
  {
    name: "BOXX Institutional Floor Sniper",
    phase: "All",
    condition: function (context) { return context.market.btcPrice > 0 && context.market.btcPrice <= (CONFIG.THRESHOLDS.INSTITUTIONAL_FLOOR_BTC || 40000); },
    getAction: function (context) {
      return {
        level: "[狙擊] 機構地板價觸發",
        message: "BTC 跌破機構地板價 $" + (CONFIG.THRESHOLDS.INSTITUTIONAL_FLOOR_BTC || 40000),
        action: "立即執行: 將所有 BOXX 流動性轉換為 IBIT/BTC。"
      };
    }
  },
  {
    name: "BTC Martingale Sniper",
    phase: "All",
    condition: function (context) { return CONFIG.BTC_MARTINGALE.ENABLED && context.market.btcPrice > 0 && context.market.btcRecentHighValid; },
    getAction: function (context) {
      const currentDrop = (context.market.btcPrice - context.market.btcRecentHigh) / context.market.btcRecentHigh;
      const strategy = CONFIG.BTC_MARTINGALE;
      let activeLevel = null;
      for (let i = strategy.LEVELS.length - 1; i >= 0; i--) { if (currentDrop <= strategy.LEVELS[i].drop) { activeLevel = strategy.LEVELS[i]; break; } }
      if (activeLevel) {
        if (context.indicators.maintenanceRatio < CONFIG.THRESHOLDS.PLEDGE_RATIO_ALERT) {
          return {
            level: "[暫停] 策略暫停 (維持率 < 2.1)",
            message: "BTC 觸及 " + activeLevel.name + " 跌幅 " + (currentDrop * 100).toFixed(1) + "%",
            action: "由於維持率過低，馬丁格爾買入暫停。"
          };
        }
        return {
          level: "[攻擊] 狙擊信號 (馬丁格爾)",
          message: "BTC 回調 " + (currentDrop * 100).toFixed(1) + "%. 進入 " + activeLevel.name,
          action: "行動: 借入/買入 BTC 金額 TWD " + (strategy.BASE_AMOUNT * activeLevel.multiplier).toLocaleString()
        };
      }
      return null;
    }
  },
  {
    name: "SAP Protocol Rebalancing",
    phase: "All",
    condition: function (context) { return context.rebalanceTargets.length > 0; },
    getAction: function (context) {
      let msg = "[戰略長] 指令:\n";
      context.rebalanceTargets.forEach(item => { msg += "\n" + item.priority + " " + item.name + "\n   - 行動: " + item.action + "\n"; });
      return { level: "[指揮] SAP 戰略指令", message: "偵測到實體配置偏移。", action: msg };
    }
  },
  {
    name: "Binance Crypto Loan Monitor",
    phase: "All",
    condition: function (context) { return context.indicators.binanceMaintenanceRatio > 0 && context.indicators.binanceMaintenanceRatio < CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_SAFE; },
    getAction: function (context) {
      const ratio = context.indicators.binanceMaintenanceRatio;
      if (ratio <= CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_CRITICAL) {
        return {
          level: "[嚴重] 幣安保證金保護 (1.3)",
          message: "幣安 BTC 質押率崩跌至 " + ratio.toFixed(2),
          action: "立即行動: 補倉 BTC 或償還幣安貸款以避免清算。"
        };
      } else if (ratio <= CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_ALERT) {
        return {
          level: "[警告] 幣安風險區 (1.5)",
          message: "幣安質押率在 " + ratio.toFixed(2),
          action: "警告: 檢測到 BTC 高波動。準備抵押品。"
        };
      }
      return null;
    }
  }
];

function showStrategicReportUI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const context = buildContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });

    let msg = "--- [SAP v24.5 指揮中心報告] ---\n";
    msg += "狀態: 活躍 | 模式: " + context.phase + "\n";

    if (alerts.length > 0) {
      alerts.forEach(a => { msg += "\n>> " + a.level + "\n   " + a.message + "\n   行動: " + a.action + "\n"; });
    } else {
      msg += "\n[OK] 實體配置平衡。主權狀態穩固。\n";
    }

    msg += generatePortfolioSnapshot(context);
    msg += "\n保持戰略。保持理性。";
    ui.alert("SAP 指揮中心", msg, ui.ButtonSet.OK);
  } catch (e) { ui.alert("錯誤: " + e.toString()); }
}

function runDailyInvestmentCheck() {
  try {
    const context = buildContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });
    if (alerts.length > 0) { sendEmailAlert(alerts, context); }
    else { sendAllClearEmail(context); }
  } catch (e) {
    const email = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (email) MailApp.sendEmail(email, "[錯誤] SAP 執行失敗", e.toString());
  }
}

function setup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const emailRes = ui.prompt("設定管理員信箱 (ADMIN_EMAIL)", "請輸入 Email:", ui.ButtonSet.OK_CANCEL);
  if (emailRes.getSelectedButton() == ui.Button.OK) props.setProperty('ADMIN_EMAIL', emailRes.getResponseText());
  const reserveRes = ui.prompt("設定預備金 (Treasury Reserve)", "請輸入預備金金額 (TWD):", ui.ButtonSet.OK_CANCEL);
  if (reserveRes.getSelectedButton() == ui.Button.OK) props.setProperty('TREASURY_RESERVE_TWD', reserveRes.getResponseText());
  ui.alert("設定完成。");
}

function buildContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balanceSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BALANCE_SHEET);
  const indicatorSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INDICATORS);

  if (!balanceSheet || !indicatorSheet) {
    throw new Error("找不到必要的工作表 (Balance Sheet / Key Market Indicators)。");
  }

  // Phase 1: 穩健數據收集
  const rawPortfolio = getPortfolioData(balanceSheet);
  const indicatorsRaw = getIndicatorData(indicatorSheet);
  const reserveValue = CONFIG.THRESHOLDS.TREASURY_RESERVE_TWD;

  // Phase 2: 資產/債務分離與聚合
  const portfolioSummary = aggregatePortfolio(rawPortfolio);

  // 總資產
  const totalGrossAssets = Object.values(portfolioSummary).reduce((sum, val) => sum + (val > 0 ? val : 0), 0);
  // 淨實體價值
  const netEntityValue = Object.values(portfolioSummary).reduce((sum, val) => sum + val, 0);

  // Phase 3: 市場數據解析
  const btcPrice = parseFloat(indicatorsRaw.BTC_Price);
  const rawRecentHigh = parseFloat(indicatorsRaw.BTC_Recent_High);
  const usdTwdRate = parseFloat(indicatorsRaw.USDT_TWD || indicatorsRaw.USD_TWD || 32.5);

  let market = {
    btcPrice: isNaN(btcPrice) ? 0 : btcPrice,
    btcRecentHigh: isNaN(rawRecentHigh) ? 0 : rawRecentHigh,
    btcRecentHighValid: !isNaN(rawRecentHigh) && rawRecentHigh > 0,
    usdTwdRate: isNaN(usdTwdRate) ? 32.5 : usdTwdRate
  };

  // Phase 4: 自動質押引擎
  const pledgeGroups = calculateAutoPledgeRatios(rawPortfolio, indicatorsRaw);

  // Phase 5: 再平衡目標
  const portfolioForRebalance = {};
  Object.keys(portfolioSummary).forEach(k => portfolioForRebalance[k] = { Market_Value_TWD: portfolioSummary[k] });
  const targets = getRebalanceTargets(portfolioForRebalance, totalGrossAssets, market);

  // 舊版兼容橋接
  const indicators = {
    isValid: pledgeGroups.length > 0,
    maintenanceRatio: (pledgeGroups.find(g => g.name === "Pledge") || pledgeGroups[0] || { ratio: 0 }).ratio,
    binanceMaintenanceRatio: (pledgeGroups.find(g => g.name === "Binance") || { ratio: 0 }).ratio
  };

  return {
    portfolioSummary,
    rawPortfolio,
    pledgeGroups,
    indicators,
    market,
    phase: "Bitcoin Standard",
    totalGrossAssets: totalGrossAssets,
    netEntityValue: netEntityValue,
    rebalanceTargets: targets,
    reserve: reserveValue
  };
}

function calculateAutoPledgeRatios(rawPortfolio, indicatorsRaw) {
  const labelMap = {};

  rawPortfolio.forEach(item => {
    const label = item.purpose ? item.purpose.trim() : "";
    if (!label || label.toLowerCase() === "none") return;

    if (!labelMap[label]) {
      labelMap[label] = { assets: 0, debt: 0 };
    }

    if (item.value > 0) {
      labelMap[label].assets += item.value;
    } else {
      labelMap[label].debt += Math.abs(item.value);
    }
  });

  const groups = [];
  Object.keys(labelMap).forEach(name => {
    const data = labelMap[name];
    if (data.debt > 0) {
      const ratio = data.assets / data.debt;

      const isCrypto = name.toLowerCase().includes("binance") || name.toLowerCase().includes("okx") || name.toLowerCase().includes("crypto");

      const alertThreshold = parseFloat(indicatorsRaw[name + "_Maint_Alert"]) ||
        (isCrypto ? CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_ALERT : CONFIG.THRESHOLDS.PLEDGE_RATIO_ALERT);

      const criticalThreshold = parseFloat(indicatorsRaw[name + "_Maint_Critical"]) ||
        (isCrypto ? CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_CRITICAL : CONFIG.THRESHOLDS.PLEDGE_RATIO_CRITICAL);

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

function aggregatePortfolio(rawPortfolio) {
  const summary = {};
  rawPortfolio.forEach(item => {
    if (!summary[item.ticker]) summary[item.ticker] = 0;
    summary[item.ticker] += item.value;
  });
  return summary;
}

function getRebalanceTargets(portfolio, assets, market) {
  let targets = [];
  if (assets <= 0) return targets;
  const activeNoise = CONFIG.NOISE_ASSETS.filter(t => portfolio[t] && portfolio[t].Market_Value_TWD > 1000);
  if (activeNoise.length > 0) {
    targets.push({
      name: "戰略雜訊清理 (" + activeNoise.join(", ") + ")",
      priority: "[警告]",
      action: "清理非核心部位，將資金回流 BTC 儲備。"
    });
  }
  CONFIG.ASSET_GROUPS.forEach(g => {
    const val = g.tickers.reduce((s, t) => s + (portfolio[t] ? portfolio[t].Market_Value_TWD : 0), 0);
    const div = (val / assets) - g.target;
    if (g.name.includes("Layer 1") && div < -0.03) {
      targets.push({ name: g.name, priority: "[攻擊]", action: "配置過低 (" + (div * 100).toFixed(1) + "%)。應積極累積 BTC 部位。" });
    } else if (g.name.includes("Layer 2") && div > 0.05) {
      targets.push({ name: g.name, priority: "[防禦]", action: "戰略緩衝過高 (" + (div * 100).toFixed(1) + "%)。可用此基盤執行 BTC 狙擊買入。" });
    }
  });
  return targets;
}

function getPortfolioData(sheet) {
  const data = sheet.getDataRange().getValues();
  const portfolio = [];
  for (let i = 1; i < data.length; i++) {
    const ticker = data[i][0];
    if (ticker) {
      portfolio.push({
        ticker: ticker,
        value: parseFloat(data[i][2]) || 0,
        purpose: data[i][3] || ""
      });
    }
  }
  return portfolio;
}

function getIndicatorData(sheet) {
  const data = sheet.getDataRange().getValues(); const ind = {};
  for (let i = 1; i < data.length; i++) { if (data[i][0]) ind[data[i][0]] = data[i][1]; }
  return ind;
}

function generatePortfolioSnapshot(context) {
  const { portfolioSummary, market, pledgeGroups, totalGrossAssets, netEntityValue, reserve } = context;

  const l1ValueTWD = CONFIG.ASSET_GROUPS[0].tickers.reduce((sum, ticker) => sum + (portfolioSummary[ticker] || 0), 0);
  const btcHeld = l1ValueTWD / (market.usdTwdRate * market.btcPrice);
  const btcGoal = 1.0;

  let s = "\n[I] 市場情報 (MARKET INTEL)\n";
  s += "- BTC 現貨價格: $" + market.btcPrice.toLocaleString() + " USD\n";
  if (market.btcRecentHighValid) {
    const drop = ((market.btcPrice - market.btcRecentHigh) / market.btcRecentHigh * 100).toFixed(1);
    s += "- BTC 高點/回調: $" + market.btcRecentHigh.toLocaleString() + " USD (" + drop + "%)\n";
  }
  s += "- USDT/TWD 匯率: " + market.usdTwdRate.toFixed(2) + " TWD/USD\n";

  if (pledgeGroups.length > 0) {
    pledgeGroups.forEach(group => {
      const safetyBuffer = group.ratio > group.critical ? ((1 - (group.critical / group.ratio)) * 100) : 0;
      s += "\n[II] 質押風險監控 (" + group.name + ")\n";
      s += "- 維持率: " + group.ratio.toFixed(2) + " (警戒線: " + group.critical + ")\n";
      s += "- 抵押價值: " + Math.round(group.collateralValue).toLocaleString() + " TWD\n";
      if (group.loanAmount > 0) {
        s += "- 貸款金額: " + Math.round(group.loanAmount).toLocaleString() + " TWD\n";
      }
      s += "- 安全緩衝: " + safetyBuffer.toFixed(1) + "% (可承受最大回調)\n";
    });
  }

  s += "\n[III] 資產配置健康度 (ALLOCATION HEALTH)\n";
  let coreGroupsValue = 0;
  CONFIG.ASSET_GROUPS.forEach(g => {
    const v = g.tickers.reduce((sum, t) => sum + (portfolioSummary[t] || 0), 0);
    coreGroupsValue += v;
    const weight = totalGrossAssets > 0 ? (v / totalGrossAssets) : 0;
    const drift = (weight - g.target) * 100;
    s += "- " + g.name + ": " + (weight * 100).toFixed(1) + "% [目標: " + (g.target * 100) + "%, 偏移: " + (drift > 0 ? "+" : "") + drift.toFixed(1) + "%]\n";
  });

  const noiseValue = totalGrossAssets - coreGroupsValue;
  const noiseWeight = totalGrossAssets > 0 ? (noiseValue / totalGrossAssets) : 0;
  if (noiseWeight > 0.001) {
    s += "- Layer 0: 雜訊與非核心部位: " + (noiseWeight * 100).toFixed(1) + "%\n";
  }

  const totalDebt = totalGrossAssets - netEntityValue;
  const leverageRatio = netEntityValue > 0 ? (totalGrossAssets / netEntityValue) : 0;

  s += "\n[IV] 流動性與槓桿 (LIQUIDITY & LEVERAGE)\n";
  s += "- 總負債金額: " + Math.round(totalDebt).toLocaleString() + " TWD\n";
  s += "- 槓桿倍數: " + leverageRatio.toFixed(2) + "x (總資產/淨資產)\n";
  s += "- 淨實體價值: " + Math.round(netEntityValue).toLocaleString() + " TWD\n";

  s += "\n[V] 目標進度 (TARGET PROGRESS)\n";
  s += "- 1 BTC 達成率: " + (isNaN(btcHeld) ? "0" : ((btcHeld / btcGoal) * 100).toFixed(1)) + "%\n";
  s += "- 預備金儲備: " + reserve.toLocaleString() + " TWD\n";
  s += "----------------------------------------\n";
  s += "最後更新: " + new Date().toLocaleString('zh-TW', { hour12: false });

  return s;
}

function sendEmailAlert(alerts, context) {
  let sub = "[SAP 戰略顧問] 需要採取行動";
  let body = "戰略夥伴，\n分析顯示需要進行再平衡：\n\n";
  alerts.forEach(a => { body += "**" + a.level + "**\n" + a.message + "\n指令: " + a.action + "\n\n"; });
  body += generatePortfolioSnapshot(context);
  MailApp.sendEmail(CONFIG.EMAIL_RECIPIENT, sub, body);
}

function sendAllClearEmail(context) {
  MailApp.sendEmail(CONFIG.EMAIL_RECIPIENT, "[SAP 每日狀態] 一切正常", generatePortfolioSnapshot(context));
}
