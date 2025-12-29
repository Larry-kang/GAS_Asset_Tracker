/**
 * @OnlyCurrentDoc
 * Sovereign Asset Protocol (SAP) v24.5 - Command Center Edition
 * Version: 24.5 (Bitcoin Standard / Treasury Edition)
 */

const CONFIG = {
  SHEET_NAMES: {
    BALANCE_SHEET: "Balance Sheet",
    INDICATORS: "Key Market Indicators",
    DASHBOARD: "Dashboard"
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
    BASE_AMOUNT: 20000,
    LEVELS: [
      { drop: -0.40, multiplier: 1, name: "Level 1 (Sniper Zone)" },
      { drop: -0.50, multiplier: 2, name: "Level 2 (Deep Value)" },
      { drop: -0.60, multiplier: 3, name: "Level 3 (Abyss)" },
      { drop: -0.70, multiplier: 4, name: "Level 4 (Capitulation)" }
    ]
  },
  ASSET_GROUPS: [
    { name: "Layer 1: Digital Reserve (Attack)", target: 0.60, tickers: ["IBIT", "BTC_Spot", "BTC"] },
    { name: "Layer 2: Credit Base (Defend)", target: 0.30, tickers: ["00713", "00662", "QQQ"] },
    { name: "Layer 3: Tactical Liquidity", target: 0.10, tickers: ["BOXX", "CASH_TWD", "USDT", "USDC"] }
  ],
  NOISE_ASSETS: ["ETH", "BNB", "TQQQ"]
};

const RULES = [
  {
    name: "ATH Breakout Monitor",
    phase: "All",
    condition: function (context) {
      return context.market.sapBaseATH > 0 && context.market.btcPrice > (context.market.sapBaseATH * 1.05);
    },
    getAction: function (context) {
      return {
        level: "[情報] 新高點偵測",
        message: "BTC 價格 ($" + context.market.btcPrice + ") 已超越定錨高點 5%。",
        action: "建議手動校準 Key Market Indicators 中的 `SAP_Base_ATH` 以重置下行狙擊線。"
      };
    }
  },
  {
    name: "BTC Martingale Sniper (v24.5)",
    phase: "All",
    condition: function (context) {
      return CONFIG.BTC_MARTINGALE.ENABLED &&
        context.market.sapBaseATH > 0 &&
        context.market.totalMartingaleSpent < context.market.maxMartingaleBudget;
    },
    getAction: function (context) {
      const currentDrop = (context.market.btcPrice - context.market.sapBaseATH) / context.market.sapBaseATH;
      const strategy = CONFIG.BTC_MARTINGALE;

      let activeLevel = null;
      for (let i = strategy.LEVELS.length - 1; i >= 0; i--) {
        if (currentDrop <= strategy.LEVELS[i].drop) {
          activeLevel = strategy.LEVELS[i];
          break;
        }
      }

      if (activeLevel) {
        const estCost = strategy.BASE_AMOUNT * activeLevel.multiplier;
        if (context.market.totalMartingaleSpent + estCost > context.market.maxMartingaleBudget) {
          return {
            level: "[警告] 狙擊預算不足",
            message: "觸發 " + activeLevel.name + " 但預算不足 (剩餘: " + (context.market.maxMartingaleBudget - context.market.totalMartingaleSpent) + ")",
            action: "請手動檢查或增加預算。"
          };
        }

        return {
          level: "[攻擊] 狙擊信號 (Sniper)",
          message: "BTC 回調 " + (currentDrop * 100).toFixed(1) + "% (基準: $" + context.market.sapBaseATH + "). 進入 " + activeLevel.name,
          action: "執行買入: TWD " + estCost.toLocaleString() + " 等值 BTC。\n(執行後請手動更新 `Total_Martingale_Spent` += " + estCost + ")"
        };
      }
      return null;
    }
  },
  {
    name: "Cashflow Rerouting Engine",
    phase: "All",
    condition: function (context) {
      return context.market.surplus > 0 || context.rebalanceTargets.length > 0;
    },
    getAction: function (context) {
      // Priority 1: Check L1 Spot Ratio
      if (context.indicators.l1SpotRatio < 0.60) { // 60% hard target
        return {
          level: "[配置] 資金流向建議 (補強地基)",
          message: "L1 現貨佔比 (" + (context.indicators.l1SpotRatio * 100).toFixed(1) + "%) 低於 60%。",
          action: "將盈餘/現金 100% 買入現貨 BTC (存放於冷錢包/OKX)。"
        };
      }
      // Priority 2: Check Overheated (Total BTC > 80%)
      else if (context.indicators.totalBtcRatio > 0.80) {
        if (context.indicators.totalBtcRatio > 0.90) {
          return {
            level: "[配置] 資金流向建議 (極度貪婪)",
            message: "BTC 總佔比 (" + (context.indicators.totalBtcRatio * 100).toFixed(1) + "%) 超過 90%。",
            action: "將盈餘 100% 轉入 USDT/USDC 或法幣現金，停止任何投資。"
          };
        }
        return {
          level: "[配置] 資金流向建議 (防禦護城河)",
          message: "BTC 總佔比高於 80%。部位過重。",
          action: "將盈餘 100% 買入 00713 (高股息 ETF) 以強化信用基底。"
        };
      }

      return null;
    }
  },
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

    updateDashboard(context);

    if (alerts.length > 0) { sendEmailAlert(alerts, context); }
    else { sendAllClearEmail(context); }
  } catch (e) {
    const email = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (email) MailApp.sendEmail(email, "[錯誤] SAP 執行失敗", e.toString());
  }
}

/**
 * Executes the frequent (30-min) strategic monitor.
 * Updates Dashboard and sends alerts only if critical.
 */
function runStrategicMonitor() {
  try {
    const context = buildContext();
    updateDashboard(context);

    let alerts = [];
    RULES.forEach(rule => {
      if (rule.condition(context)) {
        const action = rule.getAction(context);
        if (action) alerts.push(action);
      }
    });

    if (alerts.length > 0) {
      sendEmailAlert(alerts, context);
    }
  } catch (e) {
    Logger.log("[Monitor Error] " + e.toString());
  }
}

function updateDashboard(context) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  if (!sheet) return;

  // Use TextFinder to locate cells dynamically
  const metrics = {
    "Survival Runway": context.indicators.survivalRunway.toFixed(1) + " Months",
    "LTV": (context.indicators.ltv * 100).toFixed(1) + "%",
    "Net Entity Value": Math.round(context.netEntityValue).toLocaleString(),
    "BTC Price": "$" + context.market.btcPrice.toLocaleString(),
    "Last Update": new Date().toLocaleString('zh-TW', { hour12: false })
  };

  Object.keys(metrics).forEach(key => {
    const finder = sheet.createTextFinder(key);
    const cell = finder.findNext();
    if (cell) {
      cell.offset(0, 1).setValue(metrics[key]);
    }
  });
}

function setup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  // 1. 設定 Email
  const emailRes = ui.prompt("設定管理員信箱 (ADMIN_EMAIL)", "請輸入 Email:", ui.ButtonSet.OK_CANCEL);
  if (emailRes.getSelectedButton() == ui.Button.OK) props.setProperty('ADMIN_EMAIL', emailRes.getResponseText());

  // 2. 自動設定排程 (預設: 每日早上 6 點)
  const defaultHour = 6;
  props.setProperty('SCHEDULER_MODE', 'DAILY');
  props.setProperty('SCHEDULER_HOUR', defaultHour.toString());
  props.setProperty('SCHEDULER_ENABLED', 'true');

  // 清除舊觸發器
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const func = t.getHandlerFunction();
    if (func === 'runDailyInvestmentCheck' || func === 'autoRecordDailyValues' || func === 'updateAllPrices' || func === 'runAutomationMaster') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 建立新觸發器
  ScriptApp.newTrigger('runAutomationMaster').timeBased().everyDays(1).atHour(defaultHour).create();
  ScriptApp.newTrigger('autoRecordDailyValues').timeBased().everyDays(1).atHour(1).create();

  ui.alert(`設定完成！\n已為您自動部署戰略排程：\n- 每日報告: 早上 ${defaultHour}:00 (美股收盤後)\n- 資產快照: 凌晨 01:00`);
}

function buildContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balanceSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BALANCE_SHEET);
  const indicatorSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INDICATORS);

  if (!balanceSheet || !indicatorSheet) {
    throw new Error("找不到必要的工作表 (Balance Sheet / " + CONFIG.SHEET_NAMES.INDICATORS + ")。");
  }

  // Phase 1: 穩健數據收集
  const rawPortfolio = getPortfolioData(balanceSheet);
  const indicatorsRaw = fetchMarketIndicators(indicatorSheet);

  // Phase 2: 資產/債務分離與聚合
  const portfolioSummary = aggregatePortfolio(rawPortfolio);

  // 總資產
  const totalGrossAssets = Object.values(portfolioSummary).reduce((sum, val) => sum + (val > 0 ? val : 0), 0);
  // 淨實體價值
  const netEntityValue = Object.values(portfolioSummary).reduce((sum, val) => sum + val, 0);

  // Phase 3: 市場數據解析
  let market = {
    btcPrice: indicatorsRaw.Current_BTC_Price || 0,
    sapBaseATH: indicatorsRaw.SAP_Base_ATH || 0,
    totalMartingaleSpent: indicatorsRaw.Total_Martingale_Spent || 0,
    maxMartingaleBudget: indicatorsRaw.MAX_MARTINGALE_BUDGET || 437000,
    usdTwdRate: 32.5,
    surplus: 0
  };

  const monthlyDebt = indicatorsRaw.MONTHLY_DEBT_COST || 12967;
  const liquidity = (portfolioSummary["CASH_TWD"] || 0) + (portfolioSummary["USDT"] || 0) + (portfolioSummary["USDC"] || 0);
  const survivalRunway = monthlyDebt > 0 ? (liquidity / monthlyDebt) : 99;

  market.surplus = liquidity - (monthlyDebt * 3);

  // Phase 4: 自動質押引擎
  const pledgeGroups = calculateAutoPledgeRatios(rawPortfolio, {});

  // Phase 5: 再平衡目標
  const portfolioForRebalance = {};
  Object.keys(portfolioSummary).forEach(k => portfolioForRebalance[k] = { Market_Value_TWD: portfolioSummary[k] });
  const targets = getRebalanceTargets(portfolioForRebalance, totalGrossAssets, market);

  const indicators = {
    isValid: pledgeGroups.length > 0,
    maintenanceRatio: (pledgeGroups.find(g => g.name === "Pledge") || pledgeGroups[0] || { ratio: 0 }).ratio,
    binanceMaintenanceRatio: (pledgeGroups.find(g => g.name === "Binance") || { ratio: 0 }).ratio,
    l1SpotRatio: totalGrossAssets > 0 ? (calculateGroupValue(portfolioSummary, CONFIG.ASSET_GROUPS[0]) / totalGrossAssets) : 0,
    totalBtcRatio: totalGrossAssets > 0 ? (calculateGroupValue(portfolioSummary, CONFIG.ASSET_GROUPS[0]) / totalGrossAssets) : 0,
    survivalRunway: survivalRunway,
    ltv: totalGrossAssets > 0 ? (totalGrossAssets - netEntityValue) / totalGrossAssets : 0
  };

  return {
    portfolioSummary,
    rawPortfolio,
    pledgeGroups,
    indicators,
    market,
    phase: "Bitcoin Standard v24.5",
    totalGrossAssets: totalGrossAssets,
    netEntityValue: netEntityValue,
    rebalanceTargets: targets,
    reserve: liquidity
  };
}

function fetchMarketIndicators(sheet) {
  const data = sheet.getDataRange().getValues();
  const result = {};

  const keysOfInterest = [
    "SAP_Base_ATH",
    "Total_Martingale_Spent",
    // "L1_Spot_Ratio", // Deprecated: Calculated in code v24.5
    "Current_BTC_Price",
    // "Total_BTC_Ratio", // Deprecated: Calculated in code v24.5
    "MAX_MARTINGALE_BUDGET",
    "MONTHLY_DEBT_COST"
  ];

  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (keysOfInterest.includes(key) || key.indexOf("SAP_") > -1) {
      result[key] = parseFloat(data[i][1]);
    }
  }
  return result;
}

function calculateAutoPledgeRatios(rawPortfolio, indicatorsRaw) {
  const labelMap = {};
  rawPortfolio.forEach(item => {
    const label = item.purpose ? item.purpose.trim() : "";
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
      const isCrypto = name.toLowerCase().includes("binance") || name.toLowerCase().includes("okx");
      const alertThreshold = isCrypto ? CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_ALERT : CONFIG.THRESHOLDS.PLEDGE_RATIO_ALERT;
      const criticalThreshold = isCrypto ? CONFIG.THRESHOLDS.CRYPTO_LOAN_RATIO_CRITICAL : CONFIG.THRESHOLDS.PLEDGE_RATIO_CRITICAL;
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

function calculateGroupValue(summary, group) {
  let value = 0;
  group.tickers.forEach(t => value += (summary[t] || 0));
  return value;
}

function getRebalanceTargets(portfolio, assets, market) {
  let targets = [];
  if (assets <= 0) return targets;
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

function generatePortfolioSnapshot(context) {
  const { market, pledgeGroups, netEntityValue, indicators, totalGrossAssets, portfolioSummary } = context;

  let s = "\n[I] 市場情報 (MARKET INTEL)\n";
  s += "- BTC 現貨價格: $" + market.btcPrice.toLocaleString() + " USD\n";
  if (market.sapBaseATH > 0) {
    const drop = ((market.btcPrice - market.sapBaseATH) / market.sapBaseATH * 100).toFixed(1);
    s += "- 距離 ATH (" + market.sapBaseATH + "): " + drop + "%\n";
  }

  s += "\n[II] 生存指標 (SURVIVAL METRICS)\n";
  s += "- 生存跑道: " + indicators.survivalRunway.toFixed(1) + " 個月\n";
  s += "- 淨值: " + Math.round(netEntityValue).toLocaleString() + " TWD\n";
  s += "- 總資產: " + Math.round(totalGrossAssets).toLocaleString() + " TWD\n";
  s += "- 總負債: " + Math.round(totalGrossAssets - netEntityValue).toLocaleString() + " TWD\n";
  s += "- 負債比 (LTV): " + (indicators.ltv * 100).toFixed(1) + "%\n";

  if (pledgeGroups.length > 0) {
    pledgeGroups.forEach(group => {
      s += "- 維持率 (" + group.name + "): " + group.ratio.toFixed(2) + "\n";
    });
  }

  s += "\n[III] 資產配置 (ASSET ALLOCATION)\n";
  CONFIG.ASSET_GROUPS.forEach(group => {
    let groupValue = 0;
    group.tickers.forEach(t => groupValue += (portfolioSummary[t] || 0));

    // Safety check for Noise Assets not in groups might be needed, but for now stick to groups
    // Actually, calculate total based on group tickers for percentage might differ from totalGrossAssets if there are unclassified assets.
    // Use totalGrossAssets as denominator.
    const pct = totalGrossAssets > 0 ? (groupValue / totalGrossAssets * 100) : 0;
    s += "- " + group.name.split(":")[0] + ": " + Math.round(groupValue).toLocaleString() + " (" + pct.toFixed(1) + "%)\n";
  });

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
