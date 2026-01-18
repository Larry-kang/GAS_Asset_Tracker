/**
 * Sovereign Asset Protocol (SAP) - Command Center Edition
 * Version: v24.6 (Bitcoin Standard / Treasury Edition)
 */

const CONFIG = {
  SHEET_NAMES: Config.SHEET_NAMES,
  get EMAIL_RECIPIENT() {
    const email = Settings.get('ADMIN_EMAIL');
    if (!email) {
      LogService.warn("ScriptProperty 'ADMIN_EMAIL' not set.", "Config:Email");
    }
    return email || "";
  },
  THRESHOLDS: {
    ...Config.STRATEGIC,
    get TREASURY_RESERVE_TWD() {
      return Settings.getNumber('TREASURY_RESERVE_TWD', 100000);
    }
  },
  BTC_MARTINGALE: Config.BTC_MARTINGALE,
  ASSET_GROUPS: Config.ASSET_GROUPS,
  NOISE_ASSETS: Config.NOISE_ASSETS
};

/**
 * Execution Cache to prevent redundant Sheet/Properties lookups within the same run.
 */
const DataCache = {
  _sheets: {},
  getValues: function (sheetName) {
    if (!this._sheets[sheetName]) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return null;
      this._sheets[sheetName] = sheet.getDataRange().getValues();
    }
    return this._sheets[sheetName];
  }
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
    name: "BTC Martingale Sniper",
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
      const l1Target = context.assetGroups ? context.assetGroups[0].target : 0.60;
      if (context.indicators.l1SpotRatio < l1Target) {
        return {
          level: "[配置] 資金流向建議 (補強地基)",
          message: "L1 現貨佔比 (" + (context.indicators.l1SpotRatio * 100).toFixed(1) + "%) 低於 " + (l1Target * 100).toFixed(0) + "%。",
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
  },
  {
    name: "Noise Asset Cleanup Monitor",
    phase: "All",
    condition: function (context) {
      const l4 = context.assetGroups ? context.assetGroups.find(g => g.id === "L4") : null;
      return l4 && l4.value > 0;
    },
    getAction: function (context) {
      const l4 = context.assetGroups.find(g => g.id === "L4");
      return {
        level: "[注意] 雜項資產清理建議",
        message: "偵測到 Layer 4 雜項資產: " + l4.tickers.join(", ") + " (總值: " + Math.round(l4.value).toLocaleString() + " TWD)",
        action: "建議找市場高位機會清空雜項資產，回歸 L1 (BTC) 或 L2 (穩定基底)。"
      };
    }
  }
];

/**
 * Displays strategic report UI with current market status and alerts.
 * Shows portfolio snapshot, risk indicators, and action recommendations.
 * Offers option to broadcast report to Discord/Email.
 * @public
 */
function showStrategicReportUI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const context = buildContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });

    let msg = `--- [${Config.SYSTEM_NAME.split(' - ')[0]} ${Config.VERSION} 指揮中心報告] ---\n`;
    msg += "狀態: 活躍 | 模式: " + context.phase + "\n";

    if (alerts.length > 0) {
      alerts.forEach(a => { msg += "\n>> " + a.level + "\n   " + a.message + "\n   行動: " + a.action + "\n"; });
    } else {
      msg += "\n[OK] 實體配置平衡。主權狀態穩固。\n";
    }

    msg += generatePortfolioSnapshot(context);
    msg += "\n保持戰略。保持理性。";

    // [User Request P3-2] Manual Trigger Sync
    const result = ui.alert("SAP 指揮中心", msg + "\n\n是否同步發送此報告？ (Discord/Email)", ui.ButtonSet.YES_NO);
    if (result == ui.Button.YES) {
      broadcastReport_(context, alerts);
      ui.alert("✅ 報告已發送。");
    }

  } catch (e) { ui.alert("錯誤: " + e.toString()); }
}

/**
 * Executes daily investment check and broadcasts report.
 * Analyzes market conditions, triggers alerts, and updates dashboard.
 * Called by daily trigger at scheduled time.
 * @public
 */
function runDailyInvestmentCheck() {
  try {
    const context = buildContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });

    updateDashboard(context);

    // [User Request P3-2] Auto Sync
    broadcastReport_(context, alerts);

  } catch (e) {
    const email = Settings.get('ADMIN_EMAIL');
    if (email) MailApp.sendEmail(email, "[錯誤] SAP 執行失敗", e.toString());
  }
}

/**
 * Executes frequent (30-min) strategic monitor.
 * Updates dashboard and logs alerts without sending emails (noise reduction).
 * @public
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
      // sendEmailAlert(alerts, context); // [User Request 2025-12-30] 降噪模式：僅更新 Dashboard，不發信
      Logger.log("[Monitor] Alerts generated but silenced (Daily Report only).");
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

/**
 * Initial Setup Wizard
 * Guides user through configuration and initializes high-frequency monitoring.
 */
function setup() {
  const ui = SpreadsheetApp.getUi();
  // Settings manager handles internal properties

  // Step 1: Configure Admin Email
  const emailRes = ui.prompt(
    "設定管理員信箱 (ADMIN_EMAIL)",
    "請輸入接收戰略報告的 Email:",
    ui.ButtonSet.OK_CANCEL
  );

  if (emailRes.getSelectedButton() == ui.Button.OK) {
    const email = emailRes.getResponseText().trim();
    if (email) {
      Settings.set('ADMIN_EMAIL', email);
      LogService.info('Email configured: ' + email, 'Setup');
    }
  }

  // Step 1.5: Configure Discord Webhook (Optional)
  const discordRes = ui.alert(
    "設定 Discord 通知",
    "是否啟用 Discord 即時警報？\n(推薦啟用，可即時接收策略訊號)",
    ui.ButtonSet.YES_NO
  );

  if (discordRes == ui.Button.YES) {
    const webhookRes = ui.prompt(
      "設定 Discord Webhook",
      "請貼上 Webhook URL:\n(若不知如何獲取，請詢問 CTO)",
      ui.ButtonSet.OK_CANCEL
    );
    if (webhookRes.getSelectedButton() == ui.Button.OK) {
      const url = webhookRes.getResponseText().trim();
      if (url) {
        Settings.set('DISCORD_WEBHOOK_URL', url);
        LogService.info('Discord Webhook configured', 'Setup');
      }
    }
  }

  // Step 2: Configure Emergency Reserve Threshold
  const reserveRes = ui.prompt(
    "設定緊急預備金門檻 (TREASURY_RESERVE_TWD)",
    "請輸入金額（TWD，預設 100000）:\n\n" +
    "此門檻用於判斷流動性健康度。",
    ui.ButtonSet.OK_CANCEL
  );

  if (reserveRes.getSelectedButton() == ui.Button.OK) {
    const amount = parseFloat(reserveRes.getResponseText());
    if (!isNaN(amount) && amount > 0) {
      Settings.set('TREASURY_RESERVE_TWD', amount.toString());
      LogService.info('Treasury Reserve set to: ' + amount, 'Setup');
    }
  }

  // Step 3: Initialize High-Frequency Monitoring
  // Delegate to setupScheduledTriggers for standardized trigger configuration
  if (typeof setupScheduledTriggers === 'function') {
    setupScheduledTriggers();
  } else {
    ui.alert(
      '❌ 錯誤',
      '找不到 setupScheduledTriggers 函數。\n請確認 Scheduler_Triggers.js 已正確載入。',
      ui.ButtonSet.OK
    );
  }
}

function buildContext() {
  // Phase 1: 穩健數據收集
  const rawPortfolio = getPortfolioData(CONFIG.SHEET_NAMES.BALANCE_SHEET);
  const indicatorsRaw = fetchMarketIndicators(CONFIG.SHEET_NAMES.INDICATORS);

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

  // Phase 4: 動態資產配置目標注入與 Layer 4 自動化 (v24.7)
  const knownTickers = new Set();
  const assetGroups = CONFIG.ASSET_GROUPS.map(group => {
    group.tickers.forEach(t => knownTickers.add(t));
    let dynamicTarget = group.defaultTarget;
    const key = "Alloc_" + group.id + "_Target";
    if (indicatorsRaw[key] !== undefined && !isNaN(indicatorsRaw[key])) {
      dynamicTarget = indicatorsRaw[key];
    }
    return { ...group, target: dynamicTarget, value: calculateGroupValue(portfolioSummary, group) };
  });

  // 識別雜項資產 (Layer 4)
  const miscTickers = Object.keys(portfolioSummary).filter(t => !knownTickers.has(t) && !CONFIG.NOISE_ASSETS.includes(t) && portfolioSummary[t] > 0);
  const miscValue = miscTickers.reduce((sum, t) => sum + portfolioSummary[t], 0);

  assetGroups.push({
    id: "L4",
    name: "Layer 4: Miscellaneous (To Clear)",
    target: 0,
    tickers: miscTickers,
    value: miscValue,
    isMisc: true
  });

  // Phase 5: 自動質押引擎
  const pledgeGroups = calculateAutoPledgeRatios(rawPortfolio, {});

  // Phase 5: 再平衡目標
  const portfolioForRebalance = {};
  Object.keys(portfolioSummary).forEach(k => portfolioForRebalance[k] = { Market_Value_TWD: portfolioSummary[k] });
  const targets = getRebalanceTargets(portfolioForRebalance, totalGrossAssets, market);

  const indicators = {
    isValid: pledgeGroups.length > 0,
    maintenanceRatio: (pledgeGroups.find(g => g.name === "Pledge") || pledgeGroups[0] || { ratio: 0 }).ratio,
    binanceMaintenanceRatio: (pledgeGroups.find(g => g.name === "Binance") || { ratio: 0 }).ratio,
    l1SpotRatio: totalGrossAssets > 0 ? (assetGroups[0].value / totalGrossAssets) : 0,
    totalBtcRatio: totalGrossAssets > 0 ? (assetGroups[0].value / totalGrossAssets) : 0,
    survivalRunway: survivalRunway,
    ltv: totalGrossAssets > 0 ? (totalGrossAssets - netEntityValue) / totalGrossAssets : 0
  };

  return {
    portfolioSummary,
    rawPortfolio,
    pledgeGroups,
    indicators,
    market,
    assetGroups, // 注入動態生成的組態 (含 L4)
    phase: "Bitcoin Standard " + Config.VERSION,
    totalGrossAssets: totalGrossAssets,
    netEntityValue: netEntityValue,
    rebalanceTargets: targets,
    reserve: liquidity
  };
}

function fetchMarketIndicators(sheetName) {
  const data = DataCache.getValues(sheetName);
  if (!data) throw new Error("Could not find indicator sheet: " + sheetName);
  const result = {};

  const keysOfInterest = [
    "SAP_Base_ATH",
    "Total_Martingale_Spent",
    // "L1_Spot_Ratio", // Deprecated: Calculated in code v24.5
    "Current_BTC_Price",
    // "Total_BTC_Ratio", // Deprecated: Calculated in code v24.5
    "MAX_MARTINGALE_BUDGET",
    "MONTHLY_DEBT_COST",
    "Alloc_L1_Target",
    "Alloc_L2_Target",
    "Alloc_L3_Target"
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

function getPortfolioData(sheetName) {
  const data = DataCache.getValues(sheetName);
  if (!data) throw new Error("Could not find balance sheet: " + sheetName);
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
  const groupsToDisplay = context.assetGroups || CONFIG.ASSET_GROUPS;
  groupsToDisplay.forEach(group => {
    let groupValue = group.value || 0;
    if (group.value === undefined) {
      group.tickers.forEach(t => groupValue += (portfolioSummary[t] || 0));
    }

    const pct = totalGrossAssets > 0 ? (groupValue / totalGrossAssets * 100) : 0;
    const targetPct = (group.target || group.defaultTarget || 0) * 100;

    let line = "- " + group.name.split(":")[0] + ": " + Math.round(groupValue).toLocaleString() + " (" + pct.toFixed(1) + "%";
    if (!group.isMisc) {
      line += " / 目標 " + targetPct.toFixed(0) + "%)\n";
    } else {
      line += ")\n";
      if (groupValue > 0) {
        line += "  > ⚠️ 待清理: " + group.tickers.join(", ") + "\n";
      }
    }
    s += line;
  });

  s += "----------------------------------------\n";
  s += "最後更新: " + new Date().toLocaleString('zh-TW', { hour12: false });
  return s;
}

/**
 * Unified Broadcast Handler (Email + Discord)
 * @private
 */
function broadcastReport_(context, alerts = []) {
  const hasAlerts = alerts.length > 0;
  const snapshot = generatePortfolioSnapshot(context);

  // 1. Email Channel
  const emailRecipient = CONFIG.EMAIL_RECIPIENT;
  if (emailRecipient) {
    let subject = hasAlerts ? "[SAP 戰略顧問] 需要採取行動" : "[SAP 每日狀態] 一切正常";
    let body = hasAlerts ? "戰略夥伴，\n分析顯示需要進行再平衡：\n\n" : "戰略夥伴，\n目前系統運作正常。\n\n";

    if (hasAlerts) {
      alerts.forEach(a => { body += "**" + a.level + "**\n" + a.message + "\n指令: " + a.action + "\n\n"; });
    }
    body += snapshot;

    MailApp.sendEmail(emailRecipient, subject, body);
    console.log(`[Broadcast] Email sent to ${emailRecipient}`);
  }

  // 2. Discord Channel (Sync)
  if (typeof Discord !== 'undefined') {
    const title = hasAlerts ? "🚨 SAP 戰略行動報告" : "✅ SAP 每日狀態報告";
    const color = hasAlerts ? "WARNING" : "SUCCESS";

    // Format description for Embed
    let description = "";
    if (hasAlerts) {
      description += "**需要採取行動**\n";
      alerts.forEach(a => { description += `> **${a.level}**\n> ${a.message}\n> *${a.action}*\n\n`; });
      description += "\n";
    }

    // Add Snapshot in Code Block for monospace alignment
    description += "```yaml\n" + snapshot.replace(/`/g, '') + "\n```";

    Discord.sendAlert(title, description, color);
  }
}

