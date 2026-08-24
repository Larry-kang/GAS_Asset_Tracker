/**
 * Core_StrategicEngine.js
 * Sovereign Asset Protocol - Strategic Orchestration & Sheet Dashboard Bridge
 */

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
  },
  clear: function (sheetName) {
    if (sheetName) {
      delete this._sheets[sheetName];
      return;
    }
    this._sheets = {};
  }
};

/**
 * Rebuilds runtime context from fresh sheet reads.
 * Use this at top-level entrypoints when the workbook may have changed since the
 * previous context build in the same execution.
 */
function buildFreshContext() {
  DataCache.clear();
  return buildContext();
}


function showStrategicReportUI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const context = buildFreshContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });

    // Preferred Modern Experience: Show Interactive Sidebar
    if (typeof generatePortfolioSnapshotHtml === 'function' && typeof HtmlService !== 'undefined') {
      const htmlContent = generatePortfolioSnapshotHtml(context, alerts, { isInteractive: true });
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setTitle('⚡ SAP 戰略指揮中心')
        .setWidth(450);
      ui.showSidebar(htmlOutput);
      return;
    }

    // Fallback to modal dialog if HTML rendering is not available
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
    const context = buildFreshContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });

    updateDashboard(context);

    // [User Request P3-2] Auto Sync
    broadcastReport_(context, alerts);

    // [New v24.12] Trigger Snapshot with full context to ensure accuracy
    autoRecordDailyValues(context);

  } catch (e) {
    const email = Settings.get('ADMIN_EMAIL');
    if (email) MailApp.sendEmail(email, "[錯誤] SAP 執行失敗", e.toString());
  }
}


/**
 * Executes Frequent Strategic Monitor.
 * Updates dashboard and logs alerts without sending emails (noise reduction).
 * @public
 */
function runStrategicMonitor() {
  try {
    const context = buildFreshContext();
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

/**
 * [Webhook Target]
 * Manually triggered report via Discord / API.
 * Does NOT record daily snapshot history, only broadcasts current state.
 * @public
 * @returns {Object} JSON result for webhook response
 */
function triggerManualReport() {
  try {
    const context = buildFreshContext();
    let alerts = [];
    RULES.forEach(rule => { if (rule.condition(context)) { const action = rule.getAction(context); if (action) alerts.push(action); } });

    // Optional: Update dashboard on manual trigger? Yes, why not.
    updateDashboard(context);

    broadcastReport_(context, alerts);

    return {
      status: "success",
      message: "Report generated and sent to Discord/Email.",
      alertsCount: alerts.length
    };

  } catch (e) {
    console.error("Manual Trigger Failed", e);
    return { status: "error", message: e.toString() };
  }
}


function updateDashboard(context) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config.SHEET_NAMES.DASHBOARD);
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

  // Step 1.8: Configure Web App Dashboard Access Key (Optional)
  const dashboardKeyRes = ui.prompt(
    "設定網頁儀表板密鑰 (DASHBOARD_ACCESS_KEY)",
    "請輸入網頁看盤密碼 (若留空則預設採用 PROXY_PASSWORD)：",
    ui.ButtonSet.OK_CANCEL
  );

  if (dashboardKeyRes.getSelectedButton() == ui.Button.OK) {
    const dashKey = dashboardKeyRes.getResponseText().trim();
    if (dashKey) {
      Settings.set('DASHBOARD_ACCESS_KEY', dashKey);
      LogService.info('Dashboard Access Key configured', 'Setup');
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
  const rawPortfolio = getPortfolioData(Config.SHEET_NAMES.BALANCE_SHEET);
  const indicatorsRaw = fetchMarketIndicators(Config.SHEET_NAMES.INDICATORS);

  // Phase 2: 資產/債務分離與聚合
  const portfolioSummary = aggregatePortfolio(rawPortfolio);
  let inventoryExport = null;
  try {
    inventoryExport = typeof getInventoryExportBundle_ === "function"
      ? getInventoryExportBundle_()
      : null;
  } catch (e) {
    LogService.warn("Inventory export unavailable: " + e.toString(), "StrategicEngine");
  }
  const stockStrategy = buildStockExposureStrategy_(inventoryExport);

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
    // [NEW v24.10]
    btcMM: indicatorsRaw.BTC_MM || null,
    btcDrawdownFromATH: null,
    btcRegime: null,
    usdTwdRate: 32.5,
    surplus: 0,
    // [NEW v24.13] TW Weighted MM Calculation
    twWeightedMM: null,
    twMMParts: { mm713: 0, mm662: 0 },
    // [NEW v24.14] TW Stock Prices
    price713: indicatorsRaw["00713_Price"] || 0,
    price662: indicatorsRaw["00662_Price"] || 0
  };

  if (indicatorsRaw["00713_MM"] && indicatorsRaw["00662_MM"]) {
    market.twMMParts.mm713 = indicatorsRaw["00713_MM"];
    market.twMMParts.mm662 = indicatorsRaw["00662_MM"];

    // Calculate Weights based on Real Portfolio Value
    const val713 = portfolioSummary['00713'] || 0;
    const val662 = portfolioSummary['00662'] || portfolioSummary['00662_TW'] || 0; // Handle alias
    const totalTW = val713 + val662;

    if (totalTW > 0) {
      // Real-time Weight Priority
      market.twWeightedMM = (market.twMMParts.mm713 * (val713 / totalTW)) + (market.twMMParts.mm662 * (val662 / totalTW));
    } else {
      // Fallback Strategy Weight (66% : 33%)
      market.twWeightedMM = (market.twMMParts.mm713 * 0.66) + (market.twMMParts.mm662 * 0.34);
    }
  }

  const monthlyDebt = indicatorsRaw.MONTHLY_DEBT_COST || 10574;
  const liquidity = (portfolioSummary["CASH_TWD"] || 0) + (portfolioSummary["USDT"] || 0) + (portfolioSummary["USDC"] || 0);
  const survivalRunway = monthlyDebt > 0 ? (liquidity / monthlyDebt) : 99;

  market.surplus = liquidity - (monthlyDebt * 6); // Keep 6 months buffer for surplus check
  market.btcDrawdownFromATH = calculateBtcDrawdownFromATH_(market.btcPrice, market.sapBaseATH);

  // Phase 4: 自動質押引擎與 active crypto LTV
  const pledgeGroups = calculateAutoPledgeRatios(rawPortfolio, indicatorsRaw);
  let activePledgedCryptoAssets = 0;
  let totalCryptoDebt = 0;
  pledgeGroups.filter(isActiveCryptoPledgeGroup_).forEach(g => {
    activePledgedCryptoAssets += g.collateralValue;
    totalCryptoDebt += g.loanAmount;
  });

  const activeCryptoLTV = activePledgedCryptoAssets > 0 ? (totalCryptoDebt / activePledgedCryptoAssets) : 0;
  market.btcRegime = getBtcRegime_(market.btcMM, market.btcDrawdownFromATH, activeCryptoLTV, survivalRunway);
  const btcAllocationTargets = getBtcAllocationTargets_(market.btcRegime);

  // Phase 4.5: Playbook v3.0 比特幣長牛五部曲戰略狀態與動態滿額指標
  const btcSpotVal = portfolioSummary["BTC_Spot"] || portfolioSummary["BTC"] || 0;
  const ibitVal = portfolioSummary["IBIT"] || 0;
  const totalBtcExposureTwd = btcSpotVal + ibitVal;
  const totalBtcEquivalent = (market.btcPrice > 0 && market.usdTwdRate > 0)
    ? (totalBtcExposureTwd / (market.btcPrice * market.usdTwdRate))
    : 0.948;
  const foreignCashTwd = portfolioSummary["CASH_FC"] || portfolioSummary["USDT"] || 0;
  const extFixedBtc = (market.btcPrice > 0 && market.usdTwdRate > 0)
    ? (ibitVal / (market.btcPrice * market.usdTwdRate))
    : 0.502;

  market.bullStrategy = calculateBullStrategyV3_(
    totalBtcEquivalent,
    market.btcPrice,
    market.btcMM,
    activeCryptoLTV,
    totalCryptoDebt,
    foreignCashTwd,
    extFixedBtc
  );


  // Phase 5: 動態資產配置目標注入與 Layer 4 自動化
  const knownTickers = new Set();
  const assetGroups = Config.ASSET_GROUPS.map(group => {
    group.tickers.forEach(t => knownTickers.add(t));
    let dynamicTarget = group.defaultTarget;

    // Priority 1: Manual override from Sheet
    const key = "Alloc_" + group.id + "_Target";
    if (indicatorsRaw[key] !== undefined && !isNaN(indicatorsRaw[key])) {
      dynamicTarget = indicatorsRaw[key];
    }
    // Priority 2: Auto-calculate from BTC dual-factor regime
    else if (btcAllocationTargets && btcAllocationTargets[group.id] !== undefined) {
      dynamicTarget = btcAllocationTargets[group.id];
    }

    return { ...group, target: dynamicTarget, value: calculateGroupValue(portfolioSummary, group) };
  });

  // 識別雜項資產 (Layer 4)
  const miscTickers = Object.keys(portfolioSummary).filter(t => !knownTickers.has(t) && !Config.NOISE_ASSETS.includes(t) && portfolioSummary[t] > 0);
  const miscValue = miscTickers.reduce((sum, t) => sum + portfolioSummary[t], 0);
  let l4Target = Config.STRATEGIC.L4_ALLOWED_TARGET || 0;
  if (indicatorsRaw["Alloc_L4_Target"] !== undefined && !isNaN(indicatorsRaw["Alloc_L4_Target"])) {
    l4Target = indicatorsRaw["Alloc_L4_Target"];
  }
  const l4CurrentWeight = totalGrossAssets > 0 ? (miscValue / totalGrossAssets) : 0;

  assetGroups.push({
    id: "L4",
    name: "Layer 4: Miscellaneous (To Clear)",
    target: l4Target,
    tickers: miscTickers,
    value: miscValue,
    currentWeight: l4CurrentWeight,
    isMisc: true
  });

  // Phase 6: 再平衡目標
  let targets = enrichRebalanceTargets_(
    getRebalanceTargets(assetGroups, totalGrossAssets, market),
    portfolioSummary,
    assetGroups
  );
  targets = applyStockSettlementGateToRebalanceTargets_(targets, stockStrategy);

  const indicators = {
    isValid: pledgeGroups.length > 0,
    maintenanceRatio: (pledgeGroups.find(g => g.name === "Pledge") || pledgeGroups[0] || { ratio: 0 }).ratio,
    binanceMaintenanceRatio: (pledgeGroups.find(g => g.name === "Binance") || { ratio: 0 }).ratio,
    l1SpotRatio: totalGrossAssets > 0 ? (assetGroups[0].value / totalGrossAssets) : 0,
    totalBtcRatio: totalGrossAssets > 0 ? (assetGroups[0].value / totalGrossAssets) : 0,
    survivalRunway: survivalRunway,
    ltv: totalGrossAssets > 0 ? (totalGrossAssets - netEntityValue) / totalGrossAssets : 0
  };

  // Also keep Global Crypto LTV for macro view
  const l1Value = assetGroups.find(g => g.id === "L1")?.value || 0;
  const l4Value = assetGroups.find(g => g.id === "L4")?.value || 0;
  const totalCryptoAssets = l1Value + l4Value;
  const globalCryptoLTV = totalCryptoAssets > 0 ? (totalCryptoDebt / totalCryptoAssets) : 0;

  indicators.cryptoLTV = activeCryptoLTV; // Use Active LTV as default indicator
  indicators.globalCryptoLTV = globalCryptoLTV;
  indicators.stockPledgeRatio = (pledgeGroups.find(g => g.name.toLowerCase().includes("stock")) || { ratio: 999 }).ratio;

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
    reserve: liquidity,
    inventoryExport: inventoryExport,
    stockStrategy: stockStrategy
  };
}

function fetchMarketIndicators(sheetName) {
    return KeyMarketIndicatorsViewRepo.readIndicators(
      SpreadsheetApp.getActiveSpreadsheet(),
      { sheetName: sheetName }
    );
  }

function getPortfolioData(sheetName) {
    return BalanceSheetViewRepo.readPortfolio(
      SpreadsheetApp.getActiveSpreadsheet(),
      { sheetName: sheetName }
    );
  }
