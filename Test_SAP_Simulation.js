/**
 * Test_SAP_Simulation.js
 * 
 * HOW TO RUN:
 * 1. Open Apps Script editor
 * 2. Select 'debugSAPLogic' function
 * 3. Click 'Run' to simulate strategy logic without executing trades.
 * 
 * EXPECTED OUTPUT:
 * Console logs showing scenario context, decision logic, and hypothetical actions.
 * Logs output to GAS Console.
 * 
 * USAGE: Uncomment and run debugSAPLogic() in Apps Script editor
 * NOTE: This script references RULES from Core_StrategicEngine.js
 *       Ensure Core_StrategicEngine.js is loaded before running tests
 */

// Mock Dependencies
const MOCK_DATA = {
  BALANCE_SHEET: [
    ["Ticker", "Amount", "Value_TWD", "Purpose"],
    ["BTC_Spot", 0.5, 1575000, "Layer 1: Digital Reserve"],
    ["00713", 1000, 50000, "Layer 2: Credit Base"],
    ["USDT", 2000, 65000, "Layer 3: Tactical Liquidity"],
    ["ETH", 2, 100000, "Speculation"]
  ],
  INDICATORS: {
    "SAP_Base_ATH": 70000,
    "Current_BTC_Price": 65000,
    "Total_Martingale_Spent": 50000,
    "MAX_MARTINGALE_BUDGET": 500000,
    "L1_Spot_Ratio": 0.50, // Mock Low
    "Total_BTC_Ratio": 0.40,
    "MONTHLY_DEBT_COST": 15000
  }
};

function debugSAPLogic() {
  Logger.log("=== SAP v24.5 邏輯模擬測試開始 ===");

  // Scenario 1: Normal
  runScenario("情境 A: 正常市場 (平盤震盪)", { btcPrice: 65000, baseAth: 70000 });

  // Scenario 2: ATH Breakout
  runScenario("情境 B: 突破新高 (ATH Detected)", { btcPrice: 75000, baseAth: 70000 });

  // Scenario 3: Sniper Trigger
  runScenario("情境 C: 暴跌狙擊 (Sniper Level 1: -40%)", { btcPrice: 42000, baseAth: 70000 }); // 42k is -40% of 70k

  // Scenario 4: Cashflow Rerouting (L1 Low)
  runScenario("情境 D: 盈餘分配 (現貨不足)", {
    btcPrice: 65000,
    baseAth: 70000,
    l1Ratio: 0.50,
    surplus: 100000
  });

  // Scenario 5: Cashflow Rerouting (Overheated)
  runScenario("情境 E: 盈餘分配 (過熱防禦)", {
    btcPrice: 65000,
    baseAth: 70000,
    l1Ratio: 0.85,
    totalBtcRatio: 0.85,
    surplus: 100000
  });

  Logger.log("=== 測試結束 ===");
}

function runScenario(name, overrides) {
  Logger.log("\n--- " + name + " ---");

  // Build Mock Context
  const context = buildMockContext(overrides);

  let triggered = false;
  RULES.forEach(rule => {
    if (rule.condition(context)) {
      const result = rule.getAction(context);
      if (result) {
        triggered = true;
        Logger.log("[觸發規則] " + rule.name);
        Logger.log("   > 等級: " + result.level);
        Logger.log("   > 訊息: " + result.message);
        Logger.log("   > 行動: " + result.action);
      }
    }
  });

  if (!triggered) {
    Logger.log("[靜默] 無觸發任何警報或行動。");
  }
}

function buildMockContext(overrides) {
  const market = {
    btcPrice: overrides.btcPrice || 60000,
    sapBaseATH: overrides.baseAth || 70000,
    totalMartingaleSpent: overrides.spent || 0,
    maxMartingaleBudget: 500000,
    surplus: overrides.surplus || 0
  };

  const indicators = {
    isValid: true,
    maintenanceRatio: 3.0,
    binanceMaintenanceRatio: 2.5,
    l1SpotRatio: overrides.l1Ratio !== undefined ? overrides.l1Ratio : 0.65,
    totalBtcRatio: overrides.totalBtcRatio !== undefined ? overrides.totalBtcRatio : 0.65,
    survivalRunway: 12,
    ltv: 0.2
  };

  return {
    market,
    indicators,
    phase: "Simulation",
    rebalanceTargets: [],
    netEntityValue: 2000000,
    totalGrossAssets: 3000000 // Mock Total
  };
}