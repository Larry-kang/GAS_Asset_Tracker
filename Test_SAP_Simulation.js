/**
 * Test_SAP_Simulation.js
 * 獨立驗證腳本：用於模擬各種市場情境，測試 SAP v24.5 邏輯反應。
 * 不會讀取真實試算表，不會發送 Email。
 * 
 * 使用方式：在 GAS 編輯器中選擇 'debugSAPLogic' 函式並執行。
 */

function debugSAPLogic() {
    Logger.log("=== SAP v24.5 邏輯模擬測試開始 ===");

    // 1. 測試情境：正常市場 (Normal)
    runScenario("情境 A: 正常市場 (平盤震盪)", {
        btcPrice: 60000,
        sapBaseATH: 70000,
        totalMartingaleSpent: 0,
        maxMartingaleBudget: 437000,
        indicators: {
            l1SpotRatio: 0.65, // Healthy
            totalBtcRatio: 0.70,
            survivalRunway: 35.0,
            maintenanceRatio: 3.0,
            isValid: true
        }
    });

    // 2. 測試情境：ATH 突破 (New ATH)
    runScenario("情境 B: 突破新高 (ATH Detected)", {
        btcPrice: 75000,
        sapBaseATH: 70000, // 70k * 1.05 = 73500. Current 75k > 73.5k. Should trigger.
        totalMartingaleSpent: 0,
        maxMartingaleBudget: 437000,
        indicators: {
            l1SpotRatio: 0.65,
            totalBtcRatio: 0.75,
            survivalRunway: 35.0,
            maintenanceRatio: 3.5,
            isValid: true
        }
    });

    // 3. 測試情境：狙擊手觸發 (Sniper Level 1)
    runScenario("情境 C: 暴跌狙擊 (Sniper Level 1: -40%)", {
        btcPrice: 42000, // 70k * 0.6 = 42k. Exactly -40%.
        sapBaseATH: 70000,
        totalMartingaleSpent: 0,
        maxMartingaleBudget: 437000,
        indicators: {
            l1SpotRatio: 0.65,
            totalBtcRatio: 0.60,
            survivalRunway: 30.0,
            maintenanceRatio: 2.5,
            isValid: true
        }
    });

    // 4. 測試情境：現金流重導向 - 補現貨 (Need Spot)
    runScenario("情境 D: 盈餘分配 (現貨不足)", {
        btcPrice: 60000,
        sapBaseATH: 70000,
        totalMartingaleSpent: 0,
        maxMartingaleBudget: 437000,
        surplus: 50000, // Has cash
        indicators: {
            l1SpotRatio: 0.50, // < 60% Target
            totalBtcRatio: 0.55,
            survivalRunway: 40.0,
            maintenanceRatio: 3.0,
            isValid: true
        }
    });

    // 5. 測試情境：現金流重導向 - 護城河 (Overheated)
    runScenario("情境 E: 盈餘分配 (過熱防禦)", {
        btcPrice: 60000,
        sapBaseATH: 70000,
        totalMartingaleSpent: 0,
        maxMartingaleBudget: 437000,
        surplus: 50000,
        indicators: {
            l1SpotRatio: 0.65,
            totalBtcRatio: 0.85, // > 80%
            survivalRunway: 40.0,
            maintenanceRatio: 3.0,
            isValid: true
        }
    });

    Logger.log("=== 測試結束 ===");
}

function runScenario(name, data) {
    Logger.log("\n--- " + name + " ---");

    // 模擬 Context 物件
    const mockContext = {
        market: {
            btcPrice: data.btcPrice,
            sapBaseATH: data.sapBaseATH,
            totalMartingaleSpent: data.totalMartingaleSpent,
            maxMartingaleBudget: data.maxMartingaleBudget,
            surplus: data.surplus || 0,
            usdTwdRate: 32.5
        },
        indicators: {
            l1SpotRatio: data.indicators.l1SpotRatio,
            totalBtcRatio: data.indicators.totalBtcRatio,
            survivalRunway: data.indicators.survivalRunway,
            maintenanceRatio: data.indicators.maintenanceRatio,
            binanceMaintenanceRatio: 3.0,
            isValid: data.indicators.isValid,
            ltv: 0.3
        },
        rebalanceTargets: [], // default empty
        phase: "Simulation",
        portfolioSummary: {},
        pledgeGroups: [],
        netEntityValue: 1000000,
        totalGrossAssets: 1500000,
        reserve: 100000
    };

    // 執行規則檢查
    let triggered = false;
    RULES.forEach(rule => {
        try {
            if (rule.condition(mockContext)) {
                const result = rule.getAction(mockContext);
                if (result) {
                    Logger.log(`[觸發規則] ${rule.name}`);
                    Logger.log(`   > 等級: ${result.level}`);
                    Logger.log(`   > 訊息: ${result.message}`);
                    Logger.log(`   > 行動: ${result.action}`);
                    triggered = true;
                }
            }
        } catch (e) {
            Logger.log(`[錯誤] 規則 ${rule.name} 執行失敗: ${e.message}`);
        }
    });

    if (!triggered) {
        Logger.log("[靜默] 無觸發任何警報或行動。");
    }
}
