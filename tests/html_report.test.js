if (typeof require === 'function') {
  const test = require('node:test');
  const assert = require('node:assert/strict');
  const fs = require('node:fs');
  const path = require('node:path');
  const vm = require('node:vm');

  function loadHtmlReportContext() {
    const repoRoot = path.resolve(__dirname, '..');
    const sandbox = {
      JSON,
      Date,
      Math,
      parseFloat,
      parseInt,
      isFinite,
      isNaN,
      console,
      Logger: { log() {} },
      LogService: { info() {}, warn() {}, error() {} },
      Settings: { get() { return ''; } },
      SpreadsheetApp: {
        getActiveSpreadsheet() {
          return {
            getSheetByName() { return null; }
          };
        }
      }
    };
    const context = vm.createContext(sandbox);
    const files = [
      'Config.js',
      'Strategy_BullStateMachine.js',
      'Strategy_StockExposure.js',
      'Strategy_RebalanceEngine.js',
      'Strategy_ReportFormatter.js',
      'Util_HtmlReport.js'
    ];
    files.forEach(file => {
      const filePath = path.join(repoRoot, file);
      if (fs.existsSync(filePath)) {
        vm.runInContext(fs.readFileSync(filePath, 'utf8'), context, { filename: file });
      }
    });
    return context;
  }

  function createMockContext() {
    return {
      phase: "Bitcoin Standard v24.14",
      totalGrossAssets: 3030226,
      netEntityValue: 1709120,
      market: {
        btcPrice: 77264,
        sapBaseATH: 88000,
        btcDrawdownFromATH: -0.122,
        btcMM: 1.1146,
        btcRegime: {
          regime: 'ACCUMULATE',
          phaseLabel: '長牛五部曲右側擴張',
          reason: 'MM 處於健康牛市蓄勢區間',
          restockMode: 'NORMAL',
          targetLtvMin: 0.40,
          targetLtvMax: 0.45
        },
        bullStrategy: {
          phaseLabel: '長牛擴張期',
          recommendedDca: '2U / 天',
          dynamic50TargetBtc: '0.948',
          exitRoadmap: '分批止盈'
        },
        twWeightedMM: 1.05
      },
      indicators: {
        survivalRunway: 13.82,
        ltv: 0.436,
        cryptoLTV: 0.2669,
        globalCryptoLTV: 0.436
      },
      pledgeGroups: [
        { name: "OKX Loan", ratio: 3.74, critical: 1.5, alert: 2.0, loanAmount: 294976, collateralValue: 1105000 }
      ],
      assetGroups: [
        { id: "L1", name: "Layer 1: 比特幣核心", value: 1763500, target: 0.60, defaultTarget: 0.60, isMisc: false },
        { id: "L2", name: "Layer 2: 股票正二曝險", value: 742400, target: 0.25, defaultTarget: 0.25, isMisc: false },
        { id: "L3", name: "Layer 3: 流動性儲備", value: 448500, target: 0.15, defaultTarget: 0.15, isMisc: false },
        { id: "L4", name: "Layer 4: 雜項部位", value: 75826, target: 0, defaultTarget: 0, isMisc: true, tickers: ['SHIB', 'DOGE'] }
      ],
      rebalanceTargets: [
        { shortName: "L4 雜項", action: "CLEAR", currentWeight: 0.025, targetWeight: 0, deltaValue: -75826, executionHint: "市價清理", suggestedFundingSource: "轉入流動性" },
        { shortName: "L1 BTC", action: "ADD", currentWeight: 0.582, targetWeight: 0.60, deltaValue: 54600, executionHint: "分批買進", suggestedFundingSource: "L4 資金" }
      ],
      portfolioSummary: {
        "BTC": 1763500,
        "00670L": 400000,
        "00685L": 342400,
        "CASH_TWD": 159946,
        "USDT": 288554
      },
      stockStrategy: null
    };
  }

  test('generatePortfolioSnapshotHtml renders complete HTML with critical sections', () => {
    const context = loadHtmlReportContext();
    assert.equal(typeof context.generatePortfolioSnapshotHtml, 'function');

    const mockCtx = createMockContext();
    const alerts = [{ level: 'WARNING', message: 'L4 雜項部位偏高', action: '清理 L4' }];
    const html = context.generatePortfolioSnapshotHtml(mockCtx, alerts, { isEmail: true });

    assert.ok(html.includes('SAP 戰略指揮中心'));
    assert.ok(html.includes('77,264'));
    assert.ok(html.includes('1,709,120'));
    assert.ok(html.includes('13.8'));
    assert.ok(html.includes('26.7%') || html.includes('26.69%') || html.includes('26.7'));
    assert.ok(html.includes('Layer 1'));
    assert.ok(html.includes('再平衡建議'));
    assert.ok(html.includes('L4 雜項部位偏高'));
  });

  test('generatePortfolioSnapshotHtml distinguishes email and interactive mode', () => {
    const context = loadHtmlReportContext();
    const mockCtx = createMockContext();

    const emailHtml = context.generatePortfolioSnapshotHtml(mockCtx, [], { isEmail: true });
    assert.ok(!emailHtml.includes('google.script.run'));
    assert.ok(!emailHtml.includes('id="btn-sync"'));

    const interactiveHtml = context.generatePortfolioSnapshotHtml(mockCtx, [], { isInteractive: true });
    assert.ok(interactiveHtml.includes('id="btn-sync"'));
    assert.ok(interactiveHtml.includes('google.script.run'));
    assert.ok(interactiveHtml.includes('runAutomationMaster'));
  });

  test('generatePortfolioSnapshotHtml handles empty or boundary contexts gracefully', () => {
    const context = loadHtmlReportContext();
    const minimalCtx = {
      market: {},
      indicators: {},
      pledgeGroups: [],
      assetGroups: [],
      rebalanceTargets: [],
      portfolioSummary: {},
      totalGrossAssets: 0,
      netEntityValue: 0
    };

    const html = context.generatePortfolioSnapshotHtml(minimalCtx, [], { isEmail: true });
    assert.ok(typeof html === 'string');
    assert.ok(html.length > 100);
    assert.ok(html.includes('SAP 戰略指揮中心'));
  });
}
