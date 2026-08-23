/**
 * Strategy_StockExposure.js
 * Sovereign Asset Protocol - Stock Nominal Exposure & Settlement Gate
 */

function buildStockExposureStrategy_(inventoryExport) {
  if (!inventoryExport || !inventoryExport.available) return null;

  const summary = inventoryExport.summary || {};
  const rawPositions = Array.isArray(inventoryExport.positions) ? inventoryExport.positions : [];
  const positions = rawPositions
    .filter(function (position) {
      const region = String((position && position.region) || "").trim().toUpperCase();
      const role = String((position && position.strategyRole) || "").trim();
      return (region === "TW" || region === "NASDAQ") && role !== "";
    })
    .map(function (position) {
      const marketValue = toFiniteNumber_(position.valueTwd) || 0;
      const multiplier = toFiniteNumber_(position.exposureMultiplier) || 1;
      const exportedExposure = toFiniteNumber_(position.effectiveExposureTwd);
      return {
        ticker: String(position.ticker || "").trim(),
        quantity: toFiniteNumber_(position.quantity) || 0,
        marketValueTwd: marketValue,
        exposureMultiplier: multiplier,
        effectiveExposureTwd: exportedExposure === null ? marketValue * multiplier : exportedExposure,
        region: String(position.region || "").trim().toUpperCase(),
        strategyRole: String(position.strategyRole || "").trim(),
        settlementStatus: String(position.settlementStatus || "UNKNOWN").trim().toUpperCase(),
        pledgeStatus: String(position.pledgeStatus || "UNKNOWN").trim().toUpperCase()
      };
    });

  if (positions.length === 0) return null;

  const marketValueTwd = positions.reduce(function (sum, position) {
    return sum + position.marketValueTwd;
  }, 0);
  const grossExposureTwd = positions.reduce(function (sum, position) {
    return sum + position.effectiveExposureTwd;
  }, 0);
  const taiwanExposureTwd = positions
    .filter(function (position) { return position.region === "TW"; })
    .reduce(function (sum, position) { return sum + position.effectiveExposureTwd; }, 0);
  const nasdaqExposureTwd = positions
    .filter(function (position) { return position.region === "NASDAQ"; })
    .reduce(function (sum, position) { return sum + position.effectiveExposureTwd; }, 0);
  const pendingPosition = positions.find(function (position) {
    return position.settlementStatus !== "SETTLED";
  });
  const stockCapitalBaseTwd = toFiniteNumber_(summary.StockCapitalBaseTWD);
  const targetExposureRatio = toFiniteNumber_(summary.StockTargetExposureRatio) || 1;
  const targetTaiwanRatio = toFiniteNumber_(summary.TaiwanTargetExposureWeight) || 0.50;
  const targetNasdaqRatio = toFiniteNumber_(summary.NasdaqTargetExposureWeight) || 0.50;
  const taiwanRatio = grossExposureTwd > 0 ? taiwanExposureTwd / grossExposureTwd : 0;
  const nasdaqRatio = grossExposureTwd > 0 ? nasdaqExposureTwd / grossExposureTwd : 0;
  const debtStatus = String(summary.StockDebtStatus || "UNKNOWN").trim().toUpperCase();
  const settlementStatus = pendingPosition ? pendingPosition.settlementStatus : "SETTLED";

  return {
    source: "inventory_export",
    positions: positions,
    marketValueTwd: marketValueTwd,
    grossExposureTwd: grossExposureTwd,
    stockCapitalBaseTwd: stockCapitalBaseTwd !== null && stockCapitalBaseTwd > 0
      ? stockCapitalBaseTwd
      : null,
    exposureRatio: stockCapitalBaseTwd !== null && stockCapitalBaseTwd > 0
      ? grossExposureTwd / stockCapitalBaseTwd
      : null,
    targetExposureRatio: targetExposureRatio,
    taiwanExposureTwd: taiwanExposureTwd,
    nasdaqExposureTwd: nasdaqExposureTwd,
    taiwanExposureRatio: taiwanRatio,
    nasdaqExposureRatio: nasdaqRatio,
    targetTaiwanExposureRatio: targetTaiwanRatio,
    targetNasdaqExposureRatio: targetNasdaqRatio,
    taiwanDeviationPct: (taiwanRatio - targetTaiwanRatio) * 100,
    nasdaqDeviationPct: (nasdaqRatio - targetNasdaqRatio) * 100,
    cashBufferTwd: toFiniteNumber_(summary.CashBufferTWD) || 0,
    debtStatus: debtStatus,
    settlementStatus: settlementStatus,
    isPending: debtStatus !== "SETTLED" || settlementStatus !== "SETTLED",
    originalCoreTicker: String(summary.OriginalCoreTicker || "").trim(),
    originalCoreMinPreSplitEquivalentQty:
      toFiniteNumber_(summary.OriginalCoreMinPreSplitEquivalentQty) || 0,
    corporateActionStatus:
      String(summary.CorporateAction00662Status || "UNKNOWN").trim().toUpperCase()
  };
}

function buildStockExposureAlert_(stockStrategy) {
  if (!stockStrategy) return null;

  const marketValue = Math.round(stockStrategy.marketValueTwd).toLocaleString();
  const grossExposure = Math.round(stockStrategy.grossExposureTwd).toLocaleString();
  const regionMessage =
    "台灣 " + formatPercent_(stockStrategy.taiwanExposureRatio, 1) +
    " / NASDAQ " + formatPercent_(stockStrategy.nasdaqExposureRatio, 1);

  if (stockStrategy.isPending) {
    return {
      level: "[結算] 股票重整待完成",
      message:
        "股票市值 " + marketValue + " TWD，名目曝險 " + grossExposure +
        " TWD；" + regionMessage + "。交割 " + stockStrategy.settlementStatus +
        " / 股票負債 " + stockStrategy.debtStatus + "。",
      action:
        "暫停新增股票操作與股票質押；等待 T+2、實際自由現金及股票負債歸零後再評估區域微調。"
    };
  }

  const regionalDrift = Math.max(
    Math.abs(stockStrategy.taiwanDeviationPct),
    Math.abs(stockStrategy.nasdaqDeviationPct)
  );
  let action = "維持目前名目曝險，不新增股票質押。";
  if (regionalDrift >= 5) {
    action =
      "區域偏差超過 5 個百分點；待 " + (stockStrategy.originalCoreTicker || "原型 ETF") +
      " 公司行動完成後，再以新增資金或小額換倉向目標比例調整。";
  }

  return {
    level: "[戰略] 股票有效曝險監控",
    message:
      "股票市值 " + marketValue + " TWD，名目曝險 " + grossExposure +
      " TWD；" + regionMessage + "。",
    action: action
  };
}

function buildStockExposureSnapshot_(stockStrategy) {
  if (!stockStrategy) return "";

  let s = "\n[股票有效曝險]\n";
  stockStrategy.positions.forEach(function (position) {
    s +=
      "- " + position.ticker + ": " + position.quantity.toLocaleString() +
      " 股 | 市值 " + Math.round(position.marketValueTwd).toLocaleString() +
      " | " + position.exposureMultiplier.toFixed(1) + "x => " +
      Math.round(position.effectiveExposureTwd).toLocaleString() + " TWD\n";
  });
  s += "- 股票市值: " + Math.round(stockStrategy.marketValueTwd).toLocaleString() + " TWD\n";
  s += "- 名目曝險: " + Math.round(stockStrategy.grossExposureTwd).toLocaleString() + " TWD\n";
  if (stockStrategy.exposureRatio === null) {
    s += "- 曝險倍數: N/A（Strategy_Config 缺少 StockCapitalBaseTWD）\n";
  } else {
    s +=
      "- 曝險倍數: " + stockStrategy.exposureRatio.toFixed(3) + "x" +
      " / 目標 " + stockStrategy.targetExposureRatio.toFixed(2) + "x\n";
  }
  s +=
    "- 區域: 台灣 " + formatPercent_(stockStrategy.taiwanExposureRatio, 2) +
    "（目標 " + formatPercent_(stockStrategy.targetTaiwanExposureRatio, 0) +
    "，偏差 " + stockStrategy.taiwanDeviationPct.toFixed(2) + "pp）" +
    " | NASDAQ " + formatPercent_(stockStrategy.nasdaqExposureRatio, 2) +
    "（目標 " + formatPercent_(stockStrategy.targetNasdaqExposureRatio, 0) +
    "，偏差 " + stockStrategy.nasdaqDeviationPct.toFixed(2) + "pp）\n";
  s +=
    "- 狀態: 交割 " + stockStrategy.settlementStatus +
    " | 股票負債 " + stockStrategy.debtStatus +
    " | 現金底線 " + Math.round(stockStrategy.cashBufferTwd).toLocaleString() + " TWD\n";
  if (stockStrategy.originalCoreTicker) {
    s +=
      "- 原型核心: " + stockStrategy.originalCoreTicker +
      "，最低保留分割前等值 " +
      stockStrategy.originalCoreMinPreSplitEquivalentQty.toLocaleString() +
      " 股 | 公司行動 " + stockStrategy.corporateActionStatus + "\n";
  }
  return s;
}

function applyStockSettlementGateToRebalanceTargets_(targets, stockStrategy) {
  const list = Array.isArray(targets) ? targets : [];
  if (!stockStrategy || !stockStrategy.isPending) return list;
  return list.filter(function (target) {
    return target && target.id === "L4" && target.action === "CLEAR";
  });
}

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
      return Config.BTC_MARTINGALE.ENABLED &&
        context.market.sapBaseATH > 0 &&
        context.market.totalMartingaleSpent < context.market.maxMartingaleBudget;
    },
    getAction: function (context) {
      const currentDrop = (context.market.btcPrice - context.market.sapBaseATH) / context.market.sapBaseATH;
      const strategy = Config.BTC_MARTINGALE;

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

        let actionText = "執行買入: TWD " + estCost.toLocaleString() + " 等值 BTC。\n(執行後請手動更新 `Total_Martingale_Spent` += " + estCost + ")";
        if (context.stockStrategy && context.stockStrategy.isPending) {
          actionText =
            "狙擊訊號成立，但股票重整與質押清償尚未完成。\n" +
            "不執行加速買入；等待 T+2、實際自由現金與股票負債歸零，只保留既有固定 DCA。";
        } else if (shouldSuppressAggressiveBtcBuying_(context.market.btcRegime)) {
          actionText = "狙擊訊號成立，但目前 BTC Regime 為 " + context.market.btcRegime.regime +
            " / Restock Mode: " + context.market.btcRegime.restockMode + "。\n" +
            (buildCryptoLtvGuardrailAction_(context.market.btcRegime) || "暫不執行戰術買入，只保留固定 DCA 或手動檢查。");
        }

        return {
          level: "[攻擊] 狙擊信號 (Sniper)",
          message: "BTC 回調 " + (currentDrop * 100).toFixed(1) + "% (基準: $" + context.market.sapBaseATH + "). 進入 " + activeLevel.name,
          action: actionText
        };
      }
      return null;
    }
  },
  {
    name: "Crypto LTV Guardrail",
    phase: "All",
    condition: function (context) {
      return context.indicators && context.indicators.cryptoLTV >= 0.45;
    },
    getAction: function (context) {
      const ltv = context.indicators.cryptoLTV;
      const btcRegime = context.market.btcRegime || getBtcRegime_(context.market.btcMM, context.market.btcDrawdownFromATH, ltv, context.indicators.survivalRunway);
      let level = "[注意] Crypto LTV Guardrail";
      let action = "只允許用現金流 DCA，不新增質押借款。";

      if (ltv >= 0.60) {
        level = "[嚴重] Crypto LTV 硬上限突破";
        action = "強制去槓桿，停止所有買入。";
      } else if (ltv >= 0.55) {
        level = "[嚴重] DEFCON 1 Crypto LTV";
        action = "停止所有買入，優先還款降槓桿。";
      } else if (ltv >= 0.50) {
        level = "[警告] Crypto LTV Stretch 關閉";
        action = "停止新增借款，優先還款或只保留必要現金流 DCA。";
      } else if (ltv >= 0.45) {
        level = "[注意] Crypto LTV Active 上緣";
        action = "暫停提高 OKX 槓桿，只保留 Active cap 以內操作。";
      }

      return {
        level: level,
        message: "Active Crypto LTV 目前為 " + formatPercent_(ltv, 1) + "；BTC Regime: " + btcRegime.regime,
        action: btcRegime.guardrailAction || action
      };
    }
  },
  {
    name: "Cashflow Rerouting Engine",
    phase: "All",
    condition: function (context) {
      return context.market.surplus > 0 || context.rebalanceTargets.length > 0;
    },
    getAction: function (context) {
      if (context.stockStrategy && context.stockStrategy.isPending) {
        return null;
      }

      const topTarget = (context.rebalanceTargets || [])[0];
      if (topTarget) {
        return buildRebalanceAlert_(topTarget);
      }

      // Priority 1: Check L1 Spot Ratio
      const l1Target = context.assetGroups ? context.assetGroups[0].target : 0.60;
      if (context.indicators.l1SpotRatio < l1Target) {
        const btcRegime = context.market.btcRegime;
        if (shouldSuppressAggressiveBtcBuying_(btcRegime)) {
          return {
            level: "[風控] 資金流向暫停補強 BTC",
            message: "L1 現貨佔比 (" + (context.indicators.l1SpotRatio * 100).toFixed(1) + "%) 低於 " + (l1Target * 100).toFixed(0) + "%，但 BTC Regime 為 " + btcRegime.regime + " / " + btcRegime.restockMode + "。",
            action: buildCryptoLtvGuardrailAction_(btcRegime) || "只維持基本 DCA，不做一次性 BTC restock。"
          };
        }
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
          action: context.stockStrategy
            ? "先保留現金；僅在股票名目曝險低於目標且結算完成時，依台灣/NASDAQ 區域缺口補足。"
            : "將盈餘轉入股票信用基底或流動性，避免繼續提高 BTC 集中度。"
        };
      }

      return null;
    }
  },
  {
    name: "Maintenance Ratio Monitor",
    phase: "All",
    condition: function (context) {
      return !context.stockStrategy &&
        context.indicators.isValid &&
        context.indicators.maintenanceRatio < Config.STRATEGIC.PLEDGE_RATIO_SAFE;
    },
    getAction: function (context) {
      const ratio = context.indicators.maintenanceRatio;
      if (ratio <= Config.STRATEGIC.PLEDGE_RATIO_CRITICAL) {
        return {
          level: "[嚴重] 斷頭追繳警報 (1.8)",
          message: "維持率崩跌至 " + ratio.toFixed(2),
          action: "執行焦土防禦: 強制清算所有雜訊資產 (ETH/BNB/TQQQ) 以償還債務。禁止買入。"
        };
      } else if (ratio <= Config.STRATEGIC.PLEDGE_RATIO_ALERT) {
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
    condition: function (context) { return context.indicators.binanceMaintenanceRatio > 0 && context.indicators.binanceMaintenanceRatio < Config.STRATEGIC.CRYPTO_LOAN_RATIO_SAFE; },
    getAction: function (context) {
      const ratio = context.indicators.binanceMaintenanceRatio;
      if (ratio <= Config.STRATEGIC.CRYPTO_LOAN_RATIO_CRITICAL) {
        return {
          level: "[嚴重] 幣安保證金保護",
          message: "幣安質押率 (BTC/USDT) 崩跌至 " + ratio.toFixed(2) + " (LTV " + (1 / ratio * 100).toFixed(1) + "%)",
          action: "立即行動: 補倉 BTC 或償還幣安貸款以避免清算。"
        };
      } else if (ratio <= Config.STRATEGIC.CRYPTO_LOAN_RATIO_ALERT) {
        return {
          level: "[警告] 幣安風險區",
          message: "幣安質押率在 " + ratio.toFixed(2),
          action: "警告: 檢測到 BTC 高波動。準備抵押品或啟動減壓操作。"
        };
      }
      return null;
    }
  },
  {
    name: "OKX Crypto Loan Monitor",
    phase: "All",
    condition: function (context) {
      const okxGroup = context.pledgeGroups.find(g => g.name.toLowerCase().includes("okx"));
      return okxGroup && okxGroup.ratio < (okxGroup.alert || Config.STRATEGIC.CRYPTO_LOAN_RATIO_SAFE);
    },
    getAction: function (context) {
      const okxGroup = context.pledgeGroups.find(g => g.name.toLowerCase().includes("okx"));
      const ratio = okxGroup.ratio;
      if (ratio <= (okxGroup.critical || Config.STRATEGIC.CRYPTO_LOAN_RATIO_CRITICAL)) {
        return {
          level: "[嚴重] OKX 保證金保護",
          message: "OKX 質押率 (BTC/USDT) 崩跌至 " + ratio.toFixed(2) + " (LTV " + (1 / ratio * 100).toFixed(1) + "%)",
          action: "立即行動: 補倉 BTC 或償還 OKX 貸款。優先動用『戰略儲備』部位。"
        };
      } else if (ratio <= (okxGroup.alert || Config.STRATEGIC.CRYPTO_LOAN_RATIO_ALERT)) {
        return {
          level: "[警告] OKX 風險區",
          message: "OKX 質押率降至 " + ratio.toFixed(2),
          action: "警告: OKX 節點壓力增加。若持續下跌請考慮降低槓桿。"
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
      const hasL4RebalanceTarget = (context.rebalanceTargets || []).some(function (target) {
        return target && target.id === "L4";
      });
      return l4 && l4.value > 0 && (l4.currentWeight || 0) > (l4.target || 0) && !hasL4RebalanceTarget;
    },
    getAction: function (context) {
      const l4 = context.assetGroups.find(g => g.id === "L4");
      const currentPct = ((l4.currentWeight || 0) * 100).toFixed(1);
      const allowedPct = ((l4.target || 0) * 100).toFixed(1);
      return {
        level: "[注意] 雜項資產清理建議",
        message: "偵測到 Layer 4 雜項資產: " + l4.tickers.join(", ") + " (總值: " + Math.round(l4.value).toLocaleString() + " TWD, 目前 " + currentPct + "% / 容許 " + allowedPct + "%)",
        action: "建議找市場高位機會清空雜項資產，回歸 L1 (BTC) 或 L2 (穩定基底)。"
      };
    }
  },
  {
    name: "BTC 1.0 Milestone De-escalation Monitor",
    phase: "All",
    condition: function (context) {
      const btcEq = (context.inventoryExport && context.inventoryExport.summary && (context.inventoryExport.summary.BTC_Equivalent || context.inventoryExport.summary.btc_equivalent)) || (context.portfolioSummary && context.portfolioSummary.BTC_Equivalent) || 0;
      return Number(btcEq || 0) >= 1.0;
    },
    getAction: function (context) {
      const btcEq = (context.inventoryExport && context.inventoryExport.summary && (context.inventoryExport.summary.BTC_Equivalent || context.inventoryExport.summary.btc_equivalent)) || (context.portfolioSummary && context.portfolioSummary.BTC_Equivalent) || 0;
      return {
        level: "[里程碑] 1.0 BTC 達標告警",
        message: "BTC-equivalent 已達標 (" + Number(btcEq).toFixed(4) + " BTC >= 1.0000 BTC)。",
        action: "建議啟動第二階段降速協議：將 OKX DCA 降至 2U/天 (60U/月)，由日日生幣被動利息自給自足；釋出之月現金流 (~4.5萬 TWD) 轉向台股正二重整 (目標 100萬 曝險)。"
      };
    }
  },
  {
    name: "Taiwan Stock Leverage Advisor",
    phase: "All",
    condition: function (context) {
      return !!context.stockStrategy || context.market.twWeightedMM !== null;
    },
    getAction: function (context) {
      if (context.stockStrategy) {
        return buildStockExposureAlert_(context.stockStrategy);
      }

      const mm = context.market.twWeightedMM;
      let zone = "", action = "", level = "[戰略] 台股指引";

      if (mm > 1.35) {
        zone = "極度泡沫 (Bubble)";
        action = "停止加碼並檢查股票名目曝險；不新增股票質押。";
      } else if (mm > 1.15) {
        zone = "高位警戒 (Warning)";
        action = "暫停提高股票曝險，等待估值或區域配置回到目標。";
      } else if (mm > 1.00) {
        zone = "中性平衡 (Neutral)";
        action = "維持目前股票曝險；以現金流和區域偏差管理，不使用外部股票槓桿。";
      } else if (mm > 0.85) {
        zone = "低位部屬 (Accumulate)";
        action = "僅用自由現金分批補足低配區域，不新增股票質押。";
      } else {
        zone = "深水炸彈 (Deep Value)";
        action = "先確認現金底線與 Crypto LTV，再用自由現金階梯式提高股票曝險。";
        level = "[機會] 台股黃金坑";
      }

      return {
        level: level + " (" + zone + ")",
        message: "加權 MM: " + mm.toFixed(2) + " (713: " + context.market.twMMParts.mm713.toFixed(2) + " | 662: " + context.market.twMMParts.mm662.toFixed(2) + ")",
        action: action
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