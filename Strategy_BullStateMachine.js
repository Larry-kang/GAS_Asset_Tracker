/**
 * Strategy_BullStateMachine.js
 * Sovereign Asset Protocol - BTC Dual-Factor & Playbook v3.0 State Machine
 */

function toFiniteNumber_(value) {
  const num = parseFloat(value);
  return isFinite(num) ? num : null;
}

function calculateBtcDrawdownFromATH_(btcPrice, sapBaseATH) {
  const price = toFiniteNumber_(btcPrice);
  const ath = toFiniteNumber_(sapBaseATH);
  if (price === null || ath === null || price <= 0 || ath <= 0) return null;
  return (price - ath) / ath;
}

function createBtcRegime_(config) {
  return {
    regime: config.regime,
    phaseLabel: config.phaseLabel,
    action: config.action,
    restockMode: config.restockMode || "OFF",
    restockAllowed: !!config.restockAllowed,
    targetLtvMin: config.targetLtvMin || 0,
    targetLtvMax: config.targetLtvMax || 0,
    allocationBias: config.allocationBias || "MANUAL_REVIEW",
    severity: config.severity || "INFO",
    reason: config.reason || "Manual review required",
    guardrailMessage: config.guardrailMessage || "",
    guardrailAction: config.guardrailAction || ""
  };
}

function applyBtcRestockGate_(regime, cryptoLTV, survivalRunway) {
  const ltv = Math.max(toFiniteNumber_(cryptoLTV) || 0, 0);
  const runway = toFiniteNumber_(survivalRunway);
  const accumulationModes = ["NORMAL", "EXTENDED", "PANIC"];
  const isAccumulationMode = accumulationModes.indexOf(regime.restockMode) >= 0;

  if (!isAccumulationMode) return regime;

  if (ltv >= 0.50) {
    return Object.assign({}, regime, {
      restockAllowed: false,
      restockMode: "OFF",
      severity: "WARN",
      guardrailMessage: "Active Crypto LTV 已達 50%，停止新增借款並關閉 Stretch。",
      guardrailAction: "停止新增借款，優先還款或只保留必要現金流 DCA。"
    });
  }

  if (ltv >= 0.45) {
    return Object.assign({}, regime, {
      restockAllowed: false,
      restockMode: "DCA_ONLY",
      severity: "WARN",
      guardrailMessage: "Active Crypto LTV 已達 45%，只保留 Active cap 以下的保守操作。",
      guardrailAction: "只允許用現金流 DCA，不新增質押借款。"
    });
  }

  if (runway !== null && runway < 6) {
    return Object.assign({}, regime, {
      restockAllowed: false,
      restockMode: "DCA_ONLY",
      severity: "WARN",
      guardrailMessage: "生存跑道低於 6 個月，暫停戰術加碼。",
      guardrailAction: "先補足流動性緩衝，再恢復 restock。"
    });
  }

  return Object.assign({}, regime, { restockAllowed: true });
}

function getBtcRegime_(btcMM, btcDrawdown, cryptoLTV, survivalRunway) {
  const mm = toFiniteNumber_(btcMM);
  const drawdown = toFiniteNumber_(btcDrawdown);
  const ltv = Math.max(toFiniteNumber_(cryptoLTV) || 0, 0);

  if (ltv >= 0.60) {
    return createBtcRegime_({
      regime: "HARD_CAP_BREACH",
      phaseLabel: "Crypto LTV 硬上限突破",
      action: "FORCE_DELEVERAGE",
      restockMode: "DEFCON",
      targetLtvMin: 0,
      targetLtvMax: 0,
      allocationBias: "DEBT_REDUCTION",
      severity: "CRITICAL",
      reason: "Crypto LTV >= 60%",
      guardrailMessage: "Active Crypto LTV 已超過 60% 內部硬紅線。",
      guardrailAction: "強制去槓桿，停止所有買入。"
    });
  }

  if (ltv >= 0.55) {
    return createBtcRegime_({
      regime: "DEFCON_1",
      phaseLabel: "DEFCON 1 去槓桿",
      action: "STOP_BUYING_REPAY_DEBT",
      restockMode: "DEFCON",
      targetLtvMin: 0,
      targetLtvMax: 0,
      allocationBias: "DEBT_REDUCTION",
      severity: "CRITICAL",
      reason: "Crypto LTV >= 55%",
      guardrailMessage: "Active Crypto LTV 已超過 55% 戰術極限。",
      guardrailAction: "停止所有買入，優先還款降槓桿。"
    });
  }

  if (ltv >= 0.50) {
    return createBtcRegime_({
      regime: "STRETCH_LOCK",
      phaseLabel: "Stretch 關閉 / 停止新增借款",
      action: "NO_NEW_BORROWING",
      restockMode: "OFF",
      restockAllowed: false,
      targetLtvMin: 0.45,
      targetLtvMax: 0.50,
      allocationBias: "DEBT_STABILIZE",
      severity: "WARN",
      reason: "Crypto LTV >= 50%",
      guardrailMessage: "Active Crypto LTV 已超過 50%，禁止再往 Stretch 加碼。",
      guardrailAction: "停止新增借款，僅允許現金流 DCA 或優先降槓桿。"
    });
  }

  if (mm === null) {
    return createBtcRegime_({
      regime: "MANUAL_REVIEW",
      phaseLabel: "缺少 BTC_MM，需手動檢查",
      action: "MANUAL_REVIEW",
      restockMode: "OFF",
      targetLtvMin: 0,
      targetLtvMax: 0,
      allocationBias: "MANUAL_REVIEW",
      severity: "WARN",
      reason: "BTC_MM missing"
    });
  }

  if (mm >= 2.10) {
    return createBtcRegime_({
      regime: "BUBBLE",
      phaseLabel: "泡沫 / 高熱區",
      action: "DE_RISK",
      restockMode: "OFF",
      targetLtvMin: 0,
      targetLtvMax: 0,
      allocationBias: "RISK_OFF",
      severity: "WARN",
      reason: "BTC_MM >= 2.10"
    });
  }

  if (mm >= 1.50) {
    return createBtcRegime_({
      regime: "DE_RISK",
      phaseLabel: "去槓桿區",
      action: "REPAY_AND_UNWRAP",
      restockMode: "OFF",
      targetLtvMin: 0,
      targetLtvMax: 0.10,
      allocationBias: "DE_RISK",
      severity: "WARN",
      reason: "BTC_MM >= 1.50"
    });
  }

  if (mm < 0.60 || (drawdown !== null && drawdown <= -0.55)) {
    return applyBtcRestockGate_(createBtcRegime_({
      regime: "PANIC_ACCUMULATE",
      phaseLabel: "恐慌低點累積區",
      action: "PANIC_RESTOCK",
      restockMode: "PANIC",
      targetLtvMin: 0.50,
      targetLtvMax: 0.55,
      allocationBias: "MAX_L1_ACCUMULATE",
      severity: "INFO",
      reason: mm < 0.60 ? "BTC_MM < 0.60" : "Drawdown <= -55%"
    }), ltv, survivalRunway);
  }

  if (mm < 0.75) {
    return applyBtcRestockGate_(createBtcRegime_({
      regime: "DEEP_VALUE",
      phaseLabel: "深度折價主操作區",
      action: "ACTIVE_RESTOCK",
      restockMode: "NORMAL",
      targetLtvMin: 0.45,
      targetLtvMax: 0.50,
      allocationBias: "MAX_L1_ACCUMULATE",
      severity: "INFO",
      reason: "0.60 <= BTC_MM < 0.75"
    }), ltv, survivalRunway);
  }

  if (mm < 1.00) {
    return applyBtcRestockGate_(createBtcRegime_({
      regime: "ACCUMULATE",
      phaseLabel: "低位積累區",
      action: "NORMAL_RESTOCK",
      restockMode: "NORMAL",
      targetLtvMin: 0.40,
      targetLtvMax: 0.45,
      allocationBias: "L1_ACCUMULATE",
      severity: "INFO",
      reason: "0.75 <= BTC_MM < 1.00"
    }), ltv, survivalRunway);
  }

  if (mm < 1.50) {
    return createBtcRegime_({
      regime: "NEUTRAL",
      phaseLabel: "中性區 / 固定 DCA",
      action: "DCA_ONLY",
      restockMode: "DCA_ONLY",
      targetLtvMin: 0.30,
      targetLtvMax: 0.35,
      allocationBias: "NEUTRAL",
      severity: "INFO",
      reason: "BTC_MM between 1.00 and 1.50 without deep drawdown"
    });
  }

  return createBtcRegime_({
    regime: "MANUAL_REVIEW",
    phaseLabel: "未知狀態，需手動檢查",
    action: "MANUAL_REVIEW",
    restockMode: "OFF",
    targetLtvMin: 0,
    targetLtvMax: 0,
    allocationBias: "MANUAL_REVIEW",
    severity: "WARN",
    reason: "No BTC regime matched"
  });
}

function calculateBullStrategyV3_(totalBtc, btcPrice, btcMM, activeCryptoLTV, totalCryptoDebt, binanceU, extFixedBtc) {
  const price = (btcPrice > 0) ? btcPrice : 77000;
  const debtUsdt = (totalCryptoDebt > 0) ? (totalCryptoDebt / 32.5) : 9076;
  const uReserve = (binanceU > 0) ? (binanceU / 32.5) : 1818;
  const extBtc = (extFixedBtc >= 0) ? extFixedBtc : 0.502;
  const okxPledgedBtc = Math.max(totalBtc - extBtc, 0.4455);

  const baseOkxBtc = okxPledgedBtc + (uReserve / price);
  const baseColUsd = baseOkxBtc * price;

  // 50% LTV Dynamic Target
  const deltaDebt50 = Math.max((0.50 * baseColUsd - debtUsdt) / 0.50, 0);
  const dynamic50TargetBtc = Number((baseOkxBtc + (deltaDebt50 / price) + extBtc).toFixed(4));

  let phase = "PHASE_0_SPRINT";
  let phaseLabel = "階段0: 1.0 BTC 衝刺待發 (等待右側信號)";
  let recommendedDca = "每 1 小時 5 USDT (120 U/天)";
  let exitRoadmap = "$140k (賣0.10 BTC負債砍半) ➔ $160k (賣0.09 BTC負債清零，手握 1.02+ BTC 純現貨)";

  if (price >= 160000 || btcMM >= 2.25) {
    phase = "PHASE_5_EXIT";
    phaseLabel = "階段5: 頂峰清償 (負債歸0)";
    recommendedDca = "0 USDT (停投)";
  } else if (price >= 140000 || btcMM >= 1.95) {
    phase = "PHASE_5_DE_LEVERAGE";
    phaseLabel = "階段5: 狂熱減債砍半";
    recommendedDca = "0 USDT (停投)";
  } else if (price >= 120000 || btcMM >= 1.70) {
    phase = "PHASE_4_HOLD";
    phaseLabel = "階段4: 12萬狂熱停投 (一股不賣)";
    recommendedDca = "0 USDT (停投)";
  } else if (price >= 100000 || btcMM >= 1.45) {
    phase = "PHASE_3_CRUISE_2";
    phaseLabel = "階段3: 十萬大關後微速守望 ($100k~$120k)";
    recommendedDca = "每 24 小時 5 USDT (5 U/天)";
  } else if (totalBtc >= (dynamic50TargetBtc - 0.02)) {
    phase = "PHASE_3_CRUISE_1";
    phaseLabel = "階段3: 十萬大關前主升浪巡航 ($80k~$100k)";
    recommendedDca = "每 12 小時 5 USDT (10 U/天)";
  } else if (totalBtc >= 1.0 || activeCryptoLTV >= 0.38) {
    phase = "PHASE_2_CLIMB";
    phaseLabel = "階段2: 30天整數爬坡 DCA (40% ➔ 50% LTV)";
    recommendedDca = "每 1 小時 12 USDT (288 U/天)";
  }

  return {
    phase: phase,
    phaseLabel: phaseLabel,
    recommendedDca: recommendedDca,
    dynamic50TargetBtc: dynamic50TargetBtc,
    exitRoadmap: exitRoadmap
  };
}

function getBtcAllocationTargets_(btcRegime) {
  const regime = btcRegime && btcRegime.regime;
  const targets = {
    PANIC_ACCUMULATE: { L1: 0.75, L2: 0.15, L3: 0.10 },
    DEEP_VALUE: { L1: 0.72, L2: 0.18, L3: 0.10 },
    ACCUMULATE: { L1: 0.70, L2: 0.20, L3: 0.10 },
    NEUTRAL: { L1: 0.65, L2: 0.25, L3: 0.10 },
    DE_RISK: { L1: 0.60, L2: 0.30, L3: 0.10 },
    BUBBLE: { L1: 0.50, L2: 0.30, L3: 0.20 }
  };
  return targets[regime] || null;
}

function shouldSuppressAggressiveBtcBuying_(btcRegime) {
  if (!btcRegime) return false;
  return btcRegime.restockAllowed !== true;
}

function formatPercent_(value, digits) {
  const num = toFiniteNumber_(value);
  if (num === null) return "N/A";
  const precision = digits === undefined ? 0 : digits;
  return (num * 100).toFixed(precision) + "%";
}

function formatBtcLtvRange_(btcRegime) {
  if (!btcRegime) return "N/A";
  if (btcRegime.targetLtvMin === btcRegime.targetLtvMax) {
    return formatPercent_(btcRegime.targetLtvMax, 0);
  }
  return formatPercent_(btcRegime.targetLtvMin, 0) + " - " + formatPercent_(btcRegime.targetLtvMax, 0);
}

function buildCryptoLtvGuardrailAction_(btcRegime) {
  if (!btcRegime || !btcRegime.guardrailAction) return "";
  return btcRegime.guardrailAction;
}
