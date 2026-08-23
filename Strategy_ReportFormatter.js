/**
 * Strategy_ReportFormatter.js
 * Sovereign Asset Protocol - Portfolio Snapshot Generation & Notification Broadcast
 */

/**
 * Helper function to get admin email with validation
 * @private
 */
function getAdminEmail_() {
  const email = Settings.get('ADMIN_EMAIL');
  if (!email) {
    LogService.warn("ScriptProperty 'ADMIN_EMAIL' not set.", "Config:Email");
  }
  return email || "";
}

function generatePortfolioSnapshot(context) {
  const { market, pledgeGroups, netEntityValue, indicators, totalGrossAssets, portfolioSummary } = context;
  const btcRegime = market.btcRegime;

  let s = "\n[I] 市場情報 (MARKET INTEL)\n";
  s += "- BTC 現貨價格: $" + market.btcPrice.toLocaleString() + " USD\n";
  if (market.sapBaseATH > 0 && market.btcDrawdownFromATH !== null) {
    s += "- 距離 ATH (" + market.sapBaseATH + "): " + formatPercent_(market.btcDrawdownFromATH, 1) + "\n";
  } else {
    s += "- 距離 ATH: N/A（缺少 SAP_Base_ATH）\n";
  }

  if (market.btcMM) {
    s += "- Mayer Multiple: " + market.btcMM.toFixed(2) + "\n";
  }
  if (btcRegime) {
    s += "- BTC Regime: " + btcRegime.regime + "\n";
    s += "- 週期定位: " + btcRegime.phaseLabel + "（" + btcRegime.reason + "）\n";
  }

  if (market.bullStrategy) {
    s += "- 🚀 長牛戰略: " + market.bullStrategy.phaseLabel + "\n";
    s += "- 🤖 OKX DCA: " + market.bullStrategy.recommendedDca + "\n";
    s += "- 🎯 50% 滿額目標: " + market.bullStrategy.dynamic50TargetBtc + " BTC (依真實庫存動態精算)\n";
    s += "- 🏁 頂部退場: " + market.bullStrategy.exitRoadmap + "\n";
  }

  // [NEW v24.13] TW Weighted MM Display
  if (market.twWeightedMM) {
    let twPhase = "";
    if (market.twWeightedMM > 1.35) twPhase = "🔴 極度泡沫";
    else if (market.twWeightedMM > 1.15) twPhase = "🟠 高位警戒";
    else if (market.twWeightedMM > 1.00) twPhase = "🟡 中性平衡";
    else if (market.twWeightedMM > 0.85) twPhase = "🟢 低位部屬";
    else twPhase = "🟢 深水炸彈 (機會)";

    s += "- 台股加權 MM: " + market.twWeightedMM.toFixed(2) + " (" + twPhase + ")\n";
  }

  s += "\n[II] 生存指標 (SURVIVAL METRICS)\n";
  s += "- 生存跑道: " + indicators.survivalRunway.toFixed(1) + " 個月\n";
  s += "- 淨值: " + Math.round(netEntityValue).toLocaleString() + " TWD\n";
  s += "- 總資產: " + Math.round(totalGrossAssets).toLocaleString() + " TWD\n";
  s += "- 總負債: " + Math.round(totalGrossAssets - netEntityValue).toLocaleString() + " TWD\n";
  s += "- 總 LTV: " + (indicators.ltv * 100).toFixed(1) + "%\n";

  if (btcRegime) {
    s += "- 目標 LTV (Crypto): " + formatBtcLtvRange_(btcRegime) + "\n";
    s += "- 活性 LTV (Active): " + (indicators.cryptoLTV * 100).toFixed(1) + "%\n";
    s += "- 總體 LTV (Global): " + (indicators.globalCryptoLTV * 100).toFixed(1) + "%\n";
    s += "- Restock Mode: " + btcRegime.restockMode + "\n";

    if (btcRegime.guardrailMessage) {
      s += "  > " + btcRegime.guardrailMessage + "\n";
      s += "  > " + btcRegime.guardrailAction + "\n";
    } else if (btcRegime.targetLtvMax === 0 && indicators.cryptoLTV > 0) {
      s += "  > 此 regime 的目標 LTV 為 0%，優先降槓桿而非新增曝險。\n";
    } else if (btcRegime.targetLtvMax > 0 && indicators.cryptoLTV > btcRegime.targetLtvMax) {
      s += "  > Active LTV 高於 regime 目標區間上緣，新增借款前需先降槓桿。\n";
    } else if (btcRegime.restockAllowed) {
      s += "  > 質押風險位於允許區間，可依 regime 執行分批 restock。\n";
    } else if (btcRegime.restockMode === "OFF") {
      s += "  > 停止戰術 restock，優先維持流動性或降低槓桿。\n";
    } else {
      s += "  > 只維持基本 DCA 或手動檢查，不執行戰術 restock。\n";
    }
  }

  if (pledgeGroups.length > 0) {
    s += "\n[質押健康度]\n";
    pledgeGroups.forEach(group => {
      let status = "✅";
      if (group.ratio < group.critical) status = "🛑 危險";
      else if (group.ratio < group.alert) status = "⚠️ 警戒";

      let limitInfo = "";
      if (group.name.includes("Stock")) {
        limitInfo = " (安全線 > " + group.critical + ")";
      }

      const groupLTV = (1 / group.ratio * 100).toFixed(1);
      s += "- " + group.name + ": " + group.ratio.toFixed(2) + " (LTV " + groupLTV + "%)" + limitInfo + " " + status + "\n";
    });
  }

  s += buildStockExposureSnapshot_(context.stockStrategy);

  s += "\n[III] 資產配置 (ASSET ALLOCATION)\n";
  const groupsToDisplay = context.assetGroups || Config.ASSET_GROUPS;
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
      if (targetPct > 0) {
        line += " / 容許 " + targetPct.toFixed(1) + "%)\n";
      } else {
        line += ")\n";
      }
      if (groupValue > 0 && pct > targetPct) {
        line += "  > ⚠️ 待清理: " + group.tickers.join(", ") + "\n";
      }
    }
    s += line;
  });

  s += "\n[IV] 再平衡建議 (REBALANCE)\n";
  const rebalanceTargets = context.rebalanceTargets || [];
  if (rebalanceTargets.length === 0) {
    s += "- 目前配置落在可接受誤差內\n";
  } else {
    const sequenceSummary = buildRebalanceSequenceSummary_(rebalanceTargets);
    if (sequenceSummary.length > 0) {
      s += "- 建議執行順序: " + sequenceSummary.join(" -> ") + "\n";
    }
    rebalanceTargets.slice(0, 5).forEach(target => {
      const shortName = getRebalanceShortName_(target);
      const currentPct = (target.currentWeight * 100).toFixed(1);
      const targetPct = (target.targetWeight * 100).toFixed(1);
      const deltaAbs = Math.round(Math.abs(target.deltaValue)).toLocaleString();

      if (target.action === 'CLEAR') {
        s += "- " + shortName + ": 目前 " + currentPct + "%\n";
        s += "  > 建議清理 " + Math.round(Math.abs(target.deltaValue)).toLocaleString() + " TWD";
        if (target.executionHint) s += " | " + target.executionHint;
        s += "\n";
        const productHint = formatRebalanceProductHint_(target);
        if (productHint) {
          s += "  > 優先標的: " + productHint + "\n";
        }
        return;
      }

      const actionLabel = target.action === 'ADD' ? "補強" : "減碼";
      s += "- " + shortName + ": 目前 " + currentPct + "% / 目標 " + targetPct + "%\n";
      s += "  > 建議" + actionLabel + " " + deltaAbs + " TWD";
      if (target.suggestedFundingSource) s += " | 資金方向: " + target.suggestedFundingSource;
      s += "\n";
      const productHint = formatRebalanceProductHint_(target);
      if (productHint) {
        s += "  > 優先標的: " + productHint + "\n";
      }
    });
  }

  s += buildDataFreshnessReport_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
  s += "----------------------------------------\n";
  s += "最後更新: " + new Date().toLocaleString('zh-TW', { hour12: false });
  return s;
}


function broadcastReport_(context, alerts = []) {
  const hasAlerts = alerts.length > 0;
  const snapshot = generatePortfolioSnapshot(context);

  // 1. Email Channel
  const emailRecipient = getAdminEmail_();
  if (emailRecipient) {
    try {
      let subject = hasAlerts ? "[SAP 戰略顧問] 需要採取行動" : "[SAP 每日狀態] 一切正常";
      let body = hasAlerts ? "戰略夥伴，\n分析顯示需要進行再平衡：\n\n" : "戰略夥伴，\n目前系統運作正常。\n\n";

      if (hasAlerts) {
        alerts.forEach(a => { body += "**" + a.level + "**\n" + a.message + "\n指令: " + a.action + "\n\n"; });
      }
      body += snapshot;

      MailApp.sendEmail(emailRecipient, subject, body);
      console.log(`[Broadcast] Email sent to ${emailRecipient}`);
    } catch (e) {
      console.error(`[Broadcast] Email failed: ${e.toString()}`);
    }
  }

  // 2. Discord Channel (Sync)
  if (typeof sendDiscordAlert_ === 'function') {
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

    const discordSent = sendDiscordAlert_(title, description, color);
    console.log(`[Broadcast] Discord sent: ${discordSent}`);
  } else {
    console.warn("[Broadcast] Discord sender function is not loaded.");
  }
}
