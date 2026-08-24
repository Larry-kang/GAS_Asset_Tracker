/**
 * Util_HtmlReport.js
 * Sovereign Asset Protocol - Modern HTML Report & Dashboard Renderer
 *
 * Supports:
 * 1. Email-safe HTML snapshot broadcasting.
 * 2. Google Sheets interactive Sidebar view.
 * 3. Standalone Web App (doGet) monitoring page.
 */

/**
 * Helper to escape untrusted text for safe HTML output
 * @private
 */
function escapeHtml_(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Generates modern dark-themed HTML report.
 * @param {Object} context - Strategy context built from buildFreshContext() or buildContext()
 * @param {Array} alerts - Optional array of active alert objects [{ level, message, action }]
 * @param {Object} options - Optional flags { isEmail: boolean, isInteractive: boolean }
 * @returns {string} HTML string
 */
function generatePortfolioSnapshotHtml(context, alerts, options) {
  const opts = options || {};
  const isEmail = !!opts.isEmail;
  const isInteractive = !isEmail && !!opts.isInteractive;
  const ctx = context || {};

  const market = ctx.market || {};
  const indicators = ctx.indicators || {};
  const pledgeGroups = ctx.pledgeGroups || [];
  const assetGroups = ctx.assetGroups || [];
  const rebalanceTargets = ctx.rebalanceTargets || [];
  const totalGrossAssets = ctx.totalGrossAssets || 0;
  const netEntityValue = ctx.netEntityValue || 0;
  const totalDebt = Math.max(0, totalGrossAssets - netEntityValue);
  const activeAlerts = Array.isArray(alerts) ? alerts : [];

  const btcPrice = (market.btcPrice !== null && market.btcPrice !== undefined && !isNaN(market.btcPrice))
    ? "$" + Number(market.btcPrice).toLocaleString()
    : "N/A";
  const btcATH = market.sapBaseATH > 0 && market.btcDrawdownFromATH !== null && market.btcDrawdownFromATH !== undefined
    ? (Number(market.btcDrawdownFromATH) * 100).toFixed(1) + "%"
    : "N/A";
  const btcMM = market.btcMM !== null && market.btcMM !== undefined ? Number(market.btcMM).toFixed(2) : "N/A";
  const btcRegime = market.btcRegime;

  const runway = indicators.survivalRunway !== undefined ? Number(indicators.survivalRunway).toFixed(1) + " 個月" : "N/A";
  const activeLtv = indicators.cryptoLTV !== undefined ? (Number(indicators.cryptoLTV) * 100).toFixed(1) + "%" : "N/A";
  const globalLtv = indicators.ltv !== undefined ? (Number(indicators.ltv) * 100).toFixed(1) + "%" : "N/A";

  const systemName = (typeof Config !== 'undefined' && Config.SYSTEM_NAME) ? Config.SYSTEM_NAME.split(' - ')[0] : 'SAP';
  const systemVersion = (typeof Config !== 'undefined' && Config.VERSION) ? Config.VERSION : 'v24';
  const updateTime = new Date().toLocaleString('zh-TW', { hour12: false });

  // CSS Styles
  const css = `
    :root {
      --bg-primary: #0f172a;
      --bg-secondary: #1e293b;
      --border-color: #334155;
      --text-primary: #f8fafc;
      --text-secondary: #94a3b8;
      --text-muted: #64748b;
      --accent-blue: #38bdf8;
      --accent-green: #34d399;
      --accent-yellow: #fbbf24;
      --accent-orange: #fb923c;
      --accent-red: #f87171;
      --accent-purple: #c084fc;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background-color: #0f172a; color: #f8fafc; padding: 14px; line-height: 1.5; font-size: 14px; }
    .container { width: 100%; max-width: 680px; margin: 0 auto; }
    .header { background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border: 1px solid #334155; border-radius: 12px; padding: 16px; margin-bottom: 14px; border-top: 3px solid #38bdf8; }
    .header-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
    .header-title { font-size: 1.15rem; font-weight: 700; color: #f8fafc; display: flex; align-items: center; gap: 6px; }
    .badge { display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 9999px; font-size: 0.72rem; font-weight: 600; text-transform: uppercase; }
    .badge-success { background: rgba(52, 211, 153, 0.15); color: #34d399; border: 1px solid rgba(52, 211, 153, 0.3); }
    .badge-warning { background: rgba(251, 191, 36, 0.15); color: #fbbf24; border: 1px solid rgba(251, 191, 36, 0.3); }
    .badge-danger { background: rgba(248, 113, 113, 0.15); color: #f87171; border: 1px solid rgba(248, 113, 113, 0.3); }
    .badge-info { background: rgba(56, 189, 248, 0.15); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.3); }
    .header-sub { color: #94a3b8; font-size: 0.78rem; }
    .card { background: #1e293b; border: 1px solid #334155; border-radius: 12px; padding: 16px; margin-bottom: 14px; }
    .card-title { font-size: 0.9rem; font-weight: 700; color: #f8fafc; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 1px solid rgba(255, 255, 255, 0.08); text-transform: uppercase; letter-spacing: 0.04em; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
    @media (max-width: 480px) { .grid-3, .grid-2 { grid-template-columns: 1fr; } }
    .stat-box { background: rgba(15, 23, 42, 0.65); border: 1px solid rgba(255, 255, 255, 0.05); border-radius: 8px; padding: 10px; }
    .stat-label { font-size: 0.72rem; color: #94a3b8; margin-bottom: 2px; }
    .stat-value { font-size: 1.1rem; font-weight: 700; color: #f8fafc; }
    .stat-desc { font-size: 0.7rem; color: #64748b; margin-top: 2px; }
    .alloc-item { margin-bottom: 10px; }
    .alloc-header { display: flex; justify-content: space-between; font-size: 0.8rem; margin-bottom: 4px; }
    .progress-bar-bg { height: 7px; background: rgba(255, 255, 255, 0.08); border-radius: 9999px; overflow: hidden; }
    .progress-fill { height: 100%; border-radius: 9999px; }
    .action-item { background: rgba(15, 23, 42, 0.6); border-left: 4px solid #38bdf8; border-radius: 0 8px 8px 0; padding: 10px; margin-bottom: 8px; font-size: 0.82rem; }
    .action-item.add { border-left-color: #34d399; }
    .action-item.clear { border-left-color: #fb923c; }
    .action-item.danger { border-left-color: #f87171; }
    .action-header { display: flex; justify-content: space-between; font-weight: 600; margin-bottom: 2px; }
    .action-body { color: #94a3b8; font-size: 0.75rem; }
    .alert-card { background: rgba(248, 113, 113, 0.1); border: 1px solid rgba(248, 113, 113, 0.3); border-radius: 8px; padding: 12px; margin-bottom: 12px; font-size: 0.82rem; }
    .alert-card.warning { background: rgba(251, 191, 36, 0.1); border-color: rgba(251, 191, 36, 0.3); }
    .button-group { display: flex; gap: 8px; margin-top: 14px; margin-bottom: 10px; }
    .btn { flex: 1; padding: 9px 14px; border-radius: 8px; font-size: 0.82rem; font-weight: 600; cursor: pointer; border: none; text-align: center; }
    .btn-primary { background: #38bdf8; color: #0f172a; }
    .btn-secondary { background: #1e293b; color: #f8fafc; border: 1px solid #334155; }
    .btn:disabled { opacity: 0.6; cursor: not-allowed; }
    .footer { text-align: center; font-size: 0.72rem; color: #64748b; padding: 10px 0; }
    #status-bar { margin-top: 8px; font-size: 0.75rem; color: #38bdf8; text-align: center; display: none; }
  `;

  // Section 0: Active Alerts (if any)
  let alertsHtml = "";
  if (activeAlerts.length > 0) {
    alertsHtml = '<div style="margin-bottom: 14px;">';
    activeAlerts.forEach(a => {
      const isWarn = a.level && (a.level.includes('WARN') || a.level.includes('警戒'));
      const cardClass = isWarn ? 'alert-card warning' : 'alert-card';
      const badgeClass = isWarn ? 'badge badge-warning' : 'badge badge-danger';
      alertsHtml += `
        <div class="${cardClass}">
          <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
            <strong style="color: #f8fafc;">${a.message || '系統警報'}</strong>
            <span class="${badgeClass}">${a.level || 'ALERT'}</span>
          </div>
          <div style="color: #94a3b8;">行動指示: ${a.action || '請檢視資產負債表'}</div>
        </div>
      `;
    });
    alertsHtml += '</div>';
  }

  // Section 1: Market Intel
  let bullStrategyNote = "";
  if (market.bullStrategy) {
    bullStrategyNote = `
      <div style="margin-top: 10px; font-size: 0.78rem; color: #94a3b8; background: rgba(15,23,42,0.45); padding: 8px 10px; border-radius: 6px;">
        🚀 <strong>長牛戰略 (${market.bullStrategy.phaseLabel || '運作中'})</strong>：
        OKX DCA: ${market.bullStrategy.recommendedDca || 'N/A'} |
        🎯 50% 滿額目標: ${market.bullStrategy.dynamic50TargetBtc || 'N/A'} BTC |
        🏁 頂部退場: ${market.bullStrategy.exitRoadmap || 'N/A'}
      </div>
    `;
  }

  let twMmBox = "";
  if (market.twWeightedMM) {
    let twColor = "#34d399";
    let twText = "中性平衡";
    if (market.twWeightedMM > 1.35) { twColor = "#f87171"; twText = "極度泡沫"; }
    else if (market.twWeightedMM > 1.15) { twColor = "#fb923c"; twText = "高位警戒"; }
    else if (market.twWeightedMM < 0.85) { twColor = "#38bdf8"; twText = "低位機會"; }

    twMmBox = `
      <div class="stat-box">
        <div class="stat-label">台股加權 MM</div>
        <div class="stat-value" style="color: ${twColor};">${Number(market.twWeightedMM).toFixed(2)}</div>
        <div class="stat-desc">${twText}</div>
      </div>
    `;
  }

  // Section 2: Survival Metrics & LTV
  let pledgeHtml = "";
  if (pledgeGroups.length > 0) {
    pledgeHtml = '<div style="margin-top: 10px; font-size: 0.78rem; color: #94a3b8;"><strong>質押健康度：</strong>';
    pledgeGroups.forEach(g => {
      let badge = '<span class="badge badge-success">✅ 安全</span>';
      if (g.ratio < g.critical) badge = '<span class="badge badge-danger">🛑 危險</span>';
      else if (g.ratio < g.alert) badge = '<span class="badge badge-warning">⚠️ 警戒</span>';
      const ltvPct = g.ratio > 0 ? (1 / g.ratio * 100).toFixed(1) + "%" : "N/A";
      pledgeHtml += ` ${g.name} 維持率 ${g.ratio ? g.ratio.toFixed(2) : 'N/A'} (LTV ${ltvPct}) ${badge} &nbsp; `;
    });
    pledgeHtml += '</div>';
  }

  // Section 3: Asset Allocation Bars
  const colors = [
    'linear-gradient(90deg, #f59e0b, #fbbf24)',
    'linear-gradient(90deg, #3b82f6, #60a5fa)',
    'linear-gradient(90deg, #10b981, #34d399)',
    '#f97316',
    '#a855f7'
  ];
  let allocHtml = "";
  assetGroups.forEach((group, idx) => {
    const val = group.value || 0;
    const pct = totalGrossAssets > 0 ? (val / totalGrossAssets * 100) : 0;
    const targetPct = (group.target !== undefined ? group.target : (group.defaultTarget || 0)) * 100;
    const color = colors[idx % colors.length];

    allocHtml += `
      <div class="alloc-item">
        <div class="alloc-header">
          <span>${group.name ? group.name.split(':')[0] : 'Layer'}</span>
          <span><strong>NT$ ${Math.round(val).toLocaleString()}</strong> (${pct.toFixed(1)}% / 目標 ${targetPct.toFixed(0)}%)</span>
        </div>
        <div class="progress-bar-bg">
          <div class="progress-fill" style="width: ${Math.min(100, Math.max(0, pct))}%; background: ${color};"></div>
        </div>
      </div>
    `;
  });

  // Section 4: Rebalance Actions
  let rebalanceHtml = "";
  if (rebalanceTargets.length === 0) {
    rebalanceHtml = '<div style="font-size: 0.8rem; color: #94a3b8;">目前配置落在可接受誤差區間內。</div>';
  } else {
    rebalanceTargets.slice(0, 5).forEach(t => {
      const isClear = t.action === 'CLEAR';
      const isAdd = t.action === 'ADD';
      const itemClass = isClear ? 'action-item clear' : (isAdd ? 'action-item add' : 'action-item');
      const badgeClass = isClear ? 'badge badge-warning' : (isAdd ? 'badge badge-success' : 'badge badge-info');
      const actionText = isClear ? '建議清理' : (isAdd ? '建議補強' : '建議調整');
      const deltaVal = t.deltaValue !== undefined ? Math.round(Math.abs(t.deltaValue)).toLocaleString() : '0';
      const shortName = t.shortName || (t.group && t.group.name) || '資產';

      rebalanceHtml += `
        <div class="${itemClass}">
          <div class="action-header">
            <span>${shortName}</span>
            <span class="${badgeClass}">${actionText}</span>
          </div>
          <div class="action-body">
            金額: <strong>NT$ ${deltaVal}</strong> TWD
            ${t.suggestedFundingSource ? ' | 資金方向: ' + t.suggestedFundingSource : ''}
            ${t.executionHint ? ' | 提示: ' + t.executionHint : ''}
          </div>
        </div>
      `;
    });
  }

  // Interactive buttons & client script for Sidebar / WebApp
  let interactiveControls = "";
  let clientScript = "";
  if (isInteractive) {
    interactiveControls = `
      <div class="button-group">
        <button id="btn-sync" class="btn btn-primary" onclick="triggerSyncMaster()">🔄 一鍵全系統同步</button>
        <button id="btn-broadcast" class="btn btn-secondary" onclick="triggerBroadcast()">📨 立即廣播報告</button>
      </div>
      <div id="status-bar"></div>
    `;

    clientScript = `
      <script>
        function setStatus(msg, isError) {
          var bar = document.getElementById('status-bar');
          if (!bar) return;
          bar.style.display = 'block';
          bar.style.color = isError ? '#f87171' : '#38bdf8';
          bar.innerHTML = msg;
        }

        function setButtonsDisabled(disabled) {
          var b1 = document.getElementById('btn-sync');
          var b2 = document.getElementById('btn-broadcast');
          if (b1) b1.disabled = disabled;
          if (b2) b2.disabled = disabled;
        }

        function triggerSyncMaster() {
          if (typeof google === 'undefined' || !google.script || !google.script.run) {
            alert('本功能需在 Google Apps Script 環境中運行');
            return;
          }
          setButtonsDisabled(true);
          setStatus('⏳ 正在執行全系統匯率、價格與交易所資產同步...', false);
          google.script.run
            .withSuccessHandler(function(res) {
              setButtonsDisabled(false);
              setStatus('✅ 全系統同步完成！請重新整理以查看最新數據。', false);
            })
            .withFailureHandler(function(err) {
              setButtonsDisabled(false);
              setStatus('❌ 同步失敗: ' + (err.message || err), true);
            })
            .runAutomationMaster();
        }

        function triggerBroadcast() {
          if (typeof google === 'undefined' || !google.script || !google.script.run) {
            alert('本功能需在 Google Apps Script 環境中運行');
            return;
          }
          if (!confirm('確認立即發送戰略報告至 Discord 與 Email？')) return;
          setButtonsDisabled(true);
          setStatus('⏳ 正在發送戰略報告...', false);
          google.script.run
            .withSuccessHandler(function(res) {
              setButtonsDisabled(false);
              setStatus('✅ 報告已成功廣播至 Discord 與 Email！', false);
            })
            .withFailureHandler(function(err) {
              setButtonsDisabled(false);
              setStatus('❌ 廣播失敗: ' + (err.message || err), true);
            })
            .triggerManualReport();
        }
      </script>
    `;
  }

  // Full HTML Construction
  return `<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${systemName} 戰略指揮中心報告</title>
  <style>${css}</style>
</head>
<body>
<div class="container">
  <!-- Header -->
  <div class="header">
    <div class="header-top">
      <div class="header-title">
        <span>⚡ ${systemName} 戰略指揮中心</span>
      </div>
      <span class="badge ${activeAlerts.length > 0 ? 'badge-warning' : 'badge-success'}">
        ● ${activeAlerts.length > 0 ? '需要關注' : '狀態穩固'}
      </span>
    </div>
    <div class="header-sub">
      模式: ${ctx.phase || 'Bitcoin Standard ' + systemVersion} | 時間: ${updateTime}
    </div>
  </div>

  ${alertsHtml}

  <!-- Market Intel -->
  <div class="card">
    <div class="card-title">🌐 [I] 市場情報 (Market Intel)</div>
    <div class="grid-3">
      <div class="stat-box">
        <div class="stat-label">BTC 現價</div>
        <div class="stat-value" style="color: var(--accent-yellow);">${btcPrice}</div>
        <div class="stat-desc">距 ATH: ${btcATH}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Mayer Multiple (MM)</div>
        <div class="stat-value" style="color: var(--accent-green);">${btcMM}</div>
        <div class="stat-desc">200D SMA 對比</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">BTC 週期定位</div>
        <div class="stat-value" style="color: var(--accent-blue);">${btcRegime ? btcRegime.regime : 'NORMAL'}</div>
        <div class="stat-desc">${btcRegime ? btcRegime.phaseLabel : '穩定監控中'}</div>
      </div>
    </div>
    ${twMmBox ? `<div style="margin-top: 8px;">${twMmBox}</div>` : ''}
    ${bullStrategyNote}
  </div>

  <!-- Survival Metrics -->
  <div class="card">
    <div class="card-title">🛡️ [II] 生存指標與風控 (Survival & Risk)</div>
    <div class="grid-2" style="margin-bottom: 10px;">
      <div class="stat-box">
        <div class="stat-label">淨實體價值 (Net Entity Value)</div>
        <div class="stat-value" style="color: var(--accent-green); font-size: 1.25rem;">NT$ ${Math.round(netEntityValue).toLocaleString()}</div>
        <div class="stat-desc">總資產 NT$ ${Math.round(totalGrossAssets).toLocaleString()} / 負債 NT$ ${Math.round(totalDebt).toLocaleString()}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">生存跑道 (Survival Runway)</div>
        <div class="stat-value" style="color: var(--accent-blue); font-size: 1.25rem;">${runway}</div>
        <div class="stat-desc">總體 LTV: ${globalLtv}</div>
      </div>
    </div>

    <div class="grid-2">
      <div class="stat-box">
        <div class="stat-label">OKX 活性 LTV (Active Crypto)</div>
        <div class="stat-value" style="color: var(--accent-green);">${activeLtv}</div>
        <div class="stat-desc">目標區間: 40%–45%</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">總體槓桿率 (Global LTV)</div>
        <div class="stat-value">${globalLtv}</div>
        <div class="stat-desc">宏觀風控警示線</div>
      </div>
    </div>
    ${pledgeHtml}
  </div>

  <!-- Asset Allocation -->
  <div class="card">
    <div class="card-title">📊 [III] 資產配置 (Asset Allocation)</div>
    ${allocHtml}
  </div>

  <!-- Rebalance Advice -->
  <div class="card">
    <div class="card-title">⚖️ [IV] 再平衡建議 (Rebalance Advice)</div>
    ${rebalanceHtml}
  </div>

  ${interactiveControls}

  <div class="footer">
    Sovereign Asset Protocol (SAP) · 保持戰略 · 保持理性
  </div>
</div>
${clientScript}
</body>
</html>`;
}
