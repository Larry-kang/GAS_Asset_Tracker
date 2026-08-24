/**
 * Util_HtmlReport.js
 * Sovereign Asset Protocol - Modern Clean Light Mode HTML Report & Dashboard Renderer
 *
 * Supports:
 * 1. Email-safe Light Mode HTML snapshot broadcasting.
 * 2. Google Sheets interactive Light Mode Sidebar view.
 * 3. Standalone Web App (doGet) monitoring page with password lock.
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
 * Generates modern clean Light-themed HTML report.
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

  // Clean Fintech Light Theme CSS
  const css = `
    :root {
      --bg-body: #f1f5f9;
      --bg-container: #ffffff;
      --bg-card: #ffffff;
      --bg-stat: #f8fafc;
      --border-color: #e2e8f0;
      --border-subtle: #cbd5e1;
      --text-primary: #0f172a;
      --text-secondary: #475569;
      --text-muted: #64748b;
      --accent-blue: #0284c7;
      --accent-green: #059669;
      --accent-yellow: #d97706;
      --accent-orange: #ea580c;
      --accent-red: #dc2626;
      --accent-purple: #7c3aed;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background-color: #f1f5f9; color: #0f172a; padding: 14px; line-height: 1.5; font-size: 14px; }
    .container { width: 100%; max-width: 650px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.04); }
    .header { background: linear-gradient(135deg, #f8fafc 0%, #edf2f7 100%); border: 1px solid #e2e8f0; border-left: 4px solid #0284c7; border-radius: 8px; padding: 14px 16px; margin-bottom: 16px; }
    .header-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
    .header-title { font-size: 1.15rem; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 6px; }
    .badge { display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 9999px; font-size: 0.72rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em; }
    .badge-success { background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }
    .badge-warning { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
    .badge-danger { background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca; }
    .badge-info { background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; }
    .header-sub { color: #64748b; font-size: 0.78rem; }
    .card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px 16px; margin-bottom: 14px; }
    .card-title { font-size: 0.88rem; font-weight: 700; color: #1e293b; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 2px solid #f1f5f9; text-transform: uppercase; letter-spacing: 0.04em; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
    @media (max-width: 480px) { .grid-3, .grid-2 { grid-template-columns: 1fr; } }
    .stat-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; }
    .stat-label { font-size: 0.7rem; font-weight: 600; color: #64748b; margin-bottom: 2px; text-transform: uppercase; }
    .stat-value { font-size: 1.1rem; font-weight: 700; color: #0f172a; }
    .stat-desc { font-size: 0.7rem; color: #64748b; margin-top: 2px; }
    .alloc-item { margin-bottom: 10px; }
    .alloc-header { display: flex; justify-content: space-between; font-size: 0.8rem; margin-bottom: 4px; color: #334155; }
    .progress-bar-bg { height: 8px; background: #e2e8f0; border-radius: 9999px; overflow: hidden; }
    .progress-fill { height: 100%; border-radius: 9999px; }
    .action-item { background: #f8fafc; border: 1px solid #e2e8f0; border-left: 4px solid #0284c7; border-radius: 0 8px 8px 0; padding: 10px 12px; margin-bottom: 8px; font-size: 0.82rem; }
    .action-item.add { border-left-color: #059669; }
    .action-item.clear { border-left-color: #ea580c; }
    .action-item.danger { border-left-color: #dc2626; }
    .action-header { display: flex; justify-content: space-between; font-weight: 700; color: #0f172a; margin-bottom: 2px; }
    .action-body { color: #475569; font-size: 0.76rem; line-height: 1.45; }
    .alert-card { background: #fffbeb; border: 1px solid #fef3c7; border-left: 4px solid #d97706; border-radius: 0 8px 8px 0; padding: 10px 12px; margin-bottom: 12px; font-size: 0.82rem; }
    .alert-card.danger { background: #fef2f2; border-color: #fee2e2; border-left-color: #dc2626; }
    .button-group { display: flex; gap: 8px; margin-top: 14px; margin-bottom: 8px; }
    .btn { flex: 1; padding: 9px 14px; border-radius: 8px; font-size: 0.82rem; font-weight: 600; cursor: pointer; border: none; text-align: center; transition: all 0.2s; }
    .btn-primary { background: #0284c7; color: #ffffff; }
    .btn-primary:hover { background: #0369a1; }
    .btn-secondary { background: #ffffff; color: #1e293b; border: 1px solid #cbd5e1; }
    .btn-secondary:hover { background: #f8fafc; }
    .btn:disabled { opacity: 0.6; cursor: not-allowed; }
    .footer { text-align: center; font-size: 0.72rem; color: #94a3b8; padding-top: 12px; border-top: 1px solid #f1f5f9; margin-top: 14px; }
    #status-bar { margin-top: 8px; font-size: 0.75rem; color: #0284c7; text-align: center; display: none; }
  `;

  // Section 0: Active Alerts (if any)
  let alertsHtml = "";
  if (activeAlerts.length > 0) {
    alertsHtml = '<div style="margin-bottom: 14px;">';
    activeAlerts.forEach(a => {
      const isDanger = a.level && (a.level.includes('CRITICAL') || a.level.includes('緊急') || a.level.includes('DEFCON'));
      const cardClass = isDanger ? 'alert-card danger' : 'alert-card';
      const badgeClass = isDanger ? 'badge badge-danger' : 'badge badge-warning';
      alertsHtml += `
        <div class="${cardClass}">
          <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
            <strong style="color: #0f172a;">${escapeHtml_(a.message || '系統警報')}</strong>
            <span class="${badgeClass}">${escapeHtml_(a.level || 'ALERT')}</span>
          </div>
          <div style="color: #475569;">行動指示: ${escapeHtml_(a.action || '請檢視資產負債表')}</div>
        </div>
      `;
    });
    alertsHtml += '</div>';
  }

  // Section 1: Market Intel
  let bullStrategyNote = "";
  if (market.bullStrategy) {
    bullStrategyNote = `
      <div style="margin-top: 10px; font-size: 0.78rem; color: #475569; background: #f1f5f9; padding: 8px 10px; border-radius: 6px; line-height: 1.45;">
        🚀 <strong>長牛戰略 (${escapeHtml_(market.bullStrategy.phaseLabel || '運作中')})</strong>：
        OKX DCA: ${escapeHtml_(market.bullStrategy.recommendedDca || 'N/A')} |
        🎯 50% 滿額目標: ${escapeHtml_(market.bullStrategy.dynamic50TargetBtc || 'N/A')} BTC |
        🏁 頂部退場: ${escapeHtml_(market.bullStrategy.exitRoadmap || 'N/A')}
      </div>
    `;
  }

  let twMmBox = "";
  if (market.twWeightedMM) {
    let twColor = "#059669";
    let twText = "中性平衡";
    if (market.twWeightedMM > 1.35) { twColor = "#dc2626"; twText = "極度泡沫"; }
    else if (market.twWeightedMM > 1.15) { twColor = "#ea580c"; twText = "高位警戒"; }
    else if (market.twWeightedMM < 0.85) { twColor = "#0284c7"; twText = "低位機會"; }

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
    pledgeHtml = '<div style="margin-top: 10px; font-size: 0.78rem; color: #475569;"><strong>質押健康度：</strong>';
    pledgeGroups.forEach(g => {
      let badge = '<span class="badge badge-success">✅ 安全</span>';
      if (g.ratio < g.critical) badge = '<span class="badge badge-danger">🛑 危險</span>';
      else if (g.ratio < g.alert) badge = '<span class="badge badge-warning">⚠️ 警戒</span>';
      const ltvPct = g.ratio > 0 ? (1 / g.ratio * 100).toFixed(1) + "%" : "N/A";
      pledgeHtml += ` ${escapeHtml_(g.name)} 維持率 ${g.ratio ? g.ratio.toFixed(2) : 'N/A'} (LTV ${ltvPct}) ${badge} &nbsp; `;
    });
    pledgeHtml += '</div>';
  }

  // Section 3: Asset Allocation Bars
  const colors = [
    'linear-gradient(90deg, #d97706, #f59e0b)',
    'linear-gradient(90deg, #2563eb, #38bdf8)',
    'linear-gradient(90deg, #059669, #34d399)',
    '#ea580c',
    '#7c3aed'
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
          <span><strong>${escapeHtml_(group.name ? group.name.split(':')[0] : 'Layer')}</strong></span>
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
    rebalanceHtml = '<div style="font-size: 0.8rem; color: #64748b;">目前配置落在可接受誤差區間內。</div>';
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
            <span>${escapeHtml_(shortName)}</span>
            <span class="${badgeClass}">${actionText}</span>
          </div>
          <div class="action-body">
            金額: <strong>NT$ ${deltaVal}</strong> TWD
            ${t.suggestedFundingSource ? ' | 資金方向: ' + escapeHtml_(t.suggestedFundingSource) : ''}
            ${t.executionHint ? ' | 提示: ' + escapeHtml_(t.executionHint) : ''}
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
        (function() {
          try {
            var params = new URLSearchParams(window.location.search);
            var urlKey = params.get('key');
            if (urlKey) {
              localStorage.setItem('sap_dashboard_key', urlKey);
            }
          } catch(e) {}
        })();

        function setStatus(msg, isError) {
          var bar = document.getElementById('status-bar');
          if (!bar) return;
          bar.style.display = 'block';
          bar.style.color = isError ? '#dc2626' : '#0284c7';
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
  <title>${escapeHtml_(systemName)} 戰略指揮中心報告</title>
  <style>${css}</style>
</head>
<body>
<div class="container">
  <!-- Header -->
  <div class="header">
    <div class="header-top">
      <div class="header-title">
        <span>⚡ ${escapeHtml_(systemName)} 戰略指揮中心</span>
      </div>
      <span class="badge ${activeAlerts.length > 0 ? 'badge-warning' : 'badge-success'}">
        ● ${activeAlerts.length > 0 ? '需要關注' : '狀態穩固'}
      </span>
    </div>
    <div class="header-sub">
      模式: <strong>${escapeHtml_(ctx.phase || 'Bitcoin Standard ' + systemVersion)}</strong> | 時間: ${escapeHtml_(updateTime)}
    </div>
  </div>

  ${alertsHtml}

  <!-- Market Intel -->
  <div class="card">
    <div class="card-title">🌐 [I] 市場情報 (Market Intel)</div>
    <div class="grid-3">
      <div class="stat-box">
        <div class="stat-label">BTC 現價</div>
        <div class="stat-value" style="color: #b45309;">${btcPrice}</div>
        <div class="stat-desc">距 ATH: ${btcATH}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Mayer Multiple (MM)</div>
        <div class="stat-value" style="color: #059669;">${btcMM}</div>
        <div class="stat-desc">200D SMA 對比</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">BTC 週期定位</div>
        <div class="stat-value" style="color: #0284c7;">${escapeHtml_(btcRegime ? btcRegime.regime : 'NORMAL')}</div>
        <div class="stat-desc">${escapeHtml_(btcRegime ? btcRegime.phaseLabel : '穩定監控中')}</div>
      </div>
    </div>
    ${twMmBox ? `<div style="margin-top: 8px;">${twMmBox}</div>` : ''}
    ${bullStrategyNote}
  </div>

  <!-- Survival Metrics -->
  <div class="card">
    <div class="card-title">🛡️ [II] 生存指標與風控 (Survival & Risk)</div>
    <div class="grid-2" style="margin-bottom: 10px;">
      <div class="stat-box" style="background: #f0fdf4; border-color: #dcfce7;">
        <div class="stat-label" style="color: #15803d;">淨實體價值 (Net Entity Value)</div>
        <div class="stat-value" style="color: #15803d; font-size: 1.25rem;">NT$ ${Math.round(netEntityValue).toLocaleString()}</div>
        <div class="stat-desc" style="color: #166534;">總資產 NT$ ${Math.round(totalGrossAssets).toLocaleString()} / 負債 NT$ ${Math.round(totalDebt).toLocaleString()}</div>
      </div>
      <div class="stat-box" style="background: #f0f9ff; border-color: #e0f2fe;">
        <div class="stat-label" style="color: #0369a1;">生存跑道 (Survival Runway)</div>
        <div class="stat-value" style="color: #0369a1; font-size: 1.25rem;">${escapeHtml_(runway)}</div>
        <div class="stat-desc" style="color: #075985;">總體 LTV: ${escapeHtml_(globalLtv)}</div>
      </div>
    </div>

    <div class="grid-2">
      <div class="stat-box">
        <div class="stat-label">OKX 活性 LTV (Active Crypto)</div>
        <div class="stat-value" style="color: #059669;">${escapeHtml_(activeLtv)}</div>
        <div class="stat-desc">目標區間: 40%–45%</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">總體槓桿率 (Global LTV)</div>
        <div class="stat-value">${escapeHtml_(globalLtv)}</div>
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

/**
 * Generates clean Light-themed Lock Screen for Web App authentication.
 * @param {string} errorMsg - Optional error message
 * @returns {string} HTML string
 */
function generateLockScreenHtml(errorMsg) {
  const errMsg = errorMsg ? `<div style="background:#fee2e2;color:#b91c1c;border:1px solid #fecaca;padding:10px 14px;border-radius:8px;font-size:0.82rem;margin-bottom:14px;">${escapeHtml_(errorMsg)}</div>` : '';
  const systemName = (typeof Config !== 'undefined' && Config.SYSTEM_NAME) ? Config.SYSTEM_NAME.split(' - ')[0] : 'SAP';

  return `<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml_(systemName)} 戰略系統 密鑰驗證</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background: #f1f5f9; color: #0f172a; display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 20px; font-size: 14px; }
    .card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 28px; width: 100%; max-width: 400px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); text-align: center; }
    .title { font-size: 1.25rem; font-weight: 700; color: #0f172a; margin-bottom: 6px; }
    .sub { color: #64748b; font-size: 0.82rem; margin-bottom: 20px; }
    .input-group { margin-bottom: 16px; text-align: left; }
    .label { display: block; font-size: 0.75rem; font-weight: 600; color: #475569; margin-bottom: 6px; text-transform: uppercase; }
    .input { width: 100%; padding: 10px 14px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 0.9rem; outline: none; transition: border-color 0.2s; }
    .input:focus { border-color: #0284c7; box-shadow: 0 0 0 3px rgba(2,132,199,0.15); }
    .btn { width: 100%; padding: 10px 16px; background: #0284c7; color: #ffffff; border: none; border-radius: 8px; font-size: 0.9rem; font-weight: 600; cursor: pointer; transition: background 0.2s; }
    .btn:hover { background: #0369a1; }
    .footer { font-size: 0.72rem; color: #94a3b8; margin-top: 20px; }
  </style>
</head>
<body>
<div class="card">
  <div style="font-size: 2.2rem; margin-bottom: 10px;">🔒</div>
  <div class="title">SAP 戰略系統 密鑰驗證</div>
  <div class="sub">請輸入授權密鑰以解鎖即時資產儀表板</div>

  ${errMsg}

  <form id="login-form" onsubmit="handleUnlock(event)">
    <div class="input-group">
      <label class="label" for="access-key">存取密鑰 (Access Key)</label>
      <input type="password" id="access-key" class="input" placeholder="請輸入密鑰" required autofocus>
    </div>
    <button type="submit" class="btn">解鎖儀表板</button>
  </form>

  <div class="footer">
    Sovereign Asset Protocol (SAP) · 安全防護中
  </div>
</div>

<script>
  (function() {
    try {
      var savedKey = localStorage.getItem('sap_dashboard_key');
      var params = new URLSearchParams(window.location.search);
      // 若 localStorage 有儲存 key 且當前 URL 尚未附帶 key 參數，自動嘗試帶入
      if (savedKey && !params.get('key')) {
        params.set('key', savedKey);
        window.location.search = params.toString();
      }
    } catch(e) {}
  })();

  function handleUnlock(e) {
    e.preventDefault();
    var input = document.getElementById('access-key');
    var key = input.value.trim();
    if (!key) return;

    try {
      localStorage.setItem('sap_dashboard_key', key);
    } catch(err) {}

    var params = new URLSearchParams(window.location.search);
    params.set('key', key);
    window.location.search = params.toString();
  }
</script>
</body>
</html>`;
}
