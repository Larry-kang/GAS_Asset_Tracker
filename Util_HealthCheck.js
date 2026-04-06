function runSystemHealthCheck() {
    const context = "健康檢查";
    LogService.info("啟動全面系統診斷...", context);

    let report = "--- [SAP 系統健康狀態報告] ---\n";
    let issues = 0;
    let warnings = 0;
    const addWarning = function (message) {
        report += `[注意] ${message}\n`;
        warnings++;
    };
    const addFailure = function (message) {
        report += `[失敗] ${message}\n`;
        issues++;
    };
    const addPass = function (message) {
        report += `[通過] ${message}\n`;
    };

    // 1. Core Script Properties Check
    const requiredCoreProps = ["ADMIN_EMAIL"];

    report += "\n[I] 核心設定檢查\n";
    requiredCoreProps.forEach(p => {
        const val = Settings.get(p);
        if (!val) {
            addFailure(`缺少關鍵屬性: ${p}`);
        } else {
            addPass(`屬性已設定: ${p}`);
        }
    });

    report += "\n[II] 交易所整合憑證檢查\n";
    ExchangeRegistry.getActive().forEach(entry => {
        const credentialStatus = ExchangeRegistry.getCredentialStatus(entry);
        if (credentialStatus.ok) {
            addPass(`${entry.displayName} 憑證完整`);
        } else {
            addFailure(`${entry.displayName} 缺少屬性: ${credentialStatus.missing.join(", ")}`);
        }
    });

    // 2. Sheet Structure Check
    report += "\n[III] 工作表與結構檢查\n";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contractResults = WorkbookContracts.validateCoreSheets(ss);
    contractResults.forEach(result => {
        if (!result.ok) {
            addFailure(result.message);
        } else {
            addPass(`${result.sheetName} 結構正常`);
        }
    });
    const strategyInputResults = WorkbookContracts.validateStrategyInputSheets(ss);
    strategyInputResults.forEach(result => {
        if (!result.ok) {
            addFailure(result.message);
        } else {
            addPass(`${result.sheetName} 策略輸入結構正常`);
        }
    });

    // 3. Bridge / Tunnel Diagnostics
    report += "\n[IV] Bridge / Tunnel 診斷\n";
    const tunnelUrl = Settings.get("TUNNEL_URL");
    const proxyPass = Settings.get("PROXY_PASSWORD");
    const bybitEnabled = ExchangeRegistry.getCredentialStatus(ExchangeRegistry.findByFunctionName('getBybitBalance')).ok;

    if (!tunnelUrl) {
        addFailure("缺少 TUNNEL_URL，無法做共享 tunnel 診斷。");
    } else {
        addPass(`共享 Tunnel URL 已設定: ${maskUrlForReport_(tunnelUrl)}`);
    }

    if (!proxyPass) {
        addFailure("缺少 PROXY_PASSWORD，無法驗證共享 relay。");
    }

    if (tunnelUrl && proxyPass) {
        const healthRes = fetchHealthEndpoint_(tunnelUrl, proxyPass, '/healthz');
        if (healthRes.ok) {
            addPass(`共享 Bridge 健康檢查正常${healthRes.message ? ` (${healthRes.message})` : ""}`);
        } else {
            addWarning(`共享 Bridge /healthz 未回應正常，改由 relay 端點驗證 (${healthRes.message})`);
        }

        const binanceRes = fetchHealthEndpoint_(tunnelUrl, proxyPass, '/api/v3/ping');
        if (binanceRes.ok) {
            addPass("幣安 API 連線正常 (透過隧道)");
        } else {
            addFailure(`幣安 API 無法連線 (${binanceRes.message})`);
        }

        if (bybitEnabled) {
            const bybitRes = fetchHealthEndpoint_(tunnelUrl, proxyPass, '/bybit/v5/market/time');
            if (bybitRes.ok) {
                addPass("Bybit relay 連線正常 (共享 tunnel)");
            } else {
                addWarning(`Bybit relay 未就緒，可能目前為 V1 模式或 V2 未接手 (${bybitRes.message})`);
            }
        }
    }

    // 4. Sync Freshness Diagnostics
    report += "\n[V] 同步新鮮度診斷\n";
    const staleHours = (Config.THRESHOLDS && Config.THRESHOLDS.SYNC_STATUS_STALE_HOURS) || 24;
    const statusRows = SyncStatusRepo.readAll(ss);
    ExchangeRegistry.getActive().forEach(entry => {
        const status = statusRows.filter(row => row.exchange === entry.moduleName)[0] || null;
        if (!status) {
            addWarning(`${entry.displayName} 尚未建立 Sync_Status 紀錄。`);
            return;
        }

        if (status.status !== 'COMPLETE') {
            addFailure(`${entry.displayName} 最近一次同步狀態為 ${status.status || 'UNKNOWN'}${status.message ? ` | ${status.message}` : ""}`);
            return;
        }

        if (!status.lastSuccess) {
            addFailure(`${entry.displayName} 狀態為 COMPLETE，但缺少 Last Success。`);
            return;
        }

        const ageHours = (Date.now() - status.lastSuccess.getTime()) / 36e5;
        const lastSuccessLabel = Utilities.formatDate(status.lastSuccess, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        if (ageHours > staleHours) {
            addWarning(`${entry.displayName} 最近成功同步距今 ${ageHours.toFixed(1)} 小時 (${lastSuccessLabel})`);
        } else {
            addPass(`${entry.displayName} 最近成功同步於 ${lastSuccessLabel} (${status.rows} rows)`);
        }
    });

    report += "\n----------------------------------\n";
    if (issues > 0) {
        report += `總結狀態: 警告。偵測到 ${issues} 項潛在問題。`;
    } else if (warnings > 0) {
        report += `總結狀態: 注意。偵測到 ${warnings} 項提醒。`;
    } else {
        report += "總結狀態: 執行順暢。主權地位穩固。";
    }

    LogService.strategic(`健康檢查完成。問題數: ${issues}, 提醒數: ${warnings}`, context);

    if (SpreadsheetApp.getUi) {
        SpreadsheetApp.getUi().alert("SAP 系統健康檢查", report, SpreadsheetApp.getUi().ButtonSet.OK);
    }

    return report;
}

function fetchHealthEndpoint_(baseUrl, proxyPass, path) {
    const url = `${String(baseUrl || '').replace(/\/$/, '')}${path}`;
    const options = {
        muteHttpExceptions: true,
        headers: proxyPass ? { 'x-proxy-auth': proxyPass } : {}
    };

    try {
        const res = UrlFetchApp.fetch(url, options);
        const code = res.getResponseCode();
        const body = String(res.getContentText() || '').trim();
        if (code !== 200) {
            return { ok: false, message: `HTTP ${code}` };
        }

        if (path === '/bybit/v5/market/time') {
            try {
                const data = JSON.parse(body);
                return data && String(data.retCode) === '0'
                    ? { ok: true, message: "retCode=0" }
                    : { ok: false, message: `Unexpected Bybit payload` };
            } catch (e) {
                return { ok: false, message: "Invalid JSON" };
            }
        }

        if (path === '/healthz') {
            try {
                const data = JSON.parse(body);
                const mode = data.mode || data.bridge || data.name || "";
                return { ok: true, message: mode ? String(mode) : "" };
            } catch (e) {
                return { ok: true, message: "" };
            }
        }

        return { ok: true, message: "" };
    } catch (e) {
        return { ok: false, message: e.message };
    }
}

function maskUrlForReport_(url) {
    const text = String(url || '').trim();
    if (!text) return '';
    return text.replace(/^https?:\/\//i, '');
}

