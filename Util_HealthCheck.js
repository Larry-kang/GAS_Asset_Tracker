function runSystemHealthCheck() {
    const context = "健康檢查";
    LogService.info("啟動全面系統診斷...", context);

    let report = "--- [SAP 系統健康狀態報告] ---\n";
    let issues = 0;

    // 1. Core Script Properties Check
    const requiredCoreProps = ["ADMIN_EMAIL"];

    report += "\n[I] 核心設定檢查\n";
    requiredCoreProps.forEach(p => {
        const val = Settings.get(p);
        if (!val) {
            report += `[失敗] 缺少關鍵屬性: ${p}\n`;
            issues++;
        } else {
            report += `[通過] 屬性已設定: ${p}\n`;
        }
    });

    report += "\n[II] 交易所整合憑證檢查\n";
    ExchangeRegistry.getActive().forEach(entry => {
        const credentialStatus = ExchangeRegistry.getCredentialStatus(entry);
        if (credentialStatus.ok) {
            report += `[通過] ${entry.displayName} 憑證完整\n`;
        } else {
            report += `[失敗] ${entry.displayName} 缺少屬性: ${credentialStatus.missing.join(", ")}\n`;
            issues++;
        }
    });

    // 2. Sheet Structure Check
    report += "\n[III] 工作表與結構檢查\n";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contractResults = WorkbookContracts.validateCoreSheets(ss);
    contractResults.forEach(result => {
        if (!result.ok) {
            report += `[失敗] ${result.message}\n`;
            issues++;
        } else {
            report += `[通過] ${result.sheetName} 結構正常\n`;
        }
    });
    const strategyInputResults = WorkbookContracts.validateStrategyInputSheets(ss);
    strategyInputResults.forEach(result => {
        if (!result.ok) {
            report += `[失敗] ${result.message}\n`;
            issues++;
        } else {
            report += `[通過] ${result.sheetName} 策略輸入結構正常\n`;
        }
    });

    // 3. API Connectivity (Latent Check)
    report += "\n[IV] 網路連線診斷\n";
    const tunnelUrl = Settings.get("TUNNEL_URL");
    const proxyPass = Settings.get("PROXY_PASSWORD");
    const targetUrl = tunnelUrl ? `${tunnelUrl}/api/v3/ping` : "https://api.binance.com/api/v3/ping";

    const params = { muteHttpExceptions: true };
    if (proxyPass) {
        params.headers = { 'x-proxy-auth': proxyPass };
    }

    try {
        const res = UrlFetchApp.fetch(targetUrl, params);
        if (res.getResponseCode() === 200) {
            report += `[通過] 幣安 API 連線正常 ${tunnelUrl ? "(透過隧道)" : "(直接連線)"}\n`;
        } else {
            report += `[失敗] 幣安 API 回應錯誤: ${res.getResponseCode()}\n`;
            issues++;
        }
    } catch (e) {
        report += `[失敗] 幣安 API 無法連線 (${e.message})\n`;
        issues++;
    }

    report += "\n----------------------------------\n";
    report += issues === 0 ? "總結狀態: 執行順暢。主權地位穩固。" : `總結狀態: 警告。偵測到 ${issues} 項潛在問題。`;

    LogService.strategic(`健康檢查完成。問題數: ${issues}`, context);

    if (SpreadsheetApp.getUi) {
        SpreadsheetApp.getUi().alert("SAP 系統健康檢查", report, SpreadsheetApp.getUi().ButtonSet.OK);
    }

    return report;
}

