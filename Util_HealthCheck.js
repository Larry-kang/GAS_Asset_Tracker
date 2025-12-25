function runSystemHealthCheck() {
    const context = "健康檢查";
    LogService.info("啟動全面系統診斷...", context);

    let report = "--- [SAP 系統健康狀態報告] ---\n";
    let issues = 0;

    // 1. Script Properties Check
    const props = PropertiesService.getScriptProperties().getProperties();
    const requiredProps = ["ADMIN_EMAIL", "BINANCE_API_KEY", "BINANCE_API_SECRET", "BITOPRO_API_KEY", "BITOPRO_API_SECRET"];

    report += "\n[I] 憑證權限檢查\n";
    requiredProps.forEach(p => {
        if (!props[p]) {
            report += `[失敗] 缺少關鍵屬性: ${p}\n`;
            issues++;
        } else {
            report += `[通過] 屬性已設定: ${p}\n`;
        }
    });

    // 2. Sheet Structure Check
    report += "\n[II] 工作表與結構檢查\n";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = ["Balance Sheet", "Key Market Indicators", "System_Logs"];
    requiredSheets.forEach(s => {
        const sheet = ss.getSheetByName(s);
        if (!sheet) {
            report += `[失敗] 缺少工作表: ${s}\n`;
            issues++;
        } else {
            report += `[通過] 找到工作表: ${s}\n`;
        }
    });

    // 3. API Connectivity (Latent Check)
    report += "\n[III] 網路連線診斷\n";
    try {
        UrlFetchApp.fetch("https://api.binance.com/api/v3/ping");
        report += "[通過] 幣安 API 連線正常\n";
    } catch (e) {
        report += "[失敗] 幣安 API 無法連線或被阻擋\n";
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
