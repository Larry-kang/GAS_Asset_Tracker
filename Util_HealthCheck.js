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

    // 5. Price Cache Diagnostics
    report += "\n[VI] 價格資料診斷\n";
    try {
        const priceHealth = inspectPriceCacheHealth_(ss);
        if (priceHealth.missingTickers.length > 0) {
            addWarning(`價格暫存缺少資產配置標的: ${priceHealth.missingTickers.join(", ")}`);
        } else {
            addPass("價格暫存涵蓋所有資產配置標的");
        }

        if (priceHealth.blankTickers.length > 0) {
            addWarning(`價格暫存有空白或無效價格: ${priceHealth.blankTickers.join(", ")}`);
        } else {
            addPass("價格暫存沒有空白價格");
        }

        if (priceHealth.staleTickers.length > 0) {
            addWarning(`價格資料可能過期: ${priceHealth.staleTickers.join(", ")}`);
        } else {
            addPass("價格資料新鮮度正常");
        }

        if (priceHealth.lockedTickers.length > 0) {
            addWarning(`手動鎖價列: ${priceHealth.lockedTickers.join(", ")}`);
        } else {
            addPass("沒有偵測到手動鎖價列");
        }
    } catch (e) {
        addFailure(`價格資料診斷失敗: ${e.message || e}`);
    }

    // 6. FX Rate Diagnostics
    report += "\n[VII] 匯率資料診斷\n";
    try {
        const fxHealth = inspectFxRateHealth_(ss);
        if (fxHealth.missingRequiredPairs.length > 0) {
            addFailure(`參數設定缺少必要貨幣對: ${fxHealth.missingRequiredPairs.join(", ")}`);
        } else {
            addPass("參數設定涵蓋所有必要貨幣對");
        }

        if (fxHealth.invalidRequiredPairs.length > 0) {
            addFailure(`必要匯率有空白或無效值: ${fxHealth.invalidRequiredPairs.join(", ")}`);
        } else {
            addPass("必要匯率沒有空白或無效值");
        }

        if (fxHealth.invalidOptionalPairs.length > 0) {
            addWarning(`非必要匯率尚未維護: ${fxHealth.invalidOptionalPairs.join(", ")}`);
        } else {
            addPass("非必要匯率沒有空白或無效值");
        }

        if (fxHealth.blankRequiredTimestampPairs.length > 0) {
            addWarning(`必要匯率缺少更新時間: ${fxHealth.blankRequiredTimestampPairs.join(", ")}`);
        }

        if (fxHealth.blankOptionalTimestampPairs.length > 0) {
            addWarning(`非必要匯率缺少更新時間: ${fxHealth.blankOptionalTimestampPairs.join(", ")}`);
        }

        if (fxHealth.blankRequiredTimestampPairs.length === 0 && fxHealth.blankOptionalTimestampPairs.length === 0) {
            addPass("匯率矩陣更新時間完整");
        }

        if (fxHealth.staleRequiredPairs.length > 0) {
            addWarning(`必要匯率資料可能過期: ${fxHealth.staleRequiredPairs.join(", ")}`);
        } else {
            addPass("必要匯率資料新鮮度正常");
        }

        if (fxHealth.staleOptionalPairs.length > 0) {
            addWarning(`非必要匯率資料可能過期: ${fxHealth.staleOptionalPairs.join(", ")}`);
        } else {
            addPass("非必要匯率資料新鮮度正常");
        }

        if (!fxHealth.fixedUsdTwd.ok) {
            addFailure(fxHealth.fixedUsdTwd.message);
        } else {
            addPass(fxHealth.fixedUsdTwd.message);
        }

        if (fxHealth.fixedUsdTwd.staleMessage) {
            addWarning(fxHealth.fixedUsdTwd.staleMessage);
        }
    } catch (e) {
        addFailure(`匯率資料診斷失敗: ${e.message || e}`);
    }

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

function inspectPriceCacheHealth_(ss) {
    const now = new Date();
    const staleHours = (Config.THRESHOLDS && Config.THRESHOLDS.PRICE_HEALTH_STALE_HOURS) || 24;
    const staleMs = staleHours * 60 * 60 * 1000;
    const rows = PriceCacheRepo.readRows(ss);
    const cacheLookup = {};
    const blankTickers = [];
    const staleTickers = [];
    const lockedTickers = [];

    rows.forEach(row => {
        const ticker = normalizeHealthPriceTicker_(row.ticker);
        if (!ticker) return;
        cacheLookup[ticker] = true;

        const updatedAt = parseHealthPriceDate_(row.updatedAt);
        if (updatedAt && updatedAt.getTime() > now.getTime() + 1000) {
            lockedTickers.push(`${ticker} until ${formatHealthCheckDate_(updatedAt)}`);
            return;
        }

        if (!isHealthPositiveNumber_(row.price)) {
            blankTickers.push(ticker);
            return;
        }

        if (!updatedAt || now.getTime() - updatedAt.getTime() > staleMs) {
            staleTickers.push(ticker);
        }
    });

    const candidateResult = typeof readAssetAllocationPriceCandidates_ === 'function'
        ? readAssetAllocationPriceCandidates_(ss, {})
        : { rows: [], warnings: [] };
    const missingTickers = [];
    (candidateResult.rows || []).forEach(candidate => {
        const ticker = normalizeHealthPriceTicker_(candidate.ticker);
        if (ticker && !cacheLookup[ticker] && missingTickers.indexOf(ticker) < 0) {
            missingTickers.push(ticker);
        }
    });

    return {
        missingTickers: missingTickers,
        blankTickers: blankTickers,
        staleTickers: staleTickers,
        lockedTickers: lockedTickers
    };
}

function normalizeHealthPriceTicker_(value) {
    return String(value == null ? "" : value).trim().replace(/\s+/g, "").toUpperCase();
}

function parseHealthPriceDate_(value) {
    if (!value) return null;
    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value;
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
}

function isHealthPositiveNumber_(value) {
    const num = parseFloat(value);
    return isFinite(num) && num > 0;
}

function formatHealthCheckDate_(value) {
    try {
        return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    } catch (e) {
        return value.toISOString ? value.toISOString() : String(value);
    }
}

function inspectFxRateHealth_(ss) {
    const now = new Date();
    const staleHours = (Config.THRESHOLDS && Config.THRESHOLDS.FX_RATE_STALE_HOURS) || 24;
    const staleMs = staleHours * 60 * 60 * 1000;
    const settingsOptions = {};
    const matrix = SettingsMatrixRepo.readMatrix(ss, settingsOptions);
    const pairResult = SettingsMatrixRepo.listPairs(ss, settingsOptions);
    const requiredPairsResult = typeof collectRequiredCurrencyPairs_ === 'function'
        ? collectRequiredCurrencyPairs_(ss, { settingsSheetName: SettingsMatrixRepo.DEFAULT_SHEET_NAME })
        : { pairs: [] };
    const requiredPairs = buildHealthRequiredFxPairs_(requiredPairsResult.pairs);
    const requiredPairLookup = buildHealthFxPairLookup_(requiredPairs);
    const missingRequiredPairs = [];
    const invalidRequiredPairs = [];
    const invalidOptionalPairs = [];
    const blankRequiredTimestampPairs = [];
    const blankOptionalTimestampPairs = [];
    const staleRequiredPairs = [];
    const staleOptionalPairs = [];

    requiredPairs.forEach(pair => {
        const from = normalizeHealthCurrency_(pair.from);
        const to = normalizeHealthCurrency_(pair.to);
        if (!from || !to) return;
        if (!matrix.sourceRows[from] || !matrix.targetColumns[to]) {
            missingRequiredPairs.push(`${from}/${to}`);
        }
    });

    (pairResult.pairs || []).forEach(pair => {
        const pairLabel = `${pair.from}/${pair.to}`;
        const timestamp = parseHealthPriceDate_(pair.currentTimestamp);
        const isRequired = !!requiredPairLookup[buildHealthFxPairKey_(pair.from, pair.to)];

        if (!isHealthPositiveNumber_(pair.currentValue)) {
            if (isRequired) {
                invalidRequiredPairs.push(pairLabel);
            } else {
                invalidOptionalPairs.push(pairLabel);
            }
            return;
        }

        if (!timestamp) {
            if (isRequired) {
                blankRequiredTimestampPairs.push(pairLabel);
            } else {
                blankOptionalTimestampPairs.push(pairLabel);
            }
            return;
        }

        if (now.getTime() - timestamp.getTime() > staleMs) {
            const label = `${pairLabel} last updated ${formatHealthCheckDate_(timestamp)}`;
            if (isRequired) {
                staleRequiredPairs.push(label);
            } else {
                staleOptionalPairs.push(label);
            }
        }
    });

    return {
        missingRequiredPairs: missingRequiredPairs,
        invalidRequiredPairs: invalidRequiredPairs,
        invalidOptionalPairs: invalidOptionalPairs,
        blankRequiredTimestampPairs: blankRequiredTimestampPairs,
        blankOptionalTimestampPairs: blankOptionalTimestampPairs,
        staleRequiredPairs: staleRequiredPairs,
        staleOptionalPairs: staleOptionalPairs,
        fixedUsdTwd: inspectFixedUsdTwdReference_(matrix.sheet, now, staleMs)
    };
}

function buildHealthRequiredFxPairs_(pairs) {
    const required = [];
    const lookup = {};

    (pairs || []).forEach(pair => {
        const from = normalizeHealthCurrency_(pair.from);
        const to = normalizeHealthCurrency_(pair.to);
        if (!from || !to) return;
        const key = buildHealthFxPairKey_(from, to);
        if (lookup[key]) return;
        lookup[key] = true;
        required.push({ from: from, to: to });
    });

    const mandatoryKey = buildHealthFxPairKey_("USD", "TWD");
    if (!lookup[mandatoryKey]) {
        required.push({ from: "USD", to: "TWD" });
    }

    return required;
}

function buildHealthFxPairLookup_(pairs) {
    const lookup = {};
    (pairs || []).forEach(pair => {
        lookup[buildHealthFxPairKey_(pair.from, pair.to)] = true;
    });
    return lookup;
}

function buildHealthFxPairKey_(from, to) {
    return `${normalizeHealthCurrency_(from)}__${normalizeHealthCurrency_(to)}`;
}

function inspectFixedUsdTwdReference_(sheet, now, staleMs) {
    const b1 = normalizeHealthCurrency_(sheet.getRange("B1").getValue());
    const a4 = normalizeHealthCurrency_(sheet.getRange("A4").getValue());
    const b4 = sheet.getRange("B4").getValue();
    const c4 = sheet.getRange("C4").getValue();
    const errors = [];

    if (b1 !== "TWD") errors.push(`B1 expected TWD got ${b1 || "(blank)"}`);
    if (a4 !== "USD") errors.push(`A4 expected USD got ${a4 || "(blank)"}`);
    if (!isHealthPositiveNumber_(b4)) errors.push(`B4 expected positive USD/TWD rate got ${b4 || "(blank)"}`);

    const timestamp = parseHealthPriceDate_(c4);
    let staleMessage = "";
    if (timestamp && now.getTime() - timestamp.getTime() > staleMs) {
        staleMessage = `固定位置 USD/TWD 可能過期: '參數設定'!$B$4 last updated ${formatHealthCheckDate_(timestamp)}`;
    }

    if (!timestamp) {
        staleMessage = "固定位置 USD/TWD 缺少更新時間: '參數設定'!$C$4";
    }

    if (errors.length > 0) {
        return {
            ok: false,
            message: `固定位置 USD/TWD 語意異常: ${errors.join("; ")}`,
            staleMessage: staleMessage
        };
    }

    return {
        ok: true,
        message: "固定位置 USD/TWD 語意正常: '參數設定'!$B$4",
        staleMessage: staleMessage
    };
}

function normalizeHealthCurrency_(value) {
    return String(value == null ? "" : value).trim().toUpperCase();
}
