/**
 * Lib_SyncManager.js
 * 
 * [共用核心] 資產同步管理器 (Asset Sync Manager)
 * 目的：標準化所有交易所的「資料合併」、「寫入 Sheet」、「日誌記錄」與「錯誤處理」。
 * 
 * 主要功能：
 * 1. writeToSheet: 自動處理標題、ISO 時間、排序、清空舊資料、無資產顯示。
 * 2. mergeAssets: 穩健合併多個來源 (Map 或 Object)，與型別檢查。
 * 3. log: 統一透過 LogService 寫入 System_Logs，確保使用者看得到。
 */

const SyncManager = {

    /**
     * [核心] 執行同步任務的包裝器
     * 自動捕捉錯誤並寫入 LogService
     * @param {string} moduleName - 模組名稱 (如 "Binance", "Okx")
     * @param {Function} taskFn - 主要執行的函式
     */
    run: function (moduleName, taskFn) {
        try {
            this.log("INFO", `[${moduleName}] 開始同步...`, moduleName);
            taskFn();
            this.log("INFO", `[${moduleName}] 同步作業結束`, moduleName);
        } catch (e) {
            this.log("ERROR", `執行崩潰: ${e.message}`, moduleName);
            // 也顯示 Toast 讓當下操作者知道
            try { SpreadsheetApp.getActiveSpreadsheet().toast(`${moduleName} Error: ${e.message}`); } catch (ignore) { }
            console.error(e);
        }
    },

    /**
     * [核心] 通用資產合併
     * 接受任意數量的 Map 或 Object，合併為單一 Map
     * @param {...(Map|Object)} sources - 資料來源
     * @returns {Map<string, number>} 合併後的資產 Map (Key: Currency, Value: Amount)
     */
    mergeAssets: function (...sources) {
        const combined = new Map();
        let totalSources = 0;
        let totalItems = 0;

        sources.forEach((source, index) => {
            if (!source) return;
            totalSources++;
            let count = 0;

            // Case A: Map
            if (source instanceof Map) {
                source.forEach((val, key) => {
                    const current = combined.get(key) || 0;
                    combined.set(key, current + val);
                    count++;
                });
            }
            // Case B: Object
            else if (typeof source === 'object') {
                Object.keys(source).forEach(key => {
                    const val = source[key];
                    const current = combined.get(key) || 0;
                    combined.set(key, current + val);
                    count++;
                });
            }
            else {
                console.warn(`[SyncManager] Merge source #${index} is neither Map nor Object: ${typeof source}`);
            }
            totalItems += count;
        });

        console.log(`[SyncManager] Merged ${totalSources} sources, total ${totalItems} items into ${combined.size} unique assets.`);
        return combined;
    },

    /**
     * [核心] 標準化寫入 Sheet
     * @param {SpreadsheetApp.Spreadsheet} ss 
     * @param {string} sheetName - Target Sheet Name
     * @param {string[]} headers - e.g. ['Currency', 'Amount', 'Last Updated']
     * @param {Map<string, number>} dataMap - Asset Data
     */
    writeToSheet: function (ss, sheetName, headers, dataMap) {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) sheet = ss.insertSheet(sheetName);

        // 1. 強制重寫標題 (Enforce Headers)
        // 確保標題列樣式統一
        if (headers && headers.length > 0) {
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
        }

        // 2. 準備資料列
        const rows = [];
        dataMap.forEach((val, key) => {
            // 過濾極小餘額 (dust)
            if (val > 0.00000001) {
                rows.push([key, val]);
            }
        });

        // 3. 排序 (金額大到小)
        rows.sort((a, b) => b[1] - a[1]);

        // 4. 計算時間戳記 (ISO 8601)
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

        // 5. 清除舊資料 (Row 2 開始) & 寫入
        // 先清除 A2:Z (假設最大範圍)
        const maxRows = sheet.getMaxRows();
        const maxCols = sheet.getMaxColumns();
        if (maxRows > 1) {
            sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
        }

        // 寫入新資料
        if (rows.length > 0) {
            // 寫入資產數據 (Col 1, 2)
            sheet.getRange(2, 1, rows.length, 2).setValues(rows);

            // 寫入時間戳記 (預設最後一個 Header 對應的欄位)
            // 如果 headers 有 3 欄，時間就寫在 C2 (Row 2, Col 3)
            // 使用者可透過 headers 長度控制時間寫在哪，通常是最後一欄
            if (headers) {
                const timeCol = headers.length;
                // 這裡假設只需要在第一筆資料旁顯示時間，或是整列填滿?
                // 原始需求通常只在第一列或每一列顯示。
                // 為了乾淨，我們只在第二行最後一欄顯示一次，或者每一行都顯示?
                // 舊邏輯是只在 C2 顯示一次。我們維持這個行為。
                sheet.getRange(2, timeCol).setValue(timestamp);
            }

            this.log("INFO", `已寫入 ${rows.length} 筆資產至 ${sheetName}`, "SyncManager");
            ss.toast(`${sheetName} 更新成功 (${rows.length} 幣種)`);
        } else {
            // 無資產 (Empty State)
            if (headers) {
                const emptyRow = new Array(headers.length).fill("");
                emptyRow[0] = "No Assets Found (Check Logs)";
                emptyRow[1] = 0;
                emptyRow[headers.length - 1] = timestamp;

                sheet.getRange(2, 1, 1, headers.length).setValues([emptyRow]);
            }
            this.log("WARNING", `${sheetName} 無資產 (No Assets Found)`, "SyncManager");
            ss.toast(`${sheetName} 無資產`);
        }
    },

    /**
     * [核心] 統一日誌介面
     * 同時寫入 Console (給開發者) 與 System_Logs (給使用者)
     */
    log: function (level, message, context = "SyncManager") {
        // 1. Console
        console.log(`[${level}] [${context}] ${message}`);

        // 2. System_Logs (若 LogService 存在)
        if (typeof LogService !== 'undefined' && LogService.log) {
            LogService.log(level, message, context);
        }
    }
};
