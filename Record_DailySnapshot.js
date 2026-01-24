/**
 * @OnlyCurrentDoc
 * 這是一份自動化每日資產紀錄腳本（儲存格直讀版）。
 * 版本：2.7
 * 日期：2026年1月24日
 * 更新：
 * 1. 加入 SpreadsheetApp.flush() 確保公式計算完成。
 * 2. 優先使用 Context (Fast Path) 數據以提高準確性。
 */

// =================================================================================
// --- 1. 全域設定 ---
// =================================================================================
const SOURCE_SHEET_NAME_V2 = "圖表";
const DESTINATION_SHEET_NAME_V2 = "每日資產價值變化";

// =================================================================================
// --- 2. 主執行函數 (Main Execution Function) ---
// =================================================================================
function autoRecordDailyValues(context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 目的地分頁必須存在
    const destinationSheet = ss.getSheetByName(DESTINATION_SHEET_NAME_V2);
    if (!destinationSheet) throw new Error(`找不到名為 "${DESTINATION_SHEET_NAME_V2}" 的分頁。請先建立此分頁。`);

    // --- 【*** 核心修正 v2.7 - Formula Recalculation ***】 ---
    // 強制重新計算公式，防止 read stale data
    SpreadsheetApp.flush();

    let cashValue, stockValue, cryptoValue, cryptoliabilityValue, liabilityValue, netWorthValue, totalAssetValue;

    if (context && context.portfolioSummary) {
      // [Fast Path] 直接從計算上下文讀取 (更準確，不依賴 UI)
      const pSummary = context.portfolioSummary;

      // 1. 現金與流動性
      cashValue = (pSummary["CASH_TWD"] || 0) + (pSummary["USDT"] || 0) + (pSummary["USDC"] || 0) + (pSummary["BOXX"] || 0);

      // 2. 股票資產 (L2)
      stockValue = context.assetGroups.find(g => g.id === "L2")?.value || 0;

      // 3. 數位資產 (BTC + Others)
      // 使用 L1 (Reserve) + L4 (Misc) 作為數位資產總額
      const l1Value = context.assetGroups.find(g => g.id === "L1")?.value || 0;
      const l4Value = context.assetGroups.find(g => g.id === "L4")?.value || 0;
      cryptoValue = l1Value + l4Value;

      // 4. 指標數據
      totalAssetValue = context.totalGrossAssets;
      netWorthValue = context.netEntityValue;
      liabilityValue = totalAssetValue - netWorthValue;
      cryptoliabilityValue = 0; // 已整合在淨值計算中

      LogService.info(`Snapshot captured via Fast Path (Memory Context)`, 'Snapshot:Mode');

    } else {
      // [Slow Path] 讀取 Sheet (Fallback)
      const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME_V2);
      if (!sourceSheet) throw new Error(`找不到名為 "${SOURCE_SHEET_NAME_V2}" 的分頁(且無自動化數據輸入)。`);

      cashValue = sourceSheet.getRange("B2").getValue();
      stockValue = sourceSheet.getRange("B3").getValue();
      cryptoValue = sourceSheet.getRange("B4").getValue();
      cryptoliabilityValue = sourceSheet.getRange("B5").getValue();
      liabilityValue = sourceSheet.getRange("B6").getValue();
      netWorthValue = sourceSheet.getRange("B7").getValue();
      totalAssetValue = sourceSheet.getRange("D19").getValue();

      LogService.info(`Snapshot captured via Slow Path (UI Cells)`, 'Snapshot:Mode');
    }

    // 步驟二：計算「增幅百分比」
    let growthPercentage = 0;
    const lastRow = destinationSheet.getLastRow();
    if (lastRow > 1) {
      const dateColumnValues = destinationSheet.getRange(1, 1, lastRow).getValues();
      const today = new Date();
      const yesterday = new Date(today);
      yesterday.setDate(today.getDate() - 1);
      const yesterdayString = Utilities.formatDate(yesterday, "Asia/Taipei", "yyyy/MM/dd");

      for (let i = dateColumnValues.length - 1; i >= 0; i--) {
        const cellDate = new Date(dateColumnValues[i][0]);
        if (!isNaN(cellDate.getTime())) {
          const cellDateString = Utilities.formatDate(cellDate, "Asia/Taipei", "yyyy/MM/dd");
          if (cellDateString === yesterdayString) {
            const previousRowIndex = i + 1;
            const previousNetWorth = destinationSheet.getRange(previousRowIndex, 7).getValue();

            if (typeof previousNetWorth === 'number' && previousNetWorth !== 0 && typeof netWorthValue === 'number') {
              growthPercentage = (netWorthValue - previousNetWorth) / previousNetWorth;
            }
            break;
          }
        }
      }
    }

    // 步驟三：準備要寫入的新數據行
    const formattedDate = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
    const newRowData = [
      formattedDate,
      cashValue,
      cryptoValue,
      stockValue,
      growthPercentage,
      Math.abs(liabilityValue) + Math.abs(cryptoliabilityValue),
      netWorthValue,
      totalAssetValue
    ];

    // --- 【*** 核心修正 v2.7 - Upsert Logic ***】 ---
    // 步驟四：寫入或更新數據
    const lastRowIndex = destinationSheet.getLastRow();
    let isUpdateMode = false;

    if (lastRowIndex > 1) {
      // 檢查最後一列的日期
      const lastRowDateVal = destinationSheet.getRange(lastRowIndex, 1).getValue();
      const lastRowDateStr = Utilities.formatDate(new Date(lastRowDateVal), "Asia/Taipei", "yyyy/MM/dd");

      if (lastRowDateStr === formattedDate) {
        // [UPDATE] 如果日期相同，則更新該列
        isUpdateMode = true;
        // 注意：newRowData 的長度必須與目標欄位數匹配
        // 這裡我們只更新前 8 欄 (A-H)，公式欄位不動
        destinationSheet.getRange(lastRowIndex, 1, 1, newRowData.length).setValues([newRowData]);
        LogService.info(`Updated existing record for ${formattedDate}`, 'Snapshot:Upsert');
      }
    }

    if (!isUpdateMode) {
      // [INSERT] 如果日期不同 (或新表)，則新增一列
      if (lastRowIndex === 0) {
        const headers = [
          "日期", "現金", "數位幣資產", "股票資產", "增幅", "負債餘額", "淨資產", "資產總值",
          "最高點差異", "當日加權指數", "與歷史高點差異", "年月", "是否為最後一天"
        ];
        destinationSheet.appendRow(headers);
      }
      destinationSheet.appendRow(newRowData);
      LogService.info(`Inserted new record for ${formattedDate}`, 'Snapshot:Upsert');
    }

    // 步驟五：自動填充公式 (僅在新增模式或強制檢查時執行，但為確保一致性，更新模式也可以檢查)
    // 若為更新模式，公式理應已存在，但為防萬一還是執行一次 Copy
    const targetRowIndex = isUpdateMode ? lastRowIndex : destinationSheet.getLastRow();

    if (targetRowIndex > 2) {
      const previousRowIndex = targetRowIndex - 1;
      // I欄到 M欄 (9-13)
      const formulaRange = destinationSheet.getRange(previousRowIndex, 9, 1, 5);
      // 只有當目標欄位空白時才複製? 或者強制覆蓋? 這裡選擇強制刷新公式
      formulaRange.copyTo(destinationSheet.getRange(targetRowIndex, 9, 1, 5));
    }
    // --- 【*** 修正結束 ***】 ---

  } catch (e) {
    LogService.error(`每日資產紀錄腳本執行失敗: ${e.toString()}`, 'Snapshot:Error');
    // MailApp.sendEmail("YOUR_EMAIL@example.com", "【錯誤】每日資產紀錄腳本失敗", e.toString());
  }
}
