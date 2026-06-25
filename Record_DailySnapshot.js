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
const ASSET_ALLOCATION_SHEET_NAME_V2 = "資產配置";
const STOCK_PLEDGE_SHEET_NAME_V2 = "質押紀錄表";

function toSnapshotNumber_(value) {
  const num = Number(value);
  return isFinite(num) ? num : 0;
}

function buildSnapshotSummaryFromAllocationRows_(allocationRows, stockPledgeInterestAdjustment) {
  const summary = {
    cashValue: 0,
    stockValue: 0,
    cryptoValue: 0,
    cryptoliabilityValue: 0,
    liabilityValue: 0,
    netWorthValue: 0,
    totalAssetValue: 0
  };
  const stockPledgeOffset = toSnapshotNumber_(stockPledgeInterestAdjustment);

  (allocationRows || []).forEach((row) => {
    const assetType = String(row[1] || '').trim();
    const ticker = String(row[2] || '').trim().toUpperCase();
    const currentValue = toSnapshotNumber_(row[8]);

    if (!assetType && !ticker) return;

    if (assetType === '現金') {
      summary.cashValue += currentValue;
      return;
    }

    if (assetType === '股票' && ticker !== 'IBIT') {
      summary.stockValue += currentValue;
      return;
    }

    if (assetType === '數位幣' || assetType === '數位穩定幣' || ticker === 'IBIT') {
      if (currentValue >= 0) summary.cryptoValue += currentValue;
      else summary.cryptoliabilityValue += currentValue;
      return;
    }

    if (assetType === '貸款' || assetType === '質押' || assetType === '信用卡費') {
      summary.liabilityValue += currentValue;
    }
  });

  summary.liabilityValue -= stockPledgeOffset;
  summary.totalAssetValue = summary.cashValue + summary.stockValue + summary.cryptoValue;
  summary.netWorthValue =
    summary.totalAssetValue +
    summary.cryptoliabilityValue +
    summary.liabilityValue;

  return summary;
}

function readSnapshotSummaryFromSheets_(ss) {
  const assetAllocationSheet = ss.getSheetByName(ASSET_ALLOCATION_SHEET_NAME_V2);
  if (!assetAllocationSheet) {
    throw new Error(`找不到名為 "${ASSET_ALLOCATION_SHEET_NAME_V2}" 的分頁。`);
  }

  const stockPledgeSheet = ss.getSheetByName(STOCK_PLEDGE_SHEET_NAME_V2);
  const allocationValues = assetAllocationSheet.getDataRange().getValues();
  const allocationRows = allocationValues.length > 3 ? allocationValues.slice(3) : [];
  const stockPledgeInterestAdjustment = stockPledgeSheet ? stockPledgeSheet.getRange("E2").getValue() : 0;

  return buildSnapshotSummaryFromAllocationRows_(allocationRows, stockPledgeInterestAdjustment);
}

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

    try {
      const sheetSummary = readSnapshotSummaryFromSheets_(ss);
      cashValue = sheetSummary.cashValue;
      stockValue = sheetSummary.stockValue;
      cryptoValue = sheetSummary.cryptoValue;
      cryptoliabilityValue = sheetSummary.cryptoliabilityValue;
      liabilityValue = sheetSummary.liabilityValue;
      netWorthValue = sheetSummary.netWorthValue;
      totalAssetValue = sheetSummary.totalAssetValue;

      LogService.info(`Snapshot captured via Asset Allocation parity mode`, 'Snapshot:Mode');
    } catch (sheetError) {
      const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME_V2);
      if (!sourceSheet) throw new Error(`找不到名為 "${SOURCE_SHEET_NAME_V2}" 的分頁，且資產配置口徑重算失敗: ${sheetError.message}`);

      cashValue = sourceSheet.getRange("B2").getValue();
      stockValue = sourceSheet.getRange("B3").getValue();
      cryptoValue = sourceSheet.getRange("B4").getValue();
      cryptoliabilityValue = sourceSheet.getRange("B5").getValue();
      liabilityValue = sourceSheet.getRange("B6").getValue();
      netWorthValue = sourceSheet.getRange("B7").getValue();
      totalAssetValue = sourceSheet.getRange("D19").getValue();

      LogService.warn(`Asset Allocation parity mode failed, fallback to chart cells: ${sheetError.message}`, 'Snapshot:Mode');
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
