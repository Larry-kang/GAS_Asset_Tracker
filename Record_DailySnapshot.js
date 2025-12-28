/**
 * @OnlyCurrentDoc
 * 這是一份自動化每日資產紀錄腳本（儲存格直讀版）。
 * 版本：2.6
 * 日期：2025年8月5日
 * 更新：新增「自動公式填充」功能。在寫入每日數據後，自動將後續欄位的公式向下填充。
 */

// =================================================================================
// --- 1. 全域設定 ---
// =================================================================================
const SOURCE_SHEET_NAME_V2 = "圖表";
const DESTINATION_SHEET_NAME_V2 = "每日資產價值變化";

// =================================================================================
// --- 2. 主執行函數 (Main Execution Function) ---
// =================================================================================
function autoRecordDailyValues() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME_V2);
    const destinationSheet = ss.getSheetByName(DESTINATION_SHEET_NAME_V2);

    if (!sourceSheet) throw new Error(`找不到名為 "${SOURCE_SHEET_NAME_V2}" 的分頁。`);
    if (!destinationSheet) throw new Error(`找不到名為 "${DESTINATION_SHEET_NAME_V2}" 的分頁。請先建立此分頁。`);

    // 步驟一：直接從指定儲存格讀取數據
    const cashValue = sourceSheet.getRange("B2").getValue();
    const stockValue = sourceSheet.getRange("B3").getValue();
    const cryptoValue = sourceSheet.getRange("B4").getValue();
    const cryptoliabilityValue = sourceSheet.getRange("B5").getValue();
    const liabilityValue = sourceSheet.getRange("B6").getValue();
    const netWorthValue = sourceSheet.getRange("B7").getValue();
    const totalAssetValue = sourceSheet.getRange("D19").getValue();

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
        Logger.log(`[Snapshot] Updated existing record for ${formattedDate}`);
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
      Logger.log(`[Snapshot] Inserted new record for ${formattedDate}`);
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
    Logger.log(`每日資產紀錄腳本執行失敗: ${e.toString()}`);
    // MailApp.sendEmail("YOUR_EMAIL@example.com", "【錯誤】每日資產紀錄腳本失敗", e.toString());
  }
}