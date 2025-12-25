// 取得最近一個工作日（不包括週末）
function getLatestWorkingDay() {
  let today = new Date();
  let dayOfWeek = today.getDay();

  // 若今天是週六或週日，調整為最近的週五
  if (dayOfWeek === 0) {  // 週日
    today.setDate(today.getDate() - 2);  // 調整為週五
  } else if (dayOfWeek === 6) {  // 週六
    today.setDate(today.getDate() - 1);  // 調整為週五
  }

  // 格式化日期為 YYYYMMDD
  const year = today.getFullYear();
  const month = ('0' + (today.getMonth() + 1)).slice(-2);
  const day = ('0' + today.getDate()).slice(-2);

  return `${year}${month}${day}`;
}

function fetchTodayBalance(stockSymbol, queryDate) {
  Logger.log('查詢標的: ' + stockSymbol);
  Logger.log('查詢日期: ' + queryDate);
  // 設定目標的網址，將日期參數動態添加到 URL 中
  const url = `https://www.twse.com.tw/rwd/zh/marginTrading/TWTA1U?selectType=X&date=${queryDate}&response=html`;

  // 發送請求以取得網頁內容
  const response = UrlFetchApp.fetch(url);
  const html = response.getContentText();

  // 優化過的正則表達式，用來抓取「今日餘額」和「次一營業日限額」
  const regex = new RegExp(`<td>${stockSymbol}</td>(?:\\s*<td>.*?</td>){19}\\s*<td>([\\d,]+)</td>\\s*<td>([\\d,]+)</td>`, 'g');

  const matches = regex.exec(html);

  if (matches && matches.length >= 3) {
    const todayBalance = matches[1];  // 今日餘額
    const nextDayLimit = matches[2];  // 次一營業日限額
    Logger.log('今日餘額: ' + todayBalance);
    Logger.log('次一營業日限額: ' + nextDayLimit);
    return [todayBalance, nextDayLimit];  // 回傳餘額和限額
  } else {
    Logger.log('未找到股票代碼對應的數據');
    return [null, null];
  }
}

function main() {
  const stockSymbol = '00948B';  // 替換成你想要查詢的股票代碼
  const queryDate = getLatestWorkingDay();  // 自動取得最近的工作日

  const result = fetchTodayBalance(stockSymbol, queryDate);
  if (result) {
    Logger.log('今日餘額: ' + result[0]);
    Logger.log('次一營業日限額: ' + result[1]);
  }
}
