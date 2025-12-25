/**
 * JSON數據解析引擎 (IMPORTJSON Custom Function)
 * 這個強大的工具，能讓您的試算表直接讀取和解析來自專業API的JSON格式數據。
 * @param {string} url 要獲取 JSON 數據的 URL。
 * @param {string} query XPath 風格的查詢字串，用於導航 JSON 數據。
 * @return 獲取的數據。
 * @customfunction
 */
function IMPORTJSON(url, query) {
  try {
    var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    var json = response.getContentText();
    var data = JSON.parse(json);

    var parts = query.split("/");
    for (var i = 0; i < parts.length; i++) {
      if (parts[i] === "") continue;
      data = data[parts[i]];
    }

    if (typeof data === 'undefined') {
      return "查詢結果為 undefined。";
    } else if (typeof data === 'object') {
      // 如果數據是對象，則將其作為二維數組返回給工作表。
      var headers = Object.keys(data);
      var values = Object.values(data);
      return [headers, values];
    } else {
      // 如果數據是單一值，直接返回它。
      return data;
    }
  } catch (e) {
    return "錯誤: " + e.toString();
  }
}