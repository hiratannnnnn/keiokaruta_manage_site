// ============================================================
// 大会結果検索（keiokarutakai API）
// ============================================================

// name   : 選手名（全角スペースは半角に変換）
// dateStr: 日付文字列 yyyy-MM-dd（省略可）
function getTournamentResults(name, dateStr) {
  let url = CONFIG.KARUTA_SEARCH_URL;
  const date = dateStr || Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  url += '?name=' + encodeURIComponent(name.replace('　', ' '));
  url += '&date=' + encodeURIComponent(date);

  try {
    const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    return JSON.stringify({
      url:    url,
      status: response.getResponseCode(),
      body:   response.getContentText(),
    });
  } catch (e) {
    return JSON.stringify({ url, error: e.message });
  }
}
