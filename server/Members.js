// ============================================================
// 名簿関連
// ============================================================

// 名簿取得（スプレッドシート）
// シート名「名簿」の 1 行目をヘッダーとして扱う
function getMembers() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MEMBERS);
    if (!sheet) throw new Error(`「${CONFIG.SHEET_NAMES.MEMBERS}」シートが見つかりません`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify([]);

    const headers = data[0];
    const members = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

    return JSON.stringify(members);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
