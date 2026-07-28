// ============================================================
// シート・フォーム削除（deleteSheet の Web App 版）
// ============================================================

// フォームをゴミ箱へ移動し、回答先をゴミ箱用スプレッドシートへ変更した上でシートを削除する
function deleteTournament(name) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 管理シートの誤削除を防ぐガード
    const protected_ = [
      CONFIG.SHEET_NAMES.CALENDAR,
      CONFIG.SHEET_NAMES.MEMBERS,
      CONFIG.SHEET_NAMES.MAIL,
    ];
    if (protected_.includes(name)) {
      return JSON.stringify({ error: '管理シートは削除できません' });
    }

    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const formId = tournamentSheetFormId_(structure);
    if (!formId) throw new Error('フォームIDが取得できません');

    // シート名末尾の「ABC級」を除いた大会名がAPI上の大会名。
    // API削除が失敗した場合は、フォーム・シートを残して再試行できるようにする。
    const tournamentName = String(name).replace(/[A-E]+級$/, '');
    const deletedDatabase = taikaiDeleteTournament_(tournamentName);

    // フォームをゴミ箱へ移動
    const form = FormApp.openById(formId);
    DriveApp.getFileById(formId).setTrashed(true);

    // フォームの回答先をゴミ箱用スプレッドシートへ変更
    const trashSs = SpreadsheetApp.openById(CONFIG.TRASH_SPREADSHEET_ID);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, trashSs.getId());

    // シートを削除
    ss.deleteSheet(sheet);

    // カレンダーシートの該当行を削除する
    const calSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    if (calSheet) {
      const calData = calSheet.getRange(1, 1, calSheet.getLastRow(), 1).getValues();
      for (let i = 2; i < calData.length; i++) {
        if (String(calData[i][0]) === name) {
          calSheet.deleteRow(i + 1);
          break;
        }
      }
    }

    return JSON.stringify({ ok: true, database: deletedDatabase });
  } catch (err) {
    return JSON.stringify({ error: err.message });
  }
}
