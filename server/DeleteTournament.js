// ============================================================
// シート・フォーム削除（deleteSheet の Web App 版）
// ============================================================

// フォームをゴミ箱へ移動し、回答先をゴミ箱用スプレッドシートへ変更した上でシートを削除する
function deleteTournament(name) {
  let apiUpdated = false;
  let deletedDatabase = null;
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

    const tournamentName = tournamentSheetBaseName_(name);
    const grades = tournamentSheetDeclaredGrades_(name);
    const siblingSheets = ss.getSheets().filter(candidate =>
      candidate.getName() !== name
      && tournamentSheetBaseName_(candidate.getName()) === tournamentName
      && /[A-E]+級$/.test(candidate.getName())
    );

    // API削除が失敗した場合は、Google側を一切削除せず再試行可能にする。
    if (structure.version === 2) {
      const tournamentId = String(
        structure.management.metadata['tournament ID'] || ''
      ).trim();
      const scheduleIds = grades.map(grade => {
        const schedule = structure.management.schedules[grade];
        return schedule ? String(schedule.row[3] || '').trim() : '';
      });
      if (!tournamentId || scheduleIds.some(id => !id)) {
        throw new Error(
          '大会IDまたは担当級のschedule IDが不足しています。'
          + '先に完全同期を実行してください。'
        );
      }
      deletedDatabase = taikaiDeleteTournamentSchedules_(
        tournamentId, scheduleIds
      );
    } else {
      if (siblingSheets.length) {
        throw new Error(
          '複数フォーム大会の旧シートは個別削除できません。'
          + '先に大会シートv2へ移行してください。'
        );
      }
      deletedDatabase = taikaiDeleteTournament_(tournamentName);
    }
    apiUpdated = true;

    // フォームは回答先を変更してからゴミ箱へ移す。再実行時はゴミ箱済みなら省略する。
    const formFile = DriveApp.getFileById(formId);
    if (!formFile.isTrashed()) {
      const form = FormApp.openById(formId);
      const trashSs = SpreadsheetApp.openById(CONFIG.TRASH_SPREADSHEET_ID);
      form.setDestination(FormApp.DestinationType.SPREADSHEET, trashSs.getId());
      formFile.setTrashed(true);
    }

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

    // フォーム固有のメール管理行だけを削除する。
    const mailSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (mailSheet && mailSheet.getLastRow() >= 3) {
      const mailRows = mailSheet.getRange(
        3, 1, mailSheet.getLastRow() - 2, 2
      ).getValues();
      for (let index = mailRows.length - 1; index >= 0; index--) {
        if (String(mailRows[index][0]) === tournamentName
            && String(mailRows[index][1]) === grades.join('') + '級') {
          mailSheet.deleteRow(index + 3);
        }
      }
    }

    // 回答シートを最後に削除し、途中失敗時の再試行情報を残す。
    ss.deleteSheet(sheet);

    return JSON.stringify({
      ok: true,
      database: deletedDatabase,
      tournament_deleted: Boolean(
        deletedDatabase && deletedDatabase.tournament_deleted
      ),
    });
  } catch (err) {
    return JSON.stringify({
      error: (apiUpdated
        ? 'DBの日程削除は成功しましたが、Google側の後処理に失敗しました。'
          + '同じ削除操作を再実行してください: '
        : '') + err.message,
      partial: apiUpdated,
      database: deletedDatabase,
    });
  }
}
