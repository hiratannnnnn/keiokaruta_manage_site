// ============================================================
// 大会一覧・カレンダー操作
// ============================================================

// 大会一覧取得（カレンダーシート）
// 行1: メタデータ → スキップ
// 行2: ヘッダー   → スキップ
// 行3〜: データ
// 列: 0=大会名, 2=申込開始, 5=本申込期限, 7=抽選日, 10=本振込期限, 14=大会の日時
function getTournamentList() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    if (!sheet) throw new Error(`「${CONFIG.SHEET_NAMES.CALENDAR}」シートが見つかりません`);

    // 必要な最右列は col 14（O列）なので 16 列だけ取得
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 16).getValues();
    const rows = data.slice(2); // 先頭 2 行スキップ

    const list = rows
      .filter(r => r[0] !== '' && String(r[12]) !== '完了')
      .map(r => ({
        name:             r[0],
        doubleChecked:    String(r[1]) === 'レ',
        announcementHtml: String(r[15] === null || r[15] === undefined ? '' : r[15]),
        date:             formatCell(r[14]),
        applyStart:       formatCell(r[2]),
        applyDeadline:    formatCell(r[5]),
        lottery:          formatCell(r[7]),
        payDeadline:      formatCell(r[10]),
        applyDone:        String(r[6])  === '済',
        payDone:          String(r[11]) === '済',
      }));

    return JSON.stringify(list);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// カレンダーシートの指定列に値を書き込む（大会名で行を特定）
// colOneBased : 書き込む列番号（1-indexed）
function calendarTournamentSiblingRows_(data, sheetName) {
  const baseName = tournamentSheetBaseName_(sheetName);
  return (data || []).slice(2).filter(row =>
    row[0] && tournamentSheetBaseName_(row[0]) === baseName
  );
}

function calendarAggregatedCompletion_(rows, targetName, colOneBased, value) {
  return (rows || []).length > 0 && rows.every(row =>
    String(row[0]) === String(targetName)
      ? String(value || '').trim() === '済'
      : String(row[colOneBased - 1] || '').trim() === '済'
  );
}

function setCalendarColumn(name, colOneBased, value) {
  let apiUpdated = false;
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    if (!sheet) throw new Error(`「${CONFIG.SHEET_NAMES.CALENDAR}」シートが見つかりません`);

    const data = sheet.getRange(1, 1, sheet.getLastRow(), 13).getValues();
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        const baseName = tournamentSheetBaseName_(name);
        const siblingRows = calendarTournamentSiblingRows_(data, name);
        if (colOneBased === 7 || colOneBased === 12) {
          const tournament = taikaiFindTournament_(
            baseName
          );
          const field = colOneBased === 7
            ? 'registration_completed' : 'payment_completed';
          const update = {};
          update[field] = calendarAggregatedCompletion_(
            siblingRows, name, colOneBased, value
          );
          taikaiApiRequest_(
            'PATCH',
            '/tournaments/' + encodeURIComponent(String(tournament.id)),
            update
          );
          apiUpdated = true;
        }
        sheet.getRange(i + 1, colOneBased).setValue(value);
        if (colOneBased === 7 || colOneBased === 12) {
          const writebackErrors = [];
          siblingRows.forEach(row => {
            const siblingName = String(row[0]);
            const tournamentSheet = ss.getSheetByName(siblingName);
            try {
              if (!tournamentSheet) {
                throw new Error('大会シートが見つかりません。');
              }
              const structure = tournamentSheetStructure_(tournamentSheet, false);
              if (structure.version !== 2) return;
              refreshTournamentSheetV2FromApi_(tournamentSheet);
            } catch (writebackError) {
              if (tournamentSheet) {
                try {
                  markTournamentSheetV2SyncState_(
                    tournamentSheet,
                    'pending_sheet',
                    writebackError.message || writebackError
                  );
                } catch (markError) {}
              }
              writebackErrors.push(siblingName + ': ' + writebackError.message);
            }
          });
          if (writebackErrors.length) {
            return JSON.stringify({
              error: 'DBとカレンダーの更新は成功しましたが、大会シートへの書戻しに'
                + '失敗しました。同じ操作を再実行してください: '
                + writebackErrors.join(' / '),
              partial: true,
            });
          }
        }
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: '大会が見つかりません' });
  } catch (e) {
    return JSON.stringify({
      error: (apiUpdated
        ? 'DB更新後にカレンダーの更新に失敗しました。同じ操作を再実行してください: '
        : '') + e.message,
      partial: apiUpdated,
    });
  }
}

// 大会を完了済みにする（カレンダーシートの M 列に「完了」を書く）
function completeTournament(name) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const tournament = taikaiFindTournament_(
      tournamentSheetBaseName_(name)
    );
    if (!tournament.registration_completed || !tournament.payment_completed) {
      return JSON.stringify({
        error: 'DB上で申込処理・大会振込の両方が完了していないため、完了にできません。',
      });
    }

    const tournamentSheet = ss.getSheetByName(name);

    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    if (!sheet) throw new Error(`「${CONFIG.SHEET_NAMES.CALENDAR}」シートが見つかりません`);

    const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    let found = false;
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        sheet.getRange(i + 1, 13).setValue('完了'); // col M = 13
        found = true;
        break;
      }
    }
    if (!found) return JSON.stringify({ error: '大会が見つかりません' });

    // メール履歴は削除しない。大会シートの非表示失敗は完了状態を巻き戻さない。
    let warning = '';
    if (tournamentSheet) {
      try {
        tournamentSheet.hideSheet();
      } catch (hideError) {
        warning = '大会は完了しましたが、シートを非表示にできませんでした: '
          + hideError.message;
      }
    }

    return JSON.stringify({ ok: true, warning: warning });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
