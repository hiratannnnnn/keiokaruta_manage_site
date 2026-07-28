// ============================================================
// 外部DB登録（keiokarutakai）
// ============================================================

// taikai_manage APIへ出場登録する
function connectDb_(date, tournamentName, playerName, email, grade) {
  return taikaiRegisterEntry_(tournamentName, grade, date, playerName, email);
}

// 大会をデータベースに登録する
// kounin : true=公認, false=非公認
function runRegisterDatabase(name, kounin) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, true);
    if (structure.version === 2) {
      return JSON.stringify({
        error: '大会管理データv2ではregisterDatabase操作を廃止しました。年度完全同期を実行してください。',
      });
    }
    const data2 = structure.data;
    const paymentStatusIndex = structure.layout.payment_status_column - 1;
    // 選手行：回答終端より上だけを対象にする。
    const rangeVals = data2.slice(1, structure.response_end_index);
    const playerRows = [];
    for (const row of rangeVals) {
      if (typeof row[2] === 'number') break;
      if (typeof row[2] === 'string' && row[2] !== '' &&
          (row[2].includes(' ') || row[2].includes('　'))) {
        playerRows.push(row);
      }
    }

    // 抽選日を取得（col[N+3]="抽選日" の行の col[N+4]）
    let raffleDate = null;
    const legacyStartColumn = structure.layout.form_id_column;
    if (sheet.getLastColumn() >= legacyStartColumn + 1) {
      sheet.getRange(
        1, legacyStartColumn, sheet.getLastRow(), 2
      ).getValues().forEach(row => {
      if (row[0] === '抽選日' && row[1] !== '未定' && row[1] !== '') {
        raffleDate = (row[1] instanceof Date) ? row[1] : new Date(String(row[1]).replace('など', ''));
      }
      });
    }
    if (!raffleDate) raffleDate = new Date();

    // 級別大会日を取得（col[0] が A-E の行の col[count+2]）
    const gradeRegex = /^[A-E]$/;
    const dates = {};
    Object.keys(structure.grade_rows).forEach(grade => {
      dates[grade] = data2[structure.grade_rows[grade] - 1][paymentStatusIndex];
    });

    // 大会名はAPI上の大会名（級表記なし）を使う。
    // 公認・非公認は級別日程の is_sanctioned で管理する。
    const baseName = name.replace(/[A-Z]+級$/, '');
    const registered = [];
    const failed = [];
    for (const row of playerRows) {
      const payStatus  = String(row[paymentStatusIndex]);
      const isTarget   = payStatus === '' || payStatus === '済' ||
        (payStatus.includes('繰') && payStatus.includes('越')) ||
        payStatus === 'くりこし';
      if (!isTarget) continue;

      const grade     = String(row[4]);
      const gradeDate = dates[grade];
      if (!gradeDate) {
        failed.push({ name: String(row[2]), error: grade + '級の日程が未設定です' });
        continue;
      }

      const gradeDateStr = Utilities.formatDate(
        (gradeDate instanceof Date) ? gradeDate : new Date(String(gradeDate)),
        'JST', 'yyyy-MM-dd'
      );
      const email = String(row[1] || '').trim();
      if (!email) {
        failed.push({ name: String(row[2]), error: 'メールアドレスがありません' });
        continue;
      }
      try {
        connectDb_(gradeDateStr, baseName, String(row[2]), email, grade);
        registered.push(String(row[2]));
      } catch (entryError) {
        failed.push({ name: String(row[2]), error: entryError.message });
      }
    }

    // "registerDatabase" 直下セルに登録結果を書き込む
    if (structure.register_database_row) {
      const status = failed.length
        ? 'DB同期失敗: ' + failed.map(item => item.name + '（' + item.error + '）').join('、')
        : (kounin ? '公認大会として登録済み' : '非公認大会として登録済み');
      sheet.getRange(structure.register_database_row + 1, 1).setValue(status);
    }

    return JSON.stringify({
      ok: failed.length === 0,
      registered: registered,
      failed: failed,
      error: failed.length ? failed.length + '件のDB同期に失敗しました。再実行できます。' : '',
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
