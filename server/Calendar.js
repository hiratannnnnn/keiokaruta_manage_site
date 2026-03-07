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
// skipMinus   : col=12 済入力時にマイナストランザクションを計上しない（デポジット完了用）
function setCalendarColumn(name, colOneBased, value, skipMinus) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    if (!sheet) throw new Error(`「${CONFIG.SHEET_NAMES.CALENDAR}」シートが見つかりません`);

    const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][0]) === name) {
        sheet.getRange(i + 1, colOneBased).setValue(value);

        // 振込済み（col 12）が取り消された場合、マイナストランザクションを削除
        if (colOneBased === 12 && value === '') {
          removeSuitouNegTxByReason_(ss, name + '　参加費');
        }

        // 振込済み（col 12）が「済」になった場合、大会シートの済参加者にマイナス計上
        if (colOneBased === 12 && value === '済' && !skipMinus) {
          const tournSheet = ss.getSheetByName(name);
          if (tournSheet) {
            const allData    = tournSheet.getDataRange().getValues();
            const N          = getSuitouN_(allData[0]);
            if (N != null) {
              const formEndIdx = getSuitouFormEndIdx_(allData);
              const feeMap     = getSuitouFeeMap_(allData, formEndIdx);
              for (let j = 1; j < formEndIdx; j++) {
                const payStatus = String(allData[j][N + 2] || '').trim();
                const isPaid    = payStatus === '済' || payStatus === '繰越' || payStatus === 'くりこし';
                if (!isPaid) continue;
                const nameRaw  = String(allData[j][2] || '').trim();
                if (!nameRaw) continue;
                const gradeStr = String(allData[j][4] || '').trim();
                const fee      = calcFeeFromGrade_(gradeStr, feeMap);
                if (fee <= 0) continue;
                appendSuitouTx_(ss, normalizeName_(nameRaw), -fee, name + '　参加費');
              }
            }
          }
        }

        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: '大会が見つかりません' });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会を完了済みにする（カレンダーシートの M 列に「完了」を書く）
function completeTournament(name) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 出納管理の収支チェック（当該大会の参加費合計が±ゼロでなければエラー）
    const suitouCheckSheet = ss.getSheetByName('出納管理');
    if (suitouCheckSheet && suitouCheckSheet.getLastRow() >= 7) {
      const txRows = suitouCheckSheet.getRange(7, 1, suitouCheckSheet.getLastRow() - 6, 3).getValues();
      const personNet = {};
      for (const row of txRows) {
        if (String(row[2]) !== name + '　参加費') continue;
        const n = String(row[0]).trim();
        if (!n) continue;
        personNet[n] = (personNet[n] || 0) + (typeof row[1] === 'number' ? row[1] : Number(row[1]) || 0);
      }
      const unbalanced = Object.entries(personNet)
        .filter(([, net]) => net !== 0)
        .map(([n]) => n);
      if (unbalanced.length > 0) {
        return JSON.stringify({
          error: '以下の参加者の収支が合っていません。デポジット処理を行ってから完了してください。\n' + unbalanced.join('、'),
        });
      }
    }

    // 大会登録済みチェック
    const tournamentSheet = ss.getSheetByName(name);
    if (tournamentSheet) {
      const colA = tournamentSheet.getRange(1, 1, tournamentSheet.getLastRow(), 1).getValues();
      let registered = false;
      for (let i = 0; i < colA.length - 1; i++) {
        if (String(colA[i][0]).trim() === 'registerDatabase') {
          registered = String(colA[i + 1][0]).includes('登録済み');
          break;
        }
      }
      if (!registered) {
        return JSON.stringify({ error: '「大会として登録」が完了していないため、完了にできません' });
      }
    }

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

    // メール管理シートから送信済み行を削除
    const mailSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (mailSheet && mailSheet.getLastRow() >= 6) {
      const mailData = mailSheet.getRange(6, 1, mailSheet.getLastRow() - 5, 9).getValues();
      const deleteRows = [];
      for (let i = 0; i < mailData.length; i++) {
        if (String(mailData[i][0]) + String(mailData[i][1]) === name && String(mailData[i][7]) === '済') {
          deleteRows.push(i + 6); // 実際の行番号
        }
      }
      for (let i = deleteRows.length - 1; i >= 0; i--) {
        mailSheet.deleteRow(deleteRows[i]);
      }
    }

    // 出納管理シートから当該大会のトランザクションを削除
    const suitouSheet = ss.getSheetByName('出納管理');
    if (suitouSheet && suitouSheet.getLastRow() >= 7) {
      const txData = suitouSheet.getRange(7, 1, suitouSheet.getLastRow() - 6, 3).getValues();
      const txDeleteRows = [];
      for (let i = 0; i < txData.length; i++) {
        if (String(txData[i][2]).startsWith(name)) {
          txDeleteRows.push(i + 7);
        }
      }
      for (let i = txDeleteRows.length - 1; i >= 0; i--) {
        suitouSheet.deleteRow(txDeleteRows[i]);
      }
    }

    // 大会シートを非表示
    if (tournamentSheet) tournamentSheet.hideSheet();

    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
