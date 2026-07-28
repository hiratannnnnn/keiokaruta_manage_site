// ============================================================
// 大会詳細・操作パネル
// ============================================================

// 大会詳細取得（大会名と同名のシートを読む）
//
// 横方向は、行1のGoogleフォーム編集URLと左隣のフォームIDから特定する。
// 下部固定領域は、A列の連続回答が終わった位置より下だけを検索する。
function getTournamentDetail(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const responseColumns = tournamentSheetResponseColumns_(structure);
    const data = structure.data;
    const rows = data.map(row => row.map(cell => formatCell(cell)));
    if (!rows.length) return JSON.stringify({ name, personHeaders: [], personRows: [], bottomLeft: [], bottomRight: [] });

    const headerRow = rows[0];
    const rawColumnCount = structure.version === 2
      ? structure.layout.raw_response_column_count
      : structure.layout.payment_status_column - 1;
    const paymentStatusIndex = rawColumnCount;
    const personHeaders = headerRow.slice(0, rawColumnCount).concat(['振込み済みか']);
    const personRows = [];
    const formEndIdx = structure.response_end_index;
    for (let i = 1; i < formEndIdx; i++) {
      if (rows[i][responseColumns.name] !== '') {
        personRows.push(
          rows[i].slice(0, rawColumnCount).concat([
            formatCell(tournamentSheetPaymentStatus_(structure, i + 1)),
          ])
        );
      }
    }

    let bottomLeft;
    let bottomRight;
    if (structure.version === 2) {
      bottomLeft = rows.slice(
        structure.management.start_index,
        structure.management.end_index + 1
      ).map(row => row.slice(0, TOURNAMENT_SHEET_V2_WIDTH_));
      bottomRight = Object.keys(structure.management.metadata).map(key => ({
        key: key,
        value: formatCell(structure.management.metadata[key]),
      }));
    } else {
      const bottomRows = rows.slice(formEndIdx).filter(r => r[2] === '');
      bottomLeft = bottomRows.filter(r => r[0] !== '').map(r =>
        r.slice(0, structure.layout.payment_status_column)
      );
      const rightKeyIndex = structure.layout.form_id_column - 1;
      bottomRight = bottomRows.filter(r => r[rightKeyIndex] !== '').map(r => ({
        key: r[rightKeyIndex],
        value: r[rightKeyIndex + 1] || '',
      }));
    }

    // 公認/非公認・登録済み判定
    let isOfficial = tournamentSheetIsSanctioned_(structure);
    let isRegistered = structure.version === 2
      ? Boolean(structure.management.metadata['申込処理完了'])
      : false;
    if (structure.version === 1 && structure.register_database_row) {
      const registerIndex = structure.register_database_row - 1;
      isOfficial = rows[registerIndex + 1]
        ? rows[registerIndex + 1][1] === ''
        : true;
      isRegistered = rows[registerIndex + 1]
        ? rows[registerIndex + 1][0].includes('登録済み')
        : false;
    }

    // グレードサマリー（下部セクションの A〜E 行）
    const gradeSummary = [];
    const gradePattern = /^[A-E]$/;
    for (let i = formEndIdx; i < data.length; i++) {
      const gradeKey = String(data[i][0] || '').trim();
      if (!gradePattern.test(gradeKey)) continue;
      const fee = data[i][1];
      if (typeof fee !== 'number' || fee <= 0) continue;
      const date = formatCell(tournamentSheetGradeDate_(structure, gradeKey));
      // 参加者数カウント（全登録者で grade が一致するもの）
      let count = 0;
      for (let j = 1; j < formEndIdx; j++) {
        if (!data[j][2]) continue;
        const gradeStr = String(data[j][4] || '').trim()
          .replace(/級/g, '')
          .replace(/[Ａ-Ｅ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        if (gradeStr.split('').some(g => g === gradeKey)) count++;
      }
      gradeSummary.push({ grade: gradeKey, fee, count, total: fee * count, date });
    }

    return JSON.stringify({
      name,
      paymentStatusIndex,
      isOfficial,
      isRegistered,
      personHeaders,
      personRows,
      bottomLeft,
      bottomRight,
      gradeSummary,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 公認/非公認をトグルする
function toggleOfficialStatus(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    if (structure.version === 2) {
      const newIsOfficial = !tournamentSheetIsSanctioned_(structure);
      const metadataRow = structure.management.metadata_rows['公認'];
      sheet.getRange(metadataRow, 3).setValue(newIsOfficial);
      Object.keys(structure.management.schedules).forEach(grade => {
        sheet.getRange(structure.management.schedules[grade].row_number, 12)
          .setValue(newIsOfficial);
        sheet.getRange(structure.management.schedules[grade].row_number, 14)
          .setValue('pending_api');
      });
      sheet.getRange(structure.management.metadata_rows['同期状態'], 3)
        .setValue('pending_api');
      try {
        taikaiSetTournamentSanctioned_(
          String(name).replace(/[A-E]+級$/, ''), newIsOfficial
        );
      } catch (syncError) {
        sheet.getRange(structure.management.metadata_rows['同期エラー'], 3)
          .setValue(String(syncError.message || syncError).slice(0, 5000));
        throw syncError;
      }
      sheet.getRange(structure.management.metadata_rows['同期状態'], 3).setValue('synced');
      sheet.getRange(structure.management.metadata_rows['最終同期日時'], 3)
        .setValue(new Date());
      sheet.getRange(structure.management.metadata_rows['同期エラー'], 3).setValue('');
      Object.keys(structure.management.schedules).forEach(grade => {
        sheet.getRange(structure.management.schedules[grade].row_number, 14)
          .setValue('synced');
        sheet.getRange(structure.management.schedules[grade].row_number, 15)
          .setValue(new Date());
      });
      return JSON.stringify({ ok: true, isOfficial: newIsOfficial });
    }
    const row = structure.register_database_row;
    if (!row || row >= structure.data.length) {
      return JSON.stringify({ error: '"registerDatabase" 行または状態行が見つかりません' });
    }
    const isCurrentlyOfficial = String(structure.data[row][1] || '') === '';
    const newIsOfficial = !isCurrentlyOfficial;
    taikaiSetTournamentSanctioned_(String(name).replace(/[A-E]+級$/, ''), newIsOfficial);
    sheet.getRange(row + 1, 2).setValue(isCurrentlyOfficial ? '非公認' : '');
    return JSON.stringify({ ok: true, isOfficial: newIsOfficial });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートの操作トリガー：ラベル行直下の col[0] に値を書き込む
function setTournamentKeyValue(sheetName, key, newValue) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const row = tournamentSheetUniqueLabelRow_(
      structure.data, structure.response_end_index, key, true
    );
    sheet.getRange(row + 1, 1).setValue(newValue);
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートの操作トリガー：ラベル行直下 col[0] の現在値を読み取る
function getTournamentKeyValue(sheetName, key) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const row = tournamentSheetUniqueLabelRow_(
      structure.data, structure.response_end_index, key, true
    );
    return JSON.stringify({
      ok: true,
      value: formatCell((structure.data[row] || [])[0]),
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会詳細シートの振込み済みか列（col N+2）を書き換える
function setDetailPayStatus(sheetName, playerName, value, useDeposit) {
  try {
    const isPaid = value === '済' || value === '繰越' || value === 'くりこし';
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const responseColumns = tournamentSheetResponseColumns_(structure);
    const allData = structure.data;
    const formEndIdx = structure.response_end_index;
    for (let i = 1; i < formEndIdx; i++) {
      if (String(allData[i][responseColumns.name]) === playerName) {
        const currentValue = String(
          tournamentSheetPaymentStatus_(structure, i + 1) || ''
        ).trim();
        const currentPaid = currentValue === '済'
          || currentValue === '繰越' || currentValue === 'くりこし';
        if (currentPaid && value === '') {
          throw new Error(
            '支払い履歴は削除できません。出納管理から対象履歴を取り消してください。'
          );
        }
        if (isPaid) {
          if (structure.version === 2) {
            const entryRow = structure.management.entries_by_source_row[i + 1].row_number;
            tournamentSheetPaymentStatusRange_(sheet, structure, i + 1).setValue(value);
            sheet.getRange(entryRow, 15).setValue('pending_api');
            try {
              taikaiRecordFullPaymentByPlayer_(
                String(sheetName).replace(/[A-E]+級$/, ''),
                playerName,
                useDeposit === true
              );
              refreshTournamentSheetV2FromApi_(sheet);
            } catch (paymentError) {
              const currentStructure = tournamentSheetStructure_(sheet, false);
              const currentEntry = currentStructure.version === 2
                ? currentStructure.management.entries_by_source_row[i + 1]
                : null;
              if (currentEntry) {
                sheet.getRange(currentEntry.row_number, 15).setValue('pending_api');
                sheet.getRange(currentEntry.row_number, 17)
                  .setValue(String(paymentError.message || paymentError).slice(0, 5000));
              }
              throw paymentError;
            }
            return JSON.stringify({ ok: true });
          }
          taikaiRecordFullPaymentByPlayer_(
            String(sheetName).replace(/[A-E]+級$/, ''),
            playerName,
            useDeposit === true
          );
        }
        tournamentSheetPaymentStatusRange_(sheet, structure, i + 1).setValue(value);
        if (structure.version === 2) {
          const entryRow = structure.management.entries_by_source_row[i + 1].row_number;
          sheet.getRange(entryRow, 15).setValue('pending_api');
          sheet.getRange(structure.management.metadata_rows['同期状態'], 3)
            .setValue('pending_api');
        }
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: '選手が見つかりません' });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会への振込み用：済参加者の人数と合計金額を返す
function getTournamentPaySummary(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const responseColumns = tournamentSheetResponseColumns_(structure);
    const allData = structure.data;
    const formEndIdx = structure.response_end_index;
    const feeMap     = getSuitouFeeMap_(allData, formEndIdx);

    let count = 0, total = 0;
    const gradeMap = {}; // grade -> { fee, names[] }
    for (let i = 1; i < formEndIdx; i++) {
      const payStatus = String(
        tournamentSheetPaymentStatus_(structure, i + 1) || ''
      ).trim();
      const isPaid    = payStatus === '済' || payStatus === '繰越' || payStatus === 'くりこし';
      if (!isPaid) continue;
      const gradeStr = String(allData[i][responseColumns.grade] || '').trim();
      const fee      = calcFeeFromGrade_(gradeStr, feeMap);
      if (fee <= 0) continue;
      const playerName = String(allData[i][responseColumns.name] || '').trim();
      count++;
      total += fee;
      if (!gradeMap[gradeStr]) gradeMap[gradeStr] = { fee, names: [] };
      gradeMap[gradeStr].names.push(playerName);
    }
    const grades = Object.keys(gradeMap).sort().map(g => ({ grade: g, fee: gradeMap[g].fee, names: gradeMap[g].names }));
    return JSON.stringify({ ok: true, count, total, grades });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートのグレード行（下部セクション col[0]=A〜E）の参加費（col[1]）を更新する
function setGradeFee(sheetName, grade, fee) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);
    const structure = tournamentSheetStructure_(sheet, false);
    const row = structure.grade_rows[String(grade).trim()];
    if (!row) return JSON.stringify({ error: `グレード「${grade}」の行が見つかりません` });
    sheet.getRange(row, 2).setValue(fee);
    if (structure.version === 2) {
      sheet.getRange(row, 14).setValue('pending_api');
      sheet.getRange(structure.management.metadata_rows['同期状態'], 3)
        .setValue('pending_api');
    }
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 大会シートのグレード行（A〜E）に大会日を書き込む
// gradeDatesJson: JSON文字列 { A: "2026-05-01", B: "2026-05-01", ... }
function saveTournamentDates(sheetName, gradeDatesJson) {
  try {
    const gradeDates = JSON.parse(gradeDatesJson);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, true);
    Object.keys(gradeDates).forEach(grade => {
      if (/^[A-E]$/.test(grade) && gradeDates[grade]) {
        tournamentSheetGradeDateRange_(sheet, structure, grade)
          .setValue(new Date(gradeDates[grade]));
      }
    });
    if (structure.version === 2) {
      sheet.getRange(structure.management.metadata_rows['同期状態'], 3)
        .setValue('pending_api');
    }

    // シートはフォーム・表示用に残すが、運用上の大会日程は新DBへ同期する。
    try {
      taikaiSyncTournamentSchedulesFromSheet_(sheetName, gradeDates);
    } catch (syncError) {
      if (structure.version === 2) {
        sheet.getRange(structure.management.metadata_rows['同期エラー'], 3)
          .setValue(String(syncError.message || syncError).slice(0, 5000));
      }
      throw syncError;
    }
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// フォームURL取得（col N+4: フォームURL）
function getFormUrl(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);
    const structure = tournamentSheetStructure_(sheet, false);
    const formUrl = tournamentSheetFormEditUrl_(structure);
    return JSON.stringify({ formUrl });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
