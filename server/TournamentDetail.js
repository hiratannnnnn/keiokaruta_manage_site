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
    const selectionStatusIndex = rawColumnCount;
    const paymentStatusIndex = rawColumnCount + 1;
    const personHeaders = headerRow.slice(0, rawColumnCount)
      .concat(structure.version === 2
        ? ['選考状態', '支払状態']
        : ['選考状態（旧セル兼用）', '支払状態（旧セル推定）']);
    const responseRecords = tournamentSheetResponseRecords_(structure, false);
    const personRecords = responseRecords.map(record => ({
      source_row: record.source_row,
      entry_id: record.entry_id,
      player_id: record.player_id,
      name: record.name,
      selection_status: record.selection_status,
      payment_status: record.payment_status,
      paid_yen: record.paid_yen,
      balance_yen: record.balance_yen,
      payment_display_status: tournamentSheetPaymentDisplayStatus_(record),
      values: record.raw_values.map(formatCell).concat([
        record.selection_status,
        tournamentSheetPaymentDisplayStatus_(record),
      ]),
    }));
    const personRows = personRecords.map(record => record.values);
    const formEndIdx = structure.response_end_index;

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
      const count = responseRecords.filter(record =>
        record.grade.split('').includes(gradeKey)
      ).length;
      gradeSummary.push({ grade: gradeKey, fee, count, total: fee * count, date });
    }

    return JSON.stringify({
      name,
      sheetVersion: structure.version,
      selectionStatusIndex,
      paymentStatusIndex,
      isOfficial,
      isRegistered,
      personHeaders,
      personRows,
      personRecords,
      nameColumnIndex: responseColumns.name,
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
    const declaredGrades = tournamentSheetDeclaredGrades_(name);
    if (structure.version === 2) {
      const newIsOfficial = !tournamentSheetIsSanctioned_(structure);
      const missingGrades = declaredGrades.filter(
        grade => !structure.management.schedules[grade]
      );
      if (missingGrades.length) {
        throw new Error(
          name + ': 日程管理行がありません: ' + missingGrades.join(',')
        );
      }
      const metadataRow = structure.management.metadata_rows['公認'];
      sheet.getRange(metadataRow, 3).setValue(newIsOfficial);
      declaredGrades.forEach(grade => {
        sheet.getRange(structure.management.schedules[grade].row_number, 12)
          .setValue(newIsOfficial);
        sheet.getRange(structure.management.schedules[grade].row_number, 14)
          .setValue('pending_api');
      });
      sheet.getRange(structure.management.metadata_rows['同期状態'], 3)
        .setValue('pending_api');
      try {
        taikaiSetTournamentSanctioned_(
          tournamentSheetBaseName_(name), newIsOfficial, declaredGrades
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
      declaredGrades.forEach(grade => {
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
    taikaiSetTournamentSanctioned_(
      tournamentSheetBaseName_(name), newIsOfficial, declaredGrades
    );
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

// 回答行・entry IDで対象を確定し、選考状態またはAPI支払い履歴を更新する。
function setDetailPayStatus(sheetName, sourceRow, entryId, value, useDeposit) {
  try {
    const isPaid = value === '済' || value === '繰越' || value === 'くりこし';
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);

    const structure = tournamentSheetStructure_(sheet, false);
    const record = tournamentSheetResponseRecord_(structure, sourceRow, entryId);
    if (isPaid && structure.version !== 2) {
      throw new Error(
        '旧大会シートでは選考状態と支払状態を安全に分離できません。'
        + '先に大会シートv2へ移行してから支払いを登録してください。'
      );
    }
    if (isPaid && structure.version === 2) {
      if (!record.entry_id) {
        throw new Error(
          '申込IDがまだ同期されていません。先に完全同期を実行してください。'
        );
      }
      const management = structure.management.entries_by_source_row[record.source_row];
      sheet.getRange(management.row_number, 15).setValue('pending_api');
      try {
        taikaiRecordFullPaymentByEntry_(record.entry_id, useDeposit === true);
        refreshTournamentSheetV2FromApi_(sheet);
      } catch (paymentError) {
        markTournamentSheetV2SyncState_(
          sheet, 'pending_api', paymentError.message || paymentError, record.entry_id
        );
        throw paymentError;
      }
      return JSON.stringify({ ok: true });
    }
    tournamentSheetSelectionStatusRange_(
      sheet, structure, record.source_row
    ).setValue(isPaid && structure.version === 2 ? '' : value);
    if (structure.version === 2) {
      const management = structure.management.entries_by_source_row[record.source_row];
      sheet.getRange(management.row_number, 15).setValue('pending_api');
      sheet.getRange(structure.management.metadata_rows['同期状態'], 3)
        .setValue('pending_api');
    }
    return JSON.stringify({ ok: true });
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
    const records = tournamentSheetResponseRecords_(structure, false);
    const allData = structure.data;
    const formEndIdx = structure.response_end_index;
    const feeMap     = getSuitouFeeMap_(allData, formEndIdx);

    let count = 0, total = 0;
    const gradeMap = {}; // grade -> { fee, names[] }
    records.forEach(record => {
      if (!record.is_paid) return;
      const gradeStr = record.grade;
      const fee      = calcFeeFromGrade_(gradeStr, feeMap);
      if (fee <= 0) return;
      const playerName = record.name;
      count++;
      total += fee;
      if (!gradeMap[gradeStr]) gradeMap[gradeStr] = { fee, names: [] };
      gradeMap[gradeStr].names.push(playerName);
    });
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
