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
    const data = structure.data;
    const rows = data.map(row => row.map(cell => formatCell(cell)));
    if (!rows.length) return JSON.stringify({ name, personHeaders: [], personRows: [], bottomLeft: [], bottomRight: [] });

    const headerRow = rows[0];
    const paymentStatusIndex = structure.layout.payment_status_column - 1;

    const personHeaders = headerRow.slice(0, structure.layout.edit_url_column);
    const personRows = [];
    const formEndIdx = structure.response_end_index;
    for (let i = 1; i < formEndIdx; i++) {
      if (rows[i][2] !== '') {
        personRows.push(rows[i].slice(0, structure.layout.edit_url_column));
      }
    }

    const bottomRows  = rows.slice(formEndIdx).filter(r => r[2] === '');
    const bottomLeft = bottomRows.filter(r => r[0] !== '').map(r =>
      r.slice(0, structure.layout.payment_status_column)
    );
    const rightKeyIndex = structure.layout.form_id_column - 1;
    const bottomRight = bottomRows.filter(r => r[rightKeyIndex] !== '').map(r => ({
      key: r[rightKeyIndex],
      value: r[rightKeyIndex + 1] || '',
    }));

    // 公認/非公認・登録済み判定
    let isOfficial   = true;
    let isRegistered = false;
    if (structure.register_database_row) {
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
      const date = rows[i][structure.layout.payment_status_column - 1] || '';
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
    const allData = structure.data;
    const formEndIdx = structure.response_end_index;
    for (let i = 1; i < formEndIdx; i++) {
      if (String(allData[i][2]) === playerName) {
        const currentValue = String(
          allData[i][structure.layout.payment_status_column - 1] || ''
        ).trim();
        const currentPaid = currentValue === '済'
          || currentValue === '繰越' || currentValue === 'くりこし';
        if (currentPaid && value === '') {
          throw new Error(
            '支払い履歴は削除できません。出納管理から対象履歴を取り消してください。'
          );
        }
        if (isPaid) {
          taikaiRecordFullPaymentByPlayer_(
            String(sheetName).replace(/[A-E]+級$/, ''),
            playerName,
            useDeposit === true
          );
        }
        sheet.getRange(i + 1, structure.layout.payment_status_column).setValue(value);
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
    const allData = structure.data;
    const formEndIdx = structure.response_end_index;
    const feeMap     = getSuitouFeeMap_(allData, formEndIdx);

    let count = 0, total = 0;
    const gradeMap = {}; // grade -> { fee, names[] }
    for (let i = 1; i < formEndIdx; i++) {
      const payStatus = String(
        allData[i][structure.layout.payment_status_column - 1] || ''
      ).trim();
      const isPaid    = payStatus === '済' || payStatus === '繰越' || payStatus === 'くりこし';
      if (!isPaid) continue;
      const gradeStr = String(allData[i][4] || '').trim();
      const fee      = calcFeeFromGrade_(gradeStr, feeMap);
      if (fee <= 0) continue;
      const playerName = String(allData[i][2] || '').trim();
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
        sheet.getRange(
          structure.grade_rows[grade], structure.layout.payment_status_column
        ).setValue(new Date(gradeDates[grade]));
      }
    });

    // シートはフォーム・表示用に残すが、運用上の大会日程は新DBへ同期する。
    taikaiSyncTournamentSchedulesFromSheet_(sheetName, gradeDates);
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
    const layout = tournamentSheetLayout_(sheet);
    const formUrl = String(sheet.getRange(1, layout.edit_url_column).getValue());
    return JSON.stringify({ formUrl });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
