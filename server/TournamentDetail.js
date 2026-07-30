// ============================================================
// 大会詳細・操作パネル
// ============================================================

// 大会詳細取得。V2では回答表とA:C管理ブロックを共通構造パーサーで分離し、
// 大会・日程設定は管理ブロック、申込・支払情報はDBを正本として構成する。
function tournamentDetailSnapshot_(sheetName) {
  const tournament = taikaiFindTournament_(tournamentSheetBaseName_(sheetName));
  const snapshot = taikaiApiRequest_(
    'POST', '/admin/tournament-sheet-snapshot',
    { tournament_id: String(tournament.id) }
  );
  if (!snapshot || !snapshot.tournament || !Array.isArray(snapshot.entries)) {
    throw new Error('大会詳細用のDB応答が不完全です。');
  }
  return snapshot;
}

function tournamentDetailPseudonymMap_() {
  const result = {};
  emailMapRows_().forEach(row => {
    const real = String(row[0] || '').trim().toLowerCase();
    const pseudo = String(row[1] || '').trim().toLowerCase();
    if (real && pseudo) result[real] = pseudo;
  });
  return result;
}

function tournamentDetailEntryForRecord_(snapshot, record, pseudonymMap) {
  const map = pseudonymMap || tournamentDetailPseudonymMap_();
  const realEmail = String(record.email || '').trim().toLowerCase();
  const keyEmail = isPseudonymousEmail_(realEmail)
    ? realEmail : String(map[realEmail] || '');
  if (!keyEmail) return null;
  const matches = (snapshot.entries || []).filter(entry =>
    String(entry.player_email || '').toLowerCase() === keyEmail
    && String(entry.grade || '').toUpperCase() === record.grade
  );
  if (!matches.length) return null;
  const active = matches.filter(entry => !entry.canceled_at);
  const candidates = active.length ? active : matches;
  return candidates.reduce((latest, entry) =>
    !latest || taikaiCompareIds_(entry.entry_id, latest.entry_id) > 0
      ? entry : latest
  , null);
}

function tournamentDetailEnrichRecords_(sheetName, records, snapshot) {
  const currentSnapshot = snapshot || tournamentDetailSnapshot_(sheetName);
  const pseudonymMap = tournamentDetailPseudonymMap_();
  const schedules = {};
  (currentSnapshot.schedules || []).forEach(schedule => {
    schedules[String(schedule.id || '')] = schedule;
  });
  records.forEach(record => {
    const entry = tournamentDetailEntryForRecord_(
      currentSnapshot, record, pseudonymMap
    );
    if (!entry) return;
    record.entry_id = String(entry.entry_id || '');
    record.player_id = String(entry.player_id || '');
    record.schedule_id = String(entry.schedule_id || '');
    record.participation_fee_yen = entry.participation_fee_yen;
    record.paid_yen = entry.paid_yen;
    record.balance_yen = entry.balance_yen;
    record.payment_status = String(entry.payment_status || '');
    record.payment_methods = String(entry.payment_methods || '')
      .split(',').filter(Boolean);
    record.canceled_at = entry.canceled_at || null;
    const schedule = schedules[record.schedule_id] || {};
    record.lottery_result_date = schedule.lottery_result_date || null;
    record.cancellation_timing = tournamentDetailCancellationTiming_(
      record.canceled_at, record.lottery_result_date
    );
    record.is_paid = tournamentSheetPaymentIsPaid_(record);
  });
  return currentSnapshot;
}

function tournamentDetailCancellationTiming_(canceledAt, lotteryResultDate) {
  if (!canceledAt) return '';
  const canceledDate = String(canceledAt).slice(0, 10);
  const lotteryDate = String(lotteryResultDate || '').slice(0, 10);
  if (!lotteryDate) return 'canceled';
  return canceledDate < lotteryDate ? 'before' : 'after';
}

// 公開前キャンセルは出場実績・級別人数に含めない。
// 同日を含む公開後キャンセルは、連絡期限後のため出場扱いとする。
function tournamentDetailCountsAsParticipant_(record) {
  return String(record && record.cancellation_timing || '') !== 'before';
}

function tournamentDetailSelectionDisplay_(record) {
  if (record.canceled_at) return 'キャンセル';
  const value = String(record.selection_status || '').trim();
  if (!value || ['済', '繰越', '繰り越し', 'くりこし'].includes(value)) {
    return '出場可能';
  }
  return value;
}

function tournamentDetailPaymentDisplay_(record) {
  if (record.cancellation_timing === 'before' && Number(record.paid_yen || 0) === 0) {
    return '支払不要';
  }
  const status = tournamentSheetPaymentDisplayStatus_(record);
  if (!record.is_paid) return status;
  const methods = record.payment_methods || [];
  if (methods.includes('deposit')) {
    return methods.length === 1 ? '繰り越し' : status + '（デポジット併用）';
  }
  return status;
}

function tournamentDetailPaymentTimingLabel_(value) {
  const labels = {
    with_application: '申込時',
    before_tournament: '大会前',
    on_tournament_day: '大会当日',
    after_tournament: '大会後',
  };
  const normalized = String(value || '').trim();
  return labels[normalized] || normalized;
}

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
    const rawColumnCount = tournamentSheetRawResponseColumnCount_(structure);
    const selectionStatusIndex = rawColumnCount;
    const paymentStatusIndex = rawColumnCount + 1;
    const personHeaders = headerRow.slice(0, rawColumnCount)
      .concat(structure.version === 2
        ? ['選考状態', '支払状態']
        : ['選考状態（旧セル兼用）', '支払状態（旧セル推定）']);
    const responseRecords = tournamentSheetResponseRecords_(structure, false);
    let tournamentSnapshot = null;
    if (structure.version === 2) {
      tournamentSnapshot = tournamentDetailEnrichRecords_(
        name, responseRecords
      );
    }
    const personRecords = responseRecords.map(record => {
      const selectionDisplay = tournamentDetailSelectionDisplay_(record);
      const paymentDisplay = tournamentDetailPaymentDisplay_(record);
      return {
        source_row: record.source_row,
        entry_id: record.entry_id,
        player_id: record.player_id,
        name: record.name,
        selection_status: record.selection_status,
        selection_display_status: selectionDisplay,
        payment_status: record.payment_status,
        payment_methods: record.payment_methods || [],
        paid_yen: record.paid_yen,
        balance_yen: record.balance_yen,
        canceled_at: record.canceled_at || null,
        lottery_result_date: record.lottery_result_date || null,
        cancellation_timing: record.cancellation_timing || '',
        payment_display_status: paymentDisplay,
        values: record.raw_values.map(formatCell).concat([
          selectionDisplay,
          paymentDisplay,
        ]),
      };
    });
    const personRows = personRecords.map(record => record.values);
    const formEndIdx = structure.response_end_index;

    let bottomLeft;
    let bottomRight;
    if (structure.version === 2) {
      bottomLeft = rows.slice(
        structure.management.start_index,
        structure.management.end_index + 1
      ).map(row => row.slice(0, 3));
      const detailSettings = [
        ['公認', tournamentSheetIsSanctioned_(structure)
          ? '公認大会' : '非公認大会'],
        ['申込開始日', tournamentSheetManagementValue_(structure, '申込開始日')],
        ['リマインダー', tournamentSheetManagementValue_(structure, 'リマインダー')],
        ['本申込期限', tournamentSheetManagementValue_(structure, '本申込期限')],
        ['抽選日', tournamentSheetManagementValue_(structure, '抽選日')],
        ['本振込期限', tournamentSheetManagementValue_(structure, '本振込期限')],
        ['大会の日時', tournamentSheetManagementValue_(structure, '大会の日時')],
        ['メモ', tournamentSheetManagementValue_(structure, 'メモ')],
        ['支払時期', tournamentDetailPaymentTimingLabel_(
          tournamentSheetPaymentTimingFromStructure_(structure)
        )],
        ['振込先', tournamentSheetPaymentInstructionsFromStructure_(structure)],
      ];
      bottomRight = detailSettings
        .filter(item => item[1] !== null
          && item[1] !== undefined && String(item[1]).trim() !== '')
        .map(item => ({ key: item[0], value: formatCell(item[1]) }));
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
    let isRegistered = structure.version === 2 && tournamentSnapshot
      ? Boolean(Number(tournamentSnapshot.tournament.registration_completed))
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
    const gradeKeys = Object.keys(structure.grade_rows);
    gradeKeys.forEach(gradeKey => {
      const fee = tournamentSheetGradeFee_(structure, gradeKey);
      if (typeof fee !== 'number' || fee < 0) return;
      const date = formatCell(tournamentSheetGradeDate_(structure, gradeKey));
      // 公開前キャンセルを除き、公開後キャンセルは出場扱いで数える。
      const count = responseRecords.filter(record =>
        record.grade.split('').includes(gradeKey)
        && tournamentDetailCountsAsParticipant_(record)
      ).length;
      gradeSummary.push({ grade: gradeKey, fee, count, total: fee * count, date });
    });

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
      taikaiSetTournamentSanctioned_(
        tournamentSheetBaseName_(name), newIsOfficial, declaredGrades
      );
      tournamentSheetManagementRange_(sheet, structure, '公認')
        .setValue(newIsOfficial);
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
    if (structure.version === 2) {
      tournamentSheetManagementRange_(sheet, structure, key).setValue(newValue);
      return JSON.stringify({ ok: true });
    }
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
    if (structure.version === 2) {
      return JSON.stringify({
        ok: true,
        value: formatCell(tournamentSheetManagementValue_(structure, key)),
      });
    }
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
    const isPaid = value === '済' || value === '繰越'
      || value === '繰り越し' || value === 'くりこし';
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
      const expectedEntry = tournamentDetailEntryForRecord_(
        tournamentDetailSnapshot_(sheetName), record
      );
      const resolvedEntryId = String(expectedEntry && expectedEntry.entry_id || '');
      if (!resolvedEntryId) {
        throw new Error(
          '申込IDをDBから取得できません。大会詳細を再読み込みしてください。'
        );
      }
      if (String(entryId || '') !== resolvedEntryId) {
        throw new Error(
          '画面表示後に申込情報が変わりました。大会詳細を再読み込みしてください。'
        );
      }
      taikaiRecordFullPaymentByEntry_(resolvedEntryId, useDeposit === true);
      return JSON.stringify({ ok: true });
    }
    tournamentSheetSelectionStatusRange_(
      sheet, structure, record.source_row
    ).setValue(value);
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function tournamentDetailCancellationAt_(timing, lotteryResultDate) {
  const lotteryDate = String(lotteryResultDate || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(lotteryDate)) {
    throw new Error('抽選公開日が未設定のため、キャンセル区分を判定できません。');
  }
  if (timing === 'before') {
    const date = new Date(
      new Date(lotteryDate + 'T12:00:00+09:00').getTime() - 86400000
    );
    return Utilities.formatDate(
      date, 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"
    );
  }
  if (timing === 'after') {
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
    if (today < lotteryDate) {
      throw new Error('抽選公開日前のため、公開後キャンセルにはできません。');
    }
    return Utilities.formatDate(
      new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"
    );
  }
  if (timing === 'clear') return null;
  throw new Error('キャンセル区分が不正です。');
}

function setEntryCancellationStatus(
  sheetName, sourceRow, entryId, timing, movePaymentsToDeposit
) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`「${sheetName}」シートが見つかりません`);
    const structure = tournamentSheetStructure_(sheet, false);
    if (structure.version !== 2) {
      throw new Error('キャンセル管理は大会シートv2でのみ利用できます。');
    }
    const record = tournamentSheetResponseRecord_(
      structure, sourceRow, entryId
    );
    const snapshot = tournamentDetailSnapshot_(sheetName);
    const entry = tournamentDetailEntryForRecord_(snapshot, record);
    const resolvedEntryId = String(entry && entry.entry_id || '');
    if (!resolvedEntryId || resolvedEntryId !== String(entryId || '')) {
      throw new Error(
        '画面表示後に申込情報が変わりました。大会詳細を再読み込みしてください。'
      );
    }
    const schedule = (snapshot.schedules || []).find(item =>
      String(item.id || '') === String(entry.schedule_id || '')
    );
    if (!schedule) throw new Error('対象級の日程をDBから取得できません。');
    const canceledAt = tournamentDetailCancellationAt_(
      timing, schedule.lottery_result_date
    );
    const result = taikaiSetEntryCancellationStatus_(
      resolvedEntryId, canceledAt, movePaymentsToDeposit === true
    );
    return JSON.stringify({ ok: true, result: result });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getEntryPaymentDetail(entryId) {
  try {
    const normalizedEntryId = String(entryId || '');
    if (!/^\d+$/.test(normalizedEntryId)) {
      throw new Error('申込IDが不正です。');
    }
    return JSON.stringify({
      ok: true,
      payment: taikaiApiRequest_(
        'GET',
        '/entries/' + encodeURIComponent(normalizedEntryId)
          + '/payment-summary'
      ),
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function reverseEntryPaymentFromDetail(entryId, paymentId) {
  try {
    const normalizedEntryId = String(entryId || '');
    const normalizedPaymentId = String(paymentId || '');
    if (!/^\d+$/.test(normalizedEntryId)
        || !/^\d+$/.test(normalizedPaymentId)) {
      throw new Error('申込IDまたは支払履歴IDが不正です。');
    }
    const summary = taikaiApiRequest_(
      'GET',
      '/entries/' + encodeURIComponent(normalizedEntryId) + '/payment-summary'
    );
    const payment = (summary.payments || []).find(item =>
      String(item.id || '') === normalizedPaymentId
      && String(item.entry_id || '') === normalizedEntryId
    );
    if (!payment) {
      throw new Error('対象の支払履歴を申込から確認できません。');
    }
    const result = taikaiApiRequest_(
      'POST',
      '/entry-payments/' + encodeURIComponent(normalizedPaymentId) + '/reverse',
      {
        reversed_at: Utilities.formatDate(
          new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"
        ),
      }
    );
    return JSON.stringify({ ok: true, payment: result });
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
    if (structure.version === 2) {
      tournamentDetailEnrichRecords_(name, records);
    }
    const feeMap = {};
    Object.keys(structure.grade_rows).forEach(grade => {
      const fee = tournamentSheetGradeFee_(structure, grade);
      if (typeof fee === 'number' && fee > 0) feeMap[grade] = fee;
    });

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
    tournamentSheetGradeFeeRange_(sheet, structure, grade).setValue(fee);
    if (structure.version === 2) {
      const gradeDates = {};
      tournamentSheetDeclaredGrades_(sheetName).forEach(item => {
        gradeDates[item] = tournamentSheetGradeDate_(structure, item);
      });
      taikaiSyncTournamentSchedulesFromSheet_(sheetName, gradeDates);
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
    // シートはフォーム・表示用に残すが、運用上の大会日程は新DBへ同期する。
    try {
      taikaiSyncTournamentSchedulesFromSheet_(sheetName, gradeDates);
    } catch (syncError) {
      throw syncError;
    }
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// フォーム編集URL取得
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
