// ============================================================
// 大会フォーム回答シートの共通構造判定
// ============================================================

function tournamentSheetBaseName_(sheetName) {
  return String(sheetName || '').replace(/[A-E]+級$/, '');
}

function tournamentSheetDeclaredGrades_(sheetName) {
  const match = String(sheetName || '').match(/([A-E]+)級$/);
  if (!match) {
    throw new Error('大会シート名から担当級を特定できません: ' + sheetName);
  }
  const seen = {};
  return match[1].split('').map(grade => {
    if (seen[grade]) {
      throw new Error('大会シート名の担当級が重複しています: ' + sheetName);
    }
    seen[grade] = true;
    return grade;
  });
}

function tournamentSheetValidateGradeOwnership_(sheetNames) {
  const owners = {};
  (sheetNames || []).forEach(sheetName => {
    const name = String(sheetName || '').trim();
    if (!/[A-E]+級$/.test(name)) return;
    const baseName = tournamentSheetBaseName_(name);
    tournamentSheetDeclaredGrades_(name).forEach(grade => {
      const key = baseName + '|' + grade;
      if (owners[key] && owners[key] !== name) {
        throw new Error(
          baseName + grade + '級が複数フォームに重複しています: '
          + owners[key] + ' / ' + name
        );
      }
      owners[key] = name;
    });
  });
  return owners;
}

function tournamentResponseTimestampMs_(value) {
  return value && typeof value.getTime === 'function'
    && !isNaN(value.getTime())
    ? value.getTime() : null;
}

function tournamentLatestByEmail_(items) {
  // 解釈可能な新しい日時を優先する。同時刻・日時不明同士は、大会の
  // カレンダー順→シート行順で割り当てたsource_orderの後勝ちとする。
  const latestByEmail = {};
  (items || []).forEach(item => {
    const key = String(item.email || '').trim().toLowerCase();
    if (!key) return;
    const current = latestByEmail[key];
    const itemTime = item.registered_at_ms === null
      ? -Infinity : Number(item.registered_at_ms);
    const currentTime = !current || current.registered_at_ms === null
      ? -Infinity : Number(current.registered_at_ms);
    if (!current || itemTime > currentTime
        || (itemTime === currentTime
          && Number(item.source_order) > Number(current.source_order))) {
      latestByEmail[key] = item;
    }
  });
  return latestByEmail;
}

function tournamentSheetGoogleFormIdFromEditUrl_(value) {
  if (typeof value !== 'string') return null;
  const match = value.trim().match(
    /^https:\/\/docs\.google\.com\/forms\/(?:u\/\d+\/)?d\/([^/?#]+)\/edit(?:\?.*)?$/
  );
  return match ? match[1] : null;
}

// 旧横構造でも、右側旧操作列を削除した移行後構造でも利用できるアンカー。
// 保存済みcountは使わず、編集URLとその左隣のフォームIDを照合する。
function tournamentSheetLayout_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 2) throw new Error('大会シートの列数が不足しています: ' + sheet.getName());
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const candidates = [];
  for (let index = 1; index < firstRow.length; index++) {
    const formId = tournamentSheetGoogleFormIdFromEditUrl_(firstRow[index]);
    if (formId && String(firstRow[index - 1] || '').trim() === formId) {
      candidates.push(index + 1);
    }
  }
  if (candidates.length !== 1) {
    throw new Error(
      'Googleフォーム編集URL列を一意に特定できません: '
      + sheet.getName() + '（候補' + candidates.length + '件）'
    );
  }
  const editUrlColumn = candidates[0];
  if (editUrlColumn < 6) {
    throw new Error('大会シートのフォーム列配置が不正です: ' + sheet.getName());
  }
  return {
    edit_url_column: editUrlColumn,
    form_id_column: editUrlColumn - 1,
    payment_status_column: editUrlColumn - 2,
    has_legacy_columns: lastColumn > editUrlColumn,
  };
}

// 0-indexed。行1（index 0）はヘッダー、行2以降でA列が最初に空になる位置。
function tournamentSheetResponseEndIndex_(data) {
  let index = 1;
  while (
    index < (data || []).length
    && String(data[index][0] || '').trim() !== ''
  ) {
    index++;
  }
  return index;
}

function tournamentSheetRowsByLabel_(data, responseEndIndex, pattern, description) {
  const rows = {};
  for (let index = responseEndIndex; index < data.length; index++) {
    const label = String(data[index][0] || '').trim();
    if (!pattern.test(label)) continue;
    if (Object.prototype.hasOwnProperty.call(rows, label)) {
      throw new Error(
        '大会シートの' + description + 'が重複しています: ' + label
      );
    }
    rows[label] = index + 1;
  }
  return rows;
}

function tournamentSheetGradeRows_(data, responseEndIndex, requireAll) {
  const rows = tournamentSheetRowsByLabel_(
    data, responseEndIndex, /^[A-E]$/, '級別固定行'
  );
  if (requireAll) {
    const missing = ['A', 'B', 'C', 'D', 'E'].filter(grade => !rows[grade]);
    if (missing.length) {
      throw new Error('大会シートの級別固定行が不足しています: ' + missing.join(','));
    }
  }
  return rows;
}

function tournamentSheetUniqueLabelRow_(data, responseEndIndex, label, required) {
  const matches = [];
  for (let index = responseEndIndex; index < data.length; index++) {
    if (String(data[index][0] || '').trim() === label) matches.push(index + 1);
  }
  if (matches.length > 1 || (required && matches.length !== 1)) {
    throw new Error(
      '大会シートのラベル「' + label + '」を一意に特定できません'
      + '（候補' + matches.length + '件）'
    );
  }
  return matches[0] || null;
}

function tournamentSheetUniqueCellByLabel_(data, responseEndIndex, label) {
  const matches = [];
  for (let rowIndex = responseEndIndex; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex] || [];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
      if (String(row[columnIndex] || '').trim() === label) {
        matches.push({ row_index: rowIndex, column_index: columnIndex });
      }
    }
  }
  if (matches.length !== 1) {
    throw new Error(
      '大会シートのラベル「' + label + '」を一意に特定できません'
      + '（候補' + matches.length + '件）'
    );
  }
  return matches[0];
}

// 「↓振込先」と同じ列の直下5セルを、表示用プレーンテキストとして連結する。
// 空セルは除外するが、セル内の改行は維持する。
function tournamentSheetPaymentInstructions_(data, responseEndIndex) {
  const anchor = tournamentSheetUniqueCellByLabel_(
    data, responseEndIndex, '↓振込先'
  );
  const lines = [];
  for (let offset = 1; offset <= 5; offset++) {
    const row = data[anchor.row_index + offset] || [];
    const text = String(row[anchor.column_index] || '').trim();
    if (text) lines.push(text);
  }
  const result = lines.join('\n').trim();
  if (result.length > 10000) {
    throw new Error('振込先が10000文字を超えています');
  }
  return result || null;
}

function tournamentSheetStructure_(sheet, requireAllGrades) {
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error('大会シートが空です: ' + sheet.getName());
  const legacyResponseEndIndex = tournamentSheetResponseEndIndex_(data);
  const management = tournamentSheetV2Parse_(data);
  if (management) {
    // V2の開始マーカーが回答領域の確定境界。回答と管理ブロックの間の
    // 空行がフォーム追記等で埋まっても、旧レイアウト判定へ戻さない。
    const responseEndIndex = management.start_index;
    const header = data[0] || [];
    const paymentColumns = [];
    header.forEach((value, index) => {
      if (/^振込(?:み)?済みか$/.test(String(value || '').trim())) {
        paymentColumns.push(index + 1);
      }
    });
    if (paymentColumns.length !== 1) {
      throw new Error(
        '大会シートの「振込み済みか」列を一意に特定できません'
        + '（候補' + paymentColumns.length + '件）。'
      );
    }
    const paymentStatusColumn = paymentColumns[0];
    const responseColumnCount = paymentStatusColumn - 1;
    if (responseColumnCount < 2) {
      throw new Error('大会シートのフォーム回答ヘッダーが不足しています。');
    }
    const gradeRows = {};
    Object.keys(management.schedules).forEach(grade => {
      gradeRows[grade] = management.schedules[grade].row_number;
    });
    if (requireAllGrades === true) {
      const nameMatch = String(sheet.getName()).match(/([A-E]+)級$/);
      const requiredGrades = nameMatch ? nameMatch[1].split('') : Object.keys(gradeRows);
      const missing = requiredGrades.filter(grade => !gradeRows[grade]);
      if (missing.length) {
        throw new Error('大会シートの級別固定行が不足しています: ' + missing.join(','));
      }
    }
    return {
      version: 2,
      data: data,
      layout: {
        raw_response_column_count: responseColumnCount,
        payment_status_column: paymentStatusColumn,
        has_legacy_columns: false,
      },
      management: management,
      response_end_index: responseEndIndex,
      grade_rows: gradeRows,
      register_database_row: null,
    };
  }
  return {
    version: 1,
    data: data,
    layout: tournamentSheetLayout_(sheet),
    management: null,
    response_end_index: legacyResponseEndIndex,
    grade_rows: tournamentSheetGradeRows_(
      data, legacyResponseEndIndex, requireAllGrades === true
    ),
    register_database_row: tournamentSheetUniqueLabelRow_(
      data, legacyResponseEndIndex, 'registerDatabase', false
    ),
  };
}

function tournamentSheetFormId_(structure) {
  if (structure.version === 2) {
    return String(structure.management.metadata['フォームID'] || '').trim();
  }
  return String(
    (structure.data[0] || [])[structure.layout.form_id_column - 1] || ''
  ).trim();
}

function tournamentSheetFormEditUrl_(structure) {
  if (structure.version === 2) {
    return String(structure.management.metadata['フォーム編集URL'] || '').trim();
  }
  return String(
    (structure.data[0] || [])[structure.layout.edit_url_column - 1] || ''
  ).trim();
}

function tournamentSheetFormPublicUrl_(structure) {
  const formId = tournamentSheetFormId_(structure);
  return formId ? FormApp.openById(formId).getPublishedUrl() : '';
}

function tournamentSheetRawSheetStatus_(structure, sourceRow) {
  return (structure.data[sourceRow - 1] || [])[
    structure.layout.payment_status_column - 1
  ];
}

function tournamentSheetSelectionStatus_(structure, sourceRow) {
  const value = String(
    tournamentSheetRawSheetStatus_(structure, sourceRow) || ''
  ).trim();
  if (value === '済' || value === '繰越' || value === 'くりこし') {
    return '';
  }
  return value;
}

function tournamentSheetEntryPaymentStatus_(structure, sourceRow) {
  const legacy = String(
    tournamentSheetRawSheetStatus_(structure, sourceRow) || ''
  ).trim();
  return legacy === '済' || legacy === '繰越' || legacy === 'くりこし'
    ? 'paid' : 'unpaid';
}

function tournamentSheetPaymentIsPaid_(record) {
  return record.payment_status === 'paid'
    || record.payment_status === 'overpaid';
}

function tournamentSheetPaymentDisplayStatus_(record) {
  if (record.payment_status === 'overpaid') return '過払い';
  if (record.payment_status === 'paid') return '済';
  if (record.payment_status === 'partial') return '一部入金';
  if (record.payment_status === 'unpriced') return '参加費未設定';
  if (record.payment_status === 'unpaid') return '未払い';
  return '';
}

function tournamentSheetResponseRecords_(structure, includeSuperseded) {
  const columns = tournamentSheetResponseColumns_(structure);
  const width = tournamentSheetRawResponseColumnCount_(structure);
  const latestByEmail = {};
  for (let index = 1; index < structure.response_end_index; index++) {
    const email = String(
      (structure.data[index] || [])[columns.email] || ''
    ).trim().toLowerCase();
    if (email) latestByEmail[email] = index + 1;
  }
  const records = [];
  for (let index = 1; index < structure.response_end_index; index++) {
    const sourceRow = index + 1;
    const raw = structure.data[index] || [];
    const email = String(raw[columns.email] || '').trim();
    const name = String(raw[columns.name] || '').trim();
    if (!email && !name) continue;
    const superseded = Boolean(
      email && latestByEmail[email.toLowerCase()] !== sourceRow
    );
    if (superseded && includeSuperseded !== true) continue;
    const record = {
      source_row: sourceRow,
      raw_values: raw.slice(0, width),
      columns: columns,
      email: email,
      name: name,
      grade: String(raw[columns.grade] || '').replace(/級/g, '').trim().toUpperCase(),
      raw_sheet_status: String(
        tournamentSheetRawSheetStatus_(structure, sourceRow) || ''
      ).trim(),
      selection_status: tournamentSheetSelectionStatus_(structure, sourceRow),
      payment_status: tournamentSheetEntryPaymentStatus_(structure, sourceRow),
      player_id: '',
      entry_id: '',
      schedule_id: '',
      participation_fee_yen: null,
      paid_yen: null,
      balance_yen: null,
      superseded: superseded,
    };
    record.is_paid = tournamentSheetPaymentIsPaid_(record);
    records.push(record);
  }
  return records;
}

function tournamentSheetResponseRecord_(structure, sourceRow, entryId) {
  const matches = tournamentSheetResponseRecords_(structure, false).filter(record =>
    record.source_row === Number(sourceRow)
  );
  if (matches.length !== 1) {
    throw new Error(
      '大会申込を回答行とentry IDで一意に特定できません'
      + '（候補' + matches.length + '件）。'
    );
  }
  return matches[0];
}

function tournamentSheetSelectionStatusRange_(sheet, structure, sourceRow) {
  return sheet.getRange(sourceRow, structure.layout.payment_status_column);
}

function tournamentSheetRawResponseColumnCount_(structure) {
  return structure.version === 2
    ? structure.layout.raw_response_column_count
    : structure.layout.payment_status_column - 1;
}

function tournamentSheetResponseColumn_(structure, field) {
  const header = (structure.data[0] || []).slice(
    0, tournamentSheetRawResponseColumnCount_(structure)
  );
  const predicates = {
    timestamp: value => /タイムスタンプ|回答日時|送信日時/.test(value),
    email: value => /メールアドレス|e-?mail/i.test(value),
    name: value => /氏名|名前/.test(value),
    grade: value => value === '級' || /参加.*級|出場.*級/.test(value),
  };
  if (!predicates[field]) throw new Error('回答項目の指定が不正です: ' + field);
  const matches = [];
  header.forEach((value, index) => {
    if (predicates[field](String(value || '').replace(/\s+/g, '').trim())) {
      matches.push(index);
    }
  });
  if (matches.length !== 1) {
    throw new Error(
      'フォーム回答の「' + field + '」列を一意に特定できません'
      + '（候補' + matches.length + '件）。'
    );
  }
  return matches[0];
}

function tournamentSheetResponseColumns_(structure) {
  return {
    timestamp: tournamentSheetResponseColumn_(structure, 'timestamp'),
    email: tournamentSheetResponseColumn_(structure, 'email'),
    name: tournamentSheetResponseColumn_(structure, 'name'),
    grade: tournamentSheetResponseColumn_(structure, 'grade'),
  };
}

function tournamentSheetGradeFee_(structure, grade) {
  const normalized = String(grade || '').trim();
  if (structure.version === 2) {
    const item = structure.management.schedules[normalized];
    return item ? item.row[1] : null;
  }
  const rowNumber = structure.grade_rows[normalized];
  return rowNumber ? (structure.data[rowNumber - 1] || [])[1] : null;
}

function tournamentSheetGradeDate_(structure, grade) {
  const normalized = String(grade || '').trim();
  if (structure.version === 2) {
    const item = structure.management.schedules[normalized];
    return item ? item.row[2] : null;
  }
  const rowNumber = structure.grade_rows[normalized];
  if (!rowNumber) return null;
  const row = structure.data[rowNumber - 1] || [];
  return row[structure.layout.payment_status_column - 1];
}

function tournamentSheetGradeDateRange_(sheet, structure, grade) {
  const normalized = String(grade || '').trim();
  if (structure.version === 2) {
    const item = structure.management.schedules[normalized];
    const rowNumber = item && item.field_rows['開催日'];
    if (!rowNumber) {
      throw new Error('グレード「' + grade + '」の開催日行が見つかりません。');
    }
    return sheet.getRange(rowNumber, 3);
  }
  const rowNumber = structure.grade_rows[normalized];
  if (!rowNumber) throw new Error('グレード「' + grade + '」の行が見つかりません。');
  return sheet.getRange(rowNumber, structure.layout.payment_status_column);
}

function tournamentSheetIsSanctioned_(structure) {
  if (structure.version === 2) {
    const value = structure.management.metadata['公認'];
    return value === true || String(value).toLowerCase() === 'true'
      || String(value).trim() === '公認' || Number(value) === 1;
  }
  if (!structure.register_database_row) return true;
  const statusRow = structure.data[structure.register_database_row] || [];
  return String(statusRow[1] || '').trim() === '';
}

function tournamentSheetPaymentInstructionsFromStructure_(structure) {
  if (structure.version === 2) {
    return String(structure.management.metadata['振込先'] || '').trim() || null;
  }
  return tournamentSheetPaymentInstructions_(
    structure.data, structure.response_end_index
  );
}

function tournamentSheetPaymentTimingFromStructure_(structure) {
  if (structure.version !== 2) return null;
  const raw = String(
    tournamentSheetManagementValue_(structure, '後納制') || ''
  ).trim();
  const allowed = [
    'with_application',
    'before_tournament',
    'on_tournament_day',
    'after_tournament',
  ];
  if (allowed.includes(raw)) return raw;
  if (raw === '当日支払い') return 'on_tournament_day';
  if (raw) return 'after_tournament';
  return tournamentSheetManagementValue_(structure, '本振込期限')
    ? 'before_tournament' : null;
}

function tournamentSheetManagementValue_(structure, key) {
  if (structure.version !== 2) return null;
  return Object.prototype.hasOwnProperty.call(structure.management.metadata, key)
    ? structure.management.metadata[key] : null;
}

function tournamentSheetManagementRange_(sheet, structure, key) {
  if (structure.version !== 2) {
    throw new Error('大会設定の更新には新シート構造が必要です。');
  }
  const rowNumber = structure.management.metadata_rows[key];
  if (!rowNumber) throw new Error('大会設定「' + key + '」がありません。');
  return sheet.getRange(rowNumber, 3);
}

function tournamentSheetGradeFeeRange_(sheet, structure, grade) {
  const normalized = String(grade || '').trim();
  if (structure.version === 2) {
    const item = structure.management.schedules[normalized];
    const rowNumber = item && item.field_rows['参加費'];
    if (!rowNumber) {
      throw new Error('グレード「' + grade + '」の参加費行が見つかりません。');
    }
    return sheet.getRange(rowNumber, 3);
  }
  const rowNumber = structure.grade_rows[normalized];
  if (!rowNumber) throw new Error('グレード「' + grade + '」の行が見つかりません。');
  return sheet.getRange(rowNumber, 2);
}
