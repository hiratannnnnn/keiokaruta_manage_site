// ============================================================
// 大会フォーム回答シートの共通構造判定
// ============================================================

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
  const responseEndIndex = tournamentSheetResponseEndIndex_(data);
  return {
    data: data,
    layout: tournamentSheetLayout_(sheet),
    response_end_index: responseEndIndex,
    grade_rows: tournamentSheetGradeRows_(
      data, responseEndIndex, requireAllGrades === true
    ),
    register_database_row: tournamentSheetUniqueLabelRow_(
      data, responseEndIndex, 'registerDatabase', false
    ),
  };
}
