// ============================================================
// 大会フォーム回答シートの管理設定 v2（A:C縦持ち）
// ============================================================

const TOURNAMENT_SHEET_V2_MARKER_ = '__TAIKAI_MANAGEMENT_V2__';
const TOURNAMENT_SHEET_V2_END_ = '__TAIKAI_MANAGEMENT_V2_END__';
const TOURNAMENT_SHEET_V2_VERSION_ = 2;
const TOURNAMENT_SHEET_V2_WIDTH_ = 3;
const TOURNAMENT_SHEET_V2_SCHEDULE_FIELDS_ = [
  '級', '参加費', '開催日',
];

function tournamentSheetV2Cell_(value) {
  return value === undefined || value === null ? '' : value;
}

function tournamentSheetV2Rows_(snapshot) {
  const rows = [];
  const row = (type, key, value) => {
    rows.push([type, key, tournamentSheetV2Cell_(value)]);
  };
  const schedules = snapshot.schedules || [];
  const firstSchedule = schedules[0] || {};
  const settings = snapshot.settings || {};

  row(TOURNAMENT_SHEET_V2_MARKER_, TOURNAMENT_SHEET_V2_VERSION_, '');
  row('[大会]', '項目', '値');
  row('大会', 'フォームID', snapshot.form_id);
  row('大会', 'フォーム編集URL', snapshot.form_edit_url);
  row('大会', '公認', snapshot.is_sanctioned);
  [
    ['申込開始日', settings['申込開始日'] || snapshot.registration_starts_at],
    ['リマインダー', settings['リマインダー'] || snapshot.reminder_at],
    ['本申込期限', settings['本申込期限'] || firstSchedule.application_deadline],
    ['抽選日', settings['抽選日'] || firstSchedule.lottery_result_date],
    ['本振込期限', settings['本振込期限'] || firstSchedule.payment_deadline],
    ['大会の日時', settings['大会の日時'] || firstSchedule.held_on],
    ['メモ', settings['メモ']],
    ['後納制', settings['後納制'] || firstSchedule.payment_timing],
    ['振込先', settings['振込先'] || firstSchedule.payment_instructions],
  ].forEach(item => row('大会', item[0], item[1]));
  row('[日程]', '項目', '値');
  schedules.forEach(schedule => {
    const identifier = '日程:' + String(schedule.grade || '');
    row(identifier, '級', schedule.grade);
    row(identifier, '参加費', schedule.participation_fee_yen);
    row(identifier, '開催日', schedule.held_on);
  });
  row(TOURNAMENT_SHEET_V2_END_, TOURNAMENT_SHEET_V2_VERSION_, '');
  return rows;
}

function tournamentSheetV2CollectRecord_(records, id, field, value, rowNumber) {
  if (!records[id]) {
    records[id] = { values: {}, field_rows: {}, row_number: rowNumber };
  }
  if (!field || Object.prototype.hasOwnProperty.call(records[id].values, field)) {
    throw new Error('大会管理データv2の項目が空または重複しています: ' + id);
  }
  records[id].values[field] = value;
  records[id].field_rows[field] = rowNumber;
}

function tournamentSheetV2Parse_(data) {
  const markerRows = [];
  for (let index = 1; index < (data || []).length; index++) {
    if (String((data[index] || [])[0] || '') === TOURNAMENT_SHEET_V2_MARKER_) {
      markerRows.push(index);
    }
  }
  if (!markerRows.length) return null;
  if (markerRows.length !== 1) {
    throw new Error('大会管理データv2の開始マーカーが重複しています。');
  }
  const start = markerRows[0];
  if (Number((data[start] || [])[1]) !== TOURNAMENT_SHEET_V2_VERSION_) {
    throw new Error('大会管理データv2のバージョンが不正です。');
  }
  const endRows = [];
  for (let index = start + 1; index < data.length; index++) {
    if (String((data[index] || [])[0] || '') === TOURNAMENT_SHEET_V2_END_) {
      endRows.push(index);
    }
  }
  if (endRows.length !== 1
      || Number((data[endRows[0]] || [])[1]) !== TOURNAMENT_SHEET_V2_VERSION_) {
    throw new Error('大会管理データv2の終了マーカーがありません、重複、または不正です。');
  }
  const end = endRows[0];
  const metadata = {};
  const metadataRows = {};
  const schedules = {};

  for (let index = start + 1; index < end; index++) {
    const current = data[index] || [];
    const type = String(current[0] || '');
    if (type === '大会') {
      const key = String(current[1] || '');
      if (!key || Object.prototype.hasOwnProperty.call(metadata, key)) {
        throw new Error('大会管理データv2の大会項目が空または重複しています。');
      }
      metadata[key] = current[2];
      metadataRows[key] = index + 1;
    } else if (/^日程:[A-E]$/.test(type)) {
      tournamentSheetV2CollectRecord_(
        schedules, type.slice(3), String(current[1] || ''),
        current[2], index + 1
      );
    }
  }
  Object.keys(schedules).forEach(grade => {
    const item = schedules[grade];
    const logical = TOURNAMENT_SHEET_V2_SCHEDULE_FIELDS_.map(field =>
      Object.prototype.hasOwnProperty.call(item.values, field)
        ? item.values[field] : ''
    );
    if (String(logical[0] || '') !== grade) {
      throw new Error('大会管理データv2の日程級が不正です: ' + grade);
    }
    schedules[grade] = {
      row_number: item.row_number,
      row: logical,
      field_rows: item.field_rows,
    };
  });
  ['フォームID', 'フォーム編集URL', '公認'].forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(metadata, key)) {
      throw new Error('大会管理データv2の大会項目が不足しています: ' + key);
    }
  });
  return {
    version: TOURNAMENT_SHEET_V2_VERSION_,
    start_index: start,
    end_index: end,
    metadata: metadata,
    metadata_rows: metadataRows,
    schedules: schedules,
  };
}

function tournamentSheetV2ValidateSnapshot_(snapshot) {
  if (!snapshot || !snapshot.form_id || !snapshot.form_edit_url) {
    throw new Error('大会管理データv2のフォーム識別情報が不足しています。');
  }
  const grades = {};
  (snapshot.schedules || []).forEach(schedule => {
    const grade = String(schedule.grade || '');
    if (!/^[A-E]$/.test(grade) || grades[grade]) {
      throw new Error('大会管理データv2の日程級が不正または重複しています。');
    }
    grades[grade] = true;
  });
  if (!Object.keys(grades).length) {
    throw new Error('大会管理データv2に日程がありません。');
  }
  return true;
}
