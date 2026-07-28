// ============================================================
// 大会シート管理データ v2
// ============================================================

const TOURNAMENT_SHEET_V2_MARKER_ = '__TAIKAI_MANAGEMENT_V2__';
const TOURNAMENT_SHEET_V2_END_ = '__TAIKAI_MANAGEMENT_V2_END__';
const TOURNAMENT_SHEET_V2_VERSION_ = 2;
const TOURNAMENT_SHEET_V2_WIDTH_ = 17;

const TOURNAMENT_SHEET_V2_SCHEDULE_HEADERS_ = [
  '級', '参加費', '開催日', 'schedule ID', '申込期限', '会内振込期限',
  '本振込期限', '支払時期', '抽選結果日', '会場', '受付締切', '公認',
  '振込先', '同期状態', '最終同期日時',
];

const TOURNAMENT_SHEET_V2_ENTRY_HEADERS_ = [
  '申込', '回答行', 'メール', '氏名', '級', 'シート状態', 'player ID',
  'entry ID', 'schedule ID', 'キャンセル日時', '請求額', '支払済額',
  '残額', '支払状態', '同期状態', '最終同期日時',
  '同期エラー',
];

function tournamentSheetV2Cell_(value) {
  return value === undefined || value === null ? '' : value;
}

function tournamentSheetV2Rows_(snapshot) {
  const rows = [];
  const width = TOURNAMENT_SHEET_V2_WIDTH_;
  const row = values => {
    const result = new Array(width).fill('');
    (values || []).slice(0, width).forEach((value, index) => {
      result[index] = tournamentSheetV2Cell_(value);
    });
    rows.push(result);
  };

  row([TOURNAMENT_SHEET_V2_MARKER_, TOURNAMENT_SHEET_V2_VERSION_]);
  row(['[大会]', '項目', '値']);
  [
    ['大会名', snapshot.tournament_name],
    ['tournament ID', snapshot.tournament_id],
    ['フォームID', snapshot.form_id],
    ['フォーム公開URL', snapshot.form_public_url],
    ['フォーム編集URL', snapshot.form_edit_url],
    ['申込処理完了', snapshot.registration_completed],
    ['大会振込完了', snapshot.payment_completed],
    ['公認', snapshot.is_sanctioned],
    ['同期状態', snapshot.sync_status],
    ['最終同期日時', snapshot.synced_at],
    ['同期エラー', snapshot.sync_error],
  ].forEach(item => row(['大会', item[0], item[1]]));

  row(['[日程]'].concat(TOURNAMENT_SHEET_V2_SCHEDULE_HEADERS_.slice(1)));
  (snapshot.schedules || []).forEach(schedule => row([
    schedule.grade,
    schedule.participation_fee_yen,
    schedule.held_on,
    schedule.id,
    schedule.application_deadline,
    schedule.internal_payment_deadline,
    schedule.payment_deadline,
    schedule.payment_timing,
    schedule.lottery_result_date,
    schedule.venue,
    schedule.reception_ends_at,
    schedule.is_sanctioned,
    schedule.payment_instructions,
    schedule.sync_status,
    schedule.synced_at,
  ]));

  row(['[申込]'].concat(TOURNAMENT_SHEET_V2_ENTRY_HEADERS_.slice(1)));
  (snapshot.entries || []).forEach(entry => row([
    '申込',
    entry.source_row,
    entry.email,
    entry.name,
    entry.grade,
    entry.sheet_status,
    entry.player_id,
    entry.entry_id,
    entry.schedule_id,
    entry.canceled_at,
    entry.participation_fee_yen,
    entry.paid_yen,
    entry.balance_yen,
    entry.payment_status,
    entry.sync_status,
    entry.synced_at,
    entry.sync_error,
  ]));

  row([
    '[案内]', 'announcement ID', '件名', 'フォームURL', 'Gmail message ID',
    '対象schedule IDs',
  ]);
  (snapshot.announcements || []).forEach(item => row([
    '案内', item.id, item.subject, item.form_url, item.gmail_message_id,
    (item.schedule_ids || []).join(','),
  ]));

  row([
    '[メール]', 'email job ID', '種別', '送信予定日時', '処理開始日時',
    '送信完了日時', 'announcement ID', '対象schedule IDs', '配信件数',
    '配信成功件数', '配信失敗件数',
  ]);
  (snapshot.email_jobs || []).forEach(item => {
    const deliveries = item.deliveries || [];
    row([
      'メール', item.id, item.mail_type, item.scheduled_at,
      item.processing_started_at, item.sent_at, item.announcement_id,
      (item.schedule_ids || []).join(','),
      deliveries.length,
      deliveries.filter(delivery => delivery.sent_at).length,
      deliveries.filter(delivery => delivery.last_error).length,
    ]);
  });
  row([
    '[配信]', 'delivery ID', 'email job ID', 'entry ID', 'Gmail message ID',
    'Gmail draft ID', '下書き作成日時', '処理開始日時', '最終試行日時',
    '試行回数', '最終エラー', '送信日時',
  ]);
  (snapshot.email_jobs || []).forEach(item => {
    (item.deliveries || []).forEach(delivery => row([
      '配信', delivery.id, item.id, delivery.entry_id,
      delivery.gmail_message_id, delivery.gmail_draft_id,
      delivery.draft_created_at, delivery.processing_started_at,
      delivery.last_attempted_at, delivery.attempt_count,
      delivery.last_error, delivery.sent_at,
    ]));
  });

  row([
    '[旧管理監査]', '元セル', '元列見出し', '値の型', '値', '数式',
    'メモ', '数値表示形式', '背景色', '入力規則',
  ]);
  (snapshot.legacy_records || []).forEach(record => row([
    '旧管理', record[0], record[1], record[2], record[3], record[4],
    record[5], record[6], record[7], record[8],
  ]));
  row([TOURNAMENT_SHEET_V2_END_, TOURNAMENT_SHEET_V2_VERSION_]);
  return rows;
}

function tournamentSheetV2Parse_(data, responseEndIndex) {
  const markerRows = [];
  for (let index = responseEndIndex; index < (data || []).length; index++) {
    if (String((data[index] || [])[0] || '') === TOURNAMENT_SHEET_V2_MARKER_) {
      markerRows.push(index);
    }
  }
  if (markerRows.length === 0) return null;
  if (markerRows.length !== 1) {
    throw new Error('大会管理データv2の開始マーカーが重複しています。');
  }
  const start = markerRows[0];
  if (Number((data[start] || [])[1]) !== TOURNAMENT_SHEET_V2_VERSION_) {
    throw new Error('大会管理データv2のバージョンが不正です。');
  }
  let end = -1;
  let endCount = 0;
  for (let index = start + 1; index < data.length; index++) {
    if (String((data[index] || [])[0] || '') === TOURNAMENT_SHEET_V2_END_) {
      endCount++;
      if (end < 0) end = index;
    }
  }
  if (end < 0) throw new Error('大会管理データv2の終了マーカーがありません。');
  if (endCount !== 1
      || Number((data[end] || [])[1]) !== TOURNAMENT_SHEET_V2_VERSION_) {
    throw new Error('大会管理データv2の終了マーカーが重複または不正です。');
  }

  const metadata = {};
  const metadataRows = {};
  const schedules = {};
  const entriesBySourceRow = {};
  const announcements = [];
  const emailJobs = [];
  const deliveries = [];
  const legacyRecords = [];
  for (let index = start + 1; index < end; index++) {
    const row = data[index] || [];
    const type = String(row[0] || '');
    if (type === '大会') {
      const key = String(row[1] || '');
      if (!key || Object.prototype.hasOwnProperty.call(metadata, key)) {
        throw new Error('大会管理データv2の大会項目が空または重複しています。');
      }
      metadata[key] = row[2];
      metadataRows[key] = index + 1;
    }
    if (/^[A-E]$/.test(type)) {
      const grade = type;
      if (!/^[A-E]$/.test(grade) || schedules[grade]) {
        throw new Error('大会管理データv2の日程級が不正または重複しています。');
      }
      schedules[grade] = { row_number: index + 1, row: row };
    }
    if (type === '申込') {
      const sourceRow = Number(row[1]);
      if (!Number.isInteger(sourceRow) || sourceRow < 2 || entriesBySourceRow[sourceRow]) {
        throw new Error('大会管理データv2の申込回答行が不正または重複しています。');
      }
      entriesBySourceRow[sourceRow] = { row_number: index + 1, row: row };
    }
    if (type === '案内') announcements.push({ row_number: index + 1, row: row });
    if (type === 'メール') emailJobs.push({ row_number: index + 1, row: row });
    if (type === '配信') deliveries.push({ row_number: index + 1, row: row });
    if (type === '旧管理') {
      legacyRecords.push(row.slice(1, 10));
    }
  }
  [
    '大会名', 'tournament ID', 'フォームID', 'フォーム公開URL',
    'フォーム編集URL', '同期状態', '最終同期日時', '同期エラー',
  ].forEach(key => {
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
    entries_by_source_row: entriesBySourceRow,
    announcements: announcements,
    email_jobs: emailJobs,
    deliveries: deliveries,
    legacy_records: legacyRecords,
  };
}

function tournamentSheetV2ValidateSnapshot_(snapshot, requireSynced) {
  if (!snapshot || !snapshot.form_id
      || !snapshot.form_edit_url || !snapshot.form_public_url) {
    throw new Error('大会管理データv2の大会・フォーム識別情報が不足しています。');
  }
  if (requireSynced === true && !snapshot.tournament_id) {
    throw new Error('大会管理データv2のtournament IDがありません。');
  }
  const scheduleKeys = {};
  const scheduleIds = {};
  (snapshot.schedules || []).forEach(schedule => {
    if (!/^[A-E]$/.test(String(schedule.grade || ''))
        || (requireSynced === true && (!schedule.id || !schedule.held_on))) {
      throw new Error('大会管理データv2の日程情報が不足しています。');
    }
    if (scheduleKeys[schedule.grade]) {
      throw new Error('大会管理データv2の日程級が重複しています。');
    }
    scheduleKeys[schedule.grade] = true;
    if (schedule.id) {
      const id = String(schedule.id);
      if (scheduleIds[id]) {
        throw new Error('大会管理データv2のschedule IDが重複しています。');
      }
      scheduleIds[id] = true;
    }
  });
  if (!Object.keys(scheduleKeys).length) {
    throw new Error('大会管理データv2に日程がありません。');
  }
  const sourceRows = {};
  const entryIds = {};
  const allowedSyncStatuses = [
    'synced', 'pending_api', 'pending_sheet', 'not_selected', 'superseded',
  ];
  (snapshot.entries || []).forEach(entry => {
    if (!Number.isInteger(entry.source_row) || entry.source_row < 2
        || sourceRows[entry.source_row] || !entry.email || !entry.name || !entry.grade) {
      throw new Error('大会管理データv2の回答行対応が不正です。');
    }
    if (entry.sync_status === 'synced'
        && (!entry.player_id || !entry.entry_id || !entry.schedule_id)) {
      throw new Error('同期済み申込のDB識別情報が不足しています。');
    }
    if (!allowedSyncStatuses.includes(String(entry.sync_status || ''))) {
      throw new Error('大会管理データv2の申込同期状態が不正です。');
    }
    if (entry.schedule_id && !scheduleIds[String(entry.schedule_id)]) {
      throw new Error('大会管理データv2の申込schedule IDが日程に存在しません。');
    }
    if (entry.entry_id) {
      const id = String(entry.entry_id);
      if (entryIds[id]) {
        throw new Error('大会管理データv2のentry IDが重複しています。');
      }
      entryIds[id] = true;
    }
    sourceRows[entry.source_row] = true;
  });
  const announcementIds = {};
  (snapshot.announcements || []).forEach(item => {
    if (!item.id || announcementIds[String(item.id)]) {
      throw new Error('大会管理データv2のannouncement IDが不足または重複しています。');
    }
    announcementIds[String(item.id)] = true;
    (item.schedule_ids || []).forEach(id => {
      if (!scheduleIds[String(id)]) {
        throw new Error('大会管理データv2の案内対象日程が存在しません。');
      }
    });
  });
  const emailJobIds = {};
  const deliveryIds = {};
  (snapshot.email_jobs || []).forEach(item => {
    if (!item.id || emailJobIds[String(item.id)]) {
      throw new Error('大会管理データv2のemail job IDが不足または重複しています。');
    }
    emailJobIds[String(item.id)] = true;
    (item.schedule_ids || []).forEach(id => {
      if (!scheduleIds[String(id)]) {
        throw new Error('大会管理データv2のメール対象日程が存在しません。');
      }
    });
    (item.deliveries || []).forEach(delivery => {
      if (!delivery.id || deliveryIds[String(delivery.id)]
          || !delivery.entry_id) {
        throw new Error('大会管理データv2の配信識別情報が不足または重複しています。');
      }
      deliveryIds[String(delivery.id)] = true;
    });
  });
  return true;
}
