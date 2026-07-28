// ============================================================
// 既存大会シートの旧管理列移行
// ============================================================

const SHEET_MIGRATION_HISTORY_SHEET_ = 'シート構造移行履歴';
const SHEET_MIGRATION_BACKUP_PREFIX_ = '__sheet_migration_backup_';

function sheetMigrationColumnLabel_(column) {
  let value = Number(column);
  if (!Number.isInteger(value) || value < 1) {
    throw new Error('列番号が不正です。');
  }
  let label = '';
  while (value > 0) {
    value--;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function sheetMigrationCanonicalValue_(value) {
  if (value instanceof Date) {
    return { type: 'date', value: value.toISOString() };
  }
  if (typeof value === 'number') {
    return { type: 'number', value: String(value) };
  }
  if (typeof value === 'boolean') {
    return { type: 'boolean', value: value ? 'true' : 'false' };
  }
  return { type: 'string', value: String(value === null || value === undefined ? '' : value) };
}

function sheetMigrationValidationDescription_(validation) {
  if (!validation) return '';
  const values = validation.getCriteriaValues().map(value => {
    if (value instanceof Date) return sheetMigrationCanonicalValue_(value);
    if (value && typeof value.getA1Notation === 'function') {
      const sheet = typeof value.getSheet === 'function' ? value.getSheet() : null;
      return {
        type: 'range',
        value: (sheet ? sheet.getName() + '!' : '') + value.getA1Notation(),
      };
    }
    return sheetMigrationCanonicalValue_(value);
  });
  return JSON.stringify({
    criteria_type: String(validation.getCriteriaType()),
    criteria_values: values,
    allow_invalid: validation.getAllowInvalid(),
    help_text: validation.getHelpText() || '',
  });
}

function sheetMigrationLegacyRecords_(sheet, plan) {
  const seen = {};
  const records = [];
  const addRange = (startRow, startColumn, rowCount, columnCount) => {
    if (rowCount < 1 || columnCount < 1) return;
    const range = sheet.getRange(startRow, startColumn, rowCount, columnCount);
    const values = range.getValues();
    const formulas = range.getFormulas();
    const notes = range.getNotes();
    const numberFormats = range.getNumberFormats();
    const backgrounds = range.getBackgrounds();
    const validations = range.getDataValidations();
    const headerValues = sheet.getRange(1, startColumn, 1, columnCount).getDisplayValues()[0];
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < columnCount; c++) {
        const address = sheetMigrationColumnLabel_(startColumn + c) + String(startRow + r);
        if (seen[address]) continue;
        const canonical = sheetMigrationCanonicalValue_(values[r][c]);
        const formula = String(formulas[r][c] || '');
        const note = String(notes[r][c] || '');
        const validation = sheetMigrationValidationDescription_(validations[r][c]);
        const numberFormat = String(numberFormats[r][c] || '');
        const background = String(backgrounds[r][c] || '').toLowerCase();
        const formatted = (numberFormat && numberFormat !== 'General')
          || (background && background !== '#ffffff');
        if (!canonical.value && !formula && !note && !validation && !formatted) continue;
        seen[address] = true;
        records.push([
          address,
          JSON.stringify(String(headerValues[c] || '')),
          canonical.type,
          JSON.stringify(canonical.value),
          JSON.stringify(formula),
          JSON.stringify(note),
          JSON.stringify(numberFormat),
          JSON.stringify(background),
          validation,
        ]);
      }
    }
  };
  addRange(
    1, plan.delete_start_column, Math.max(1, sheet.getLastRow()),
    plan.delete_column_count
  );
  addRange(
    plan.response_end_index + 1, 1,
    Math.max(0, sheet.getLastRow() - plan.response_end_index),
    sheet.getLastColumn()
  );
  return records;
}

function sheetMigrationIsoDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'JST', 'yyyy-MM-dd');
  }
  const match = String(value).trim().match(/^\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : '';
}

function sheetMigrationSnapshot_(sheet, structure, legacyRecords) {
  const baseName = String(sheet.getName()).replace(/[A-E]+級$/, '');
  const tournament = taikaiFindTournament_(baseName);
  const pseudoEmails = [];
  for (let index = 1; index < structure.response_end_index; index++) {
    const email = String((structure.data[index] || [])[1] || '').trim();
    if (email) pseudoEmails.push(pseudonymousEmailFor_(email));
  }
  const api = taikaiApiRequest_(
    'POST',
    '/admin/tournament-sheet-snapshot',
    {
      tournament_id: String(tournament.id),
      pseudonymous_emails: pseudoEmails.filter(
        (value, index, values) => values.indexOf(value) === index
      ),
    },
    null,
    { tournament_name: baseName, operation: '大会シート構造移行' }
  );
  if (!api || api.complete !== true || !api.tournament
      || !Array.isArray(api.schedules) || !Array.isArray(api.entries)
      || !Array.isArray(api.announcements) || !Array.isArray(api.email_jobs)) {
    throw new Error('APIが大会シート移行用の全件スナップショットを返しませんでした。');
  }
  if (String(api.tournament.id) !== String(tournament.id)
      || String(api.tournament.name) !== baseName) {
    throw new Error('大会シートとAPIの大会識別情報が一致しません。');
  }

  const apiSchedules = {};
  const localSanctioned = tournamentSheetIsSanctioned_(structure);
  const legacyPaymentInstructions = structure.version === 1
    ? tournamentSheetPaymentInstructionsFromStructure_(structure)
    : null;
  api.schedules.forEach(schedule => {
    const grade = String(schedule.grade || '').trim().toUpperCase();
    if (!/^[A-E]$/.test(grade) || apiSchedules[grade]) {
      throw new Error('APIの日程を級で一意に対応できません: ' + grade);
    }
    const legacyRow = structure.grade_rows[grade];
    if (!legacyRow) throw new Error('旧シートに' + grade + '級の日程行がありません。');
    const oldFee = tournamentSheetGradeFee_(structure, grade);
    const oldDate = tournamentSheetGradeDate_(structure, grade);
    if (Number(oldFee) !== Number(schedule.participation_fee_yen)
        || sheetMigrationIsoDate_(oldDate) !== String(schedule.held_on || '')) {
      throw new Error(
        grade + '級の開催日または参加費がAPIと一致しません。'
        + ' シート=' + sheetMigrationIsoDate_(oldDate) + '/' + oldFee
        + ' API=' + String(schedule.held_on || '') + '/'
        + String(schedule.participation_fee_yen)
      );
    }
    const apiSanctioned = schedule.is_sanctioned === true
      || String(schedule.is_sanctioned).toLowerCase() === 'true'
      || Number(schedule.is_sanctioned) === 1;
    const localScheduleValue = structure.version === 2
      ? structure.management.schedules[grade].row[11]
      : localSanctioned;
    const localScheduleSanctioned = localScheduleValue === true
      || String(localScheduleValue).toLowerCase() === 'true'
      || Number(localScheduleValue) === 1;
    const localPaymentInstructions = structure.version === 2
      ? String(structure.management.schedules[grade].row[12] || '').trim()
      : legacyPaymentInstructions;
    if (apiSanctioned !== localScheduleSanctioned
        || String(schedule.payment_instructions || '').trim()
          !== String(localPaymentInstructions || '').trim()) {
      throw new Error(
        grade + '級の公認状態または振込先がAPIと一致しません。'
      );
    }
    apiSchedules[grade] = schedule;
  });
  if (!Object.keys(apiSchedules).length) {
    throw new Error('APIに大会日程がありません。');
  }

  const apiEntries = {};
  api.entries.forEach(entry => {
    const key = String(entry.player_email || '').toLowerCase()
      + '|' + String(entry.grade || '').toUpperCase();
    if (apiEntries[key]) {
      throw new Error('API申込を疑似メールと級で一意に対応できません。');
    }
    apiEntries[key] = entry;
  });
  const now = new Date();
  const entries = [];
  for (let index = 1; index < structure.response_end_index; index++) {
    const row = structure.data[index] || [];
    const email = String(row[1] || '').trim();
    const name = String(row[2] || '').trim();
    const grade = String(row[4] || '').replace(/級/g, '').trim().toUpperCase();
    if (!email && !name) continue;
    if (!email || !name || !/^[A-E]$/.test(grade)) {
      throw new Error('回答行' + (index + 1) + 'のメール・氏名・級が不足しています。');
    }
    const key = pseudonymousEmailFor_(email).toLowerCase() + '|' + grade;
    const apiEntry = apiEntries[key];
    const sheetStatus = String(
      tournamentSheetPaymentStatus_(structure, index + 1) || ''
    ).trim();
    const isCarriedOver = (sheetStatus.includes('繰') && sheetStatus.includes('越'))
      || sheetStatus === 'くりこし';
    const requiresApiEntry = sheetStatus === '' || sheetStatus === '済'
      || isCarriedOver;
    if (!apiEntry && requiresApiEntry) {
      throw new Error('回答行' + (index + 1) + 'をAPI申込へ対応できません。');
    }
    entries.push({
      source_row: index + 1,
      email: email,
      name: name,
      grade: grade,
      sheet_status: sheetStatus,
      player_id: apiEntry ? apiEntry.player_id : '',
      entry_id: apiEntry ? apiEntry.entry_id : '',
      schedule_id: apiEntry ? apiEntry.schedule_id : apiSchedules[grade].id,
      canceled_at: apiEntry ? apiEntry.canceled_at : '',
      participation_fee_yen: apiEntry
        ? apiEntry.participation_fee_yen : apiSchedules[grade].participation_fee_yen,
      paid_yen: apiEntry ? apiEntry.paid_yen : 0,
      balance_yen: apiEntry
        ? apiEntry.balance_yen : apiSchedules[grade].participation_fee_yen,
      payment_status: apiEntry ? apiEntry.payment_status : 'not_selected',
      sync_status: apiEntry ? 'synced' : 'not_selected',
      synced_at: apiEntry ? now : '',
      sync_error: '',
    });
  }
  const formId = tournamentSheetFormId_(structure);
  const editUrl = tournamentSheetFormEditUrl_(structure);
  const form = FormApp.openById(formId);
  const snapshot = {
    tournament_name: baseName,
    tournament_id: api.tournament.id,
    form_id: formId,
    form_public_url: form.getPublishedUrl(),
    form_edit_url: editUrl,
    registration_completed: Boolean(Number(api.tournament.registration_completed)),
    payment_completed: Boolean(Number(api.tournament.payment_completed)),
    is_sanctioned: localSanctioned,
    sync_status: 'synced',
    synced_at: now,
    sync_error: '',
    schedules: api.schedules.map(schedule => Object.assign({}, schedule, {
      sync_status: 'synced',
      synced_at: now,
    })),
    entries: entries,
    announcements: api.announcements,
    email_jobs: api.email_jobs,
    legacy_records: legacyRecords,
  };
  tournamentSheetV2ValidateSnapshot_(snapshot, true);
  return snapshot;
}

function sheetMigrationEnabled_() {
  return String(
    PropertiesService.getScriptProperties()
      .getProperty('SHEET_STRUCTURE_MIGRATION_ENABLED') || ''
  ).toLowerCase() === 'true';
}

function tournamentSheetNamesForMigration_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const calendar = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
  if (!calendar || calendar.getLastRow() < 3) return [];
  const values = calendar.getRange(3, 1, calendar.getLastRow() - 2, 1).getValues();
  const seen = {};
  return values.map(row => String(row[0] || '').trim()).filter(name => {
    if (!name || seen[name]) return false;
    seen[name] = true;
    return true;
  });
}

function inspectTournamentSheetMigration_(sheet) {
  const result = {
    sheet_name: sheet.getName(),
    executable: false,
    status: 'blocked',
    warnings: [],
  };
  try {
    const structure = tournamentSheetStructure_(sheet, true);
    if (structure.version === 2) {
      result.status = 'migrated';
      result.reason = '大会管理データv2です。';
      result.response_count = Math.max(0, structure.response_end_index - 1);
      result.tournament_id = structure.management.metadata['tournament ID'];
      return result;
    }
    const layout = structure.layout;
    const deleteStart = layout.payment_status_column;
    const deleteCount = Math.max(0, sheet.getLastColumn() - deleteStart + 1);
    result.edit_url_column = layout.edit_url_column;
    result.form_id_column = layout.form_id_column;
    result.payment_status_column = layout.payment_status_column;
    result.response_count = Math.max(0, structure.response_end_index - 1);
    result.grade_rows = structure.grade_rows;
    result.current_column_count = sheet.getLastColumn();
    result.delete_start_column = deleteStart;
    result.delete_column_count = deleteCount;
    result.delete_column_headers = sheet.getRange(
      1, deleteStart, 1, deleteCount
    ).getDisplayValues()[0].map((value, index) =>
      String(value || '').trim() || '（見出しなし・' + (deleteStart + index) + '列目）'
    );
    result.response_end_index = structure.response_end_index;
    result.non_empty_cell_count = sheetMigrationLegacyRecords_(
      sheet, result
    ).length;
    result.move_target =
      '回答終端下の大会管理データv2（大会・日程・申込/支払・案内・メール・同期・旧管理監査）';
    if (result.non_empty_cell_count) {
      result.warnings.push(
        '旧管理列・旧下部領域の情報' + result.non_empty_cell_count
        + 'セルを監査保存し、業務項目へ正規化してから旧領域を撤去します。'
      );
    }
    result.warnings.push('API全件取得・日程/参加費・疑似メール/級の照合失敗時は実行しません。');
    result.status = 'ready';
    result.executable = true;
    return result;
  } catch (e) {
    result.reason = e.message;
    return result;
  }
}

function sheetMigrationHistorySheet_(createIfMissing) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_MIGRATION_HISTORY_SHEET_);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(SHEET_MIGRATION_HISTORY_SHEET_);
    sheet.getRange(1, 1, 1, 10).setValues([[
      '移行ID', '対象シート', 'バックアップシート', '編集URL列',
      '削除開始列', '削除列数', '状態', '実行日時', '復元日時', 'エラー',
    ]]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function sheetMigrationHistory_() {
  const sheet = sheetMigrationHistorySheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
    .map((row, index) => ({
      row_number: index + 2,
      migration_id: String(row[0] || ''),
      sheet_name: String(row[1] || ''),
      backup_sheet_name: String(row[2] || ''),
      edit_url_column: Number(row[3] || 0),
      delete_start_column: Number(row[4] || 0),
      delete_column_count: Number(row[5] || 0),
      status: String(row[6] || ''),
      executed_at: formatCell(row[7]),
      restored_at: formatCell(row[8]),
      error: String(row[9] || ''),
      restorable: String(row[6] || '') === 'success' && !row[8],
    })).reverse();
}

function previewTournamentSheetMigrations() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const plans = tournamentSheetNamesForMigration_().map(name => {
      const sheet = ss.getSheetByName(name);
      return sheet
        ? inspectTournamentSheetMigration_(sheet)
        : {
          sheet_name: name,
          executable: false,
          status: 'blocked',
          warnings: [],
          reason: 'シートが見つかりません。',
        };
    });
    return JSON.stringify({
      ok: true,
      execution_enabled: sheetMigrationEnabled_(),
      plans: plans,
      history: sheetMigrationHistory_(),
      summary: {
        ready: plans.filter(plan => plan.status === 'ready').length,
        migrated: plans.filter(plan => plan.status === 'migrated').length,
        blocked: plans.filter(plan => plan.status === 'blocked').length,
      },
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function sheetMigrationId_() {
  return Utilities.formatDate(new Date(), 'JST', 'yyyyMMddHHmmss')
    + '-' + Utilities.getUuid().slice(0, 8);
}

function backupNameForMigration_(migrationId) {
  return (SHEET_MIGRATION_BACKUP_PREFIX_ + migrationId).slice(0, 99);
}

function sheetMigrationRestoreFromBackup_(target, backup) {
  if (target.getMaxColumns() < backup.getMaxColumns()) {
    target.insertColumnsAfter(
      target.getMaxColumns(), backup.getMaxColumns() - target.getMaxColumns()
    );
  } else if (target.getMaxColumns() > backup.getMaxColumns()) {
    target.deleteColumns(
      backup.getMaxColumns() + 1, target.getMaxColumns() - backup.getMaxColumns()
    );
  }
  if (target.getMaxRows() < backup.getMaxRows()) {
    target.insertRowsAfter(target.getMaxRows(), backup.getMaxRows() - target.getMaxRows());
  } else if (target.getMaxRows() > backup.getMaxRows()) {
    target.deleteRows(backup.getMaxRows() + 1, target.getMaxRows() - backup.getMaxRows());
  }
  target.clear();
  backup.getRange(
    1, 1, backup.getMaxRows(), backup.getMaxColumns()
  ).copyTo(target.getRange(
    1, 1, backup.getMaxRows(), backup.getMaxColumns()
  ));
}

function executeOneTournamentSheetMigration_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('対象シートが見つかりません。');
  const plan = inspectTournamentSheetMigration_(sheet);
  if (!plan.executable) {
    return { sheet_name: sheetName, status: 'skipped', reason: plan.reason || plan.status };
  }

  const migrationId = sheetMigrationId_();
  const backupName = backupNameForMigration_(migrationId);
  let backup = null;
  const history = sheetMigrationHistorySheet_(true);
  const historyRow = history.getLastRow() + 1;
  try {
    const originalStructure = tournamentSheetStructure_(sheet, true);
    const records = sheetMigrationLegacyRecords_(sheet, plan);
    if (records.length !== plan.non_empty_cell_count) {
      throw new Error('dry-run後に旧管理情報が変更されました。再検査してください。');
    }
    const snapshot = sheetMigrationSnapshot_(sheet, originalStructure, records);
    const v2Rows = tournamentSheetV2Rows_(snapshot);
    backup = sheet.copyTo(ss).setName(backupName);
    backup.hideSheet();
    history.getRange(historyRow, 1, 1, 10).setValues([[
      migrationId, sheetName, backupName, plan.edit_url_column,
      plan.delete_start_column, plan.delete_column_count,
      'backed_up', new Date(), '', '',
    ]]);

    const clearStartRow = originalStructure.response_end_index + 1;
    if (clearStartRow <= sheet.getMaxRows()) {
      sheet.getRange(
        clearStartRow, 1,
        sheet.getMaxRows() - clearStartRow + 1, sheet.getMaxColumns()
      ).clear();
    }
    sheet.deleteColumns(plan.delete_start_column, plan.delete_column_count);
    if (sheet.getMaxColumns() < TOURNAMENT_SHEET_V2_WIDTH_) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(), TOURNAMENT_SHEET_V2_WIDTH_ - sheet.getMaxColumns()
      );
    }
    const startRow = originalStructure.response_end_index + 2;
    const requiredLastRow = startRow + v2Rows.length - 1;
    if (sheet.getMaxRows() < requiredLastRow) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredLastRow - sheet.getMaxRows());
    }
    sheet.getRange(
      startRow, 1, v2Rows.length, TOURNAMENT_SHEET_V2_WIDTH_
    ).setValues(v2Rows);
    sheet.getRange(startRow, 1, 1, TOURNAMENT_SHEET_V2_WIDTH_)
      .setBackground('#d9ead3').setFontWeight('bold');

    const converted = tournamentSheetStructure_(sheet, true);
    if (converted.version !== 2
        || String(converted.management.metadata['tournament ID'])
          !== String(snapshot.tournament_id)
        || Object.keys(converted.management.schedules).length !== snapshot.schedules.length
        || Object.keys(converted.management.entries_by_source_row).length
          !== snapshot.entries.length
        || converted.management.announcements.length !== snapshot.announcements.length
        || converted.management.email_jobs.length !== snapshot.email_jobs.length) {
      throw new Error('大会管理データv2の再読取検証に失敗しました。');
    }
    history.getRange(historyRow, 7).setValue('success');
    return {
      sheet_name: sheetName,
      migration_id: migrationId,
      status: 'success',
      backup_sheet_name: backupName,
      moved_cell_count: records.length,
      management_start_row: startRow,
      deleted_column_count: plan.delete_column_count,
      schedule_count: snapshot.schedules.length,
      entry_count: snapshot.entries.length,
      announcement_count: snapshot.announcements.length,
      email_job_count: snapshot.email_jobs.length,
      restorable: true,
    };
  } catch (e) {
    let rollbackError = '';
    try {
      if (backup) sheetMigrationRestoreFromBackup_(sheet, backup);
    } catch (rollbackException) {
      rollbackError = ' ロールバックにも失敗しました: ' + rollbackException.message;
    }
    const errorMessage = e.message + rollbackError;
    if (history.getLastRow() < historyRow) {
      history.getRange(historyRow, 1, 1, 10).setValues([[
        migrationId, sheetName, backup ? backup.getName() : '',
        plan.edit_url_column, plan.delete_start_column, plan.delete_column_count,
        'failed', new Date(), '', errorMessage,
      ]]);
    } else {
      history.getRange(historyRow, 7).setValue('failed');
      history.getRange(historyRow, 10).setValue(errorMessage);
    }
    return { sheet_name: sheetName, status: 'failed', error: errorMessage };
  }
}

function executeTournamentSheetMigrations(sheetNamesJson) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return JSON.stringify({ error: '別の大会シート移行または同期が実行中です。' });
  }
  try {
    try {
    if (!sheetMigrationEnabled_()) {
      throw new Error(
        '大会シート構造移行はまだ有効化されていません。'
        + ' taikai_manage #54/#55完了後に実行してください。'
      );
    }
    const names = JSON.parse(sheetNamesJson);
    if (!Array.isArray(names) || !names.length) {
      throw new Error('移行対象シートを1件以上選択してください。');
    }
    const allowed = {};
    tournamentSheetNamesForMigration_().forEach(name => { allowed[name] = true; });
    const unique = [];
    names.forEach(name => {
      const clean = String(name || '').trim();
      if (!allowed[clean]) throw new Error('移行対象外のシートです: ' + clean);
      if (!unique.includes(clean)) unique.push(clean);
    });
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    return JSON.stringify({
      ok: true,
      results: unique.map(name => executeOneTournamentSheetMigration_(ss, name)),
    });
    } catch (e) {
      return JSON.stringify({ error: e.message });
    }
  } finally {
    lock.releaseLock();
  }
}

function refreshTournamentSheetV2FromApi_(sheet) {
  const structure = tournamentSheetStructure_(sheet, false);
  if (structure.version !== 2) return { sheet_name: sheet.getName(), skipped: true };
  const snapshot = sheetMigrationSnapshot_(
    sheet, structure, structure.management.legacy_records || []
  );
  const rows = tournamentSheetV2Rows_(snapshot);
  const startRow = structure.management.start_index + 1;
  const oldRowCount = structure.management.end_index
    - structure.management.start_index + 1;
  const oldValues = sheet.getRange(
    startRow, 1, oldRowCount, TOURNAMENT_SHEET_V2_WIDTH_
  ).getValues();
  try {
    const requiredLastRow = startRow + Math.max(rows.length, oldRowCount) - 1;
    if (sheet.getMaxRows() < requiredLastRow) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredLastRow - sheet.getMaxRows());
    }
    sheet.getRange(
      startRow, 1, Math.max(rows.length, oldRowCount), TOURNAMENT_SHEET_V2_WIDTH_
    ).clearContent();
    sheet.getRange(
      startRow, 1, rows.length, TOURNAMENT_SHEET_V2_WIDTH_
    ).setValues(rows);
    const verified = tournamentSheetStructure_(sheet, false);
    if (verified.version !== 2
        || String(verified.management.metadata['tournament ID'])
          !== String(snapshot.tournament_id)
        || Object.keys(verified.management.entries_by_source_row).length
          !== snapshot.entries.length) {
      throw new Error(sheet.getName() + ': API書戻し後の再読取検証に失敗しました。');
    }
  } catch (writeError) {
    sheet.getRange(
      startRow, 1, Math.max(rows.length, oldRowCount), TOURNAMENT_SHEET_V2_WIDTH_
    ).clearContent();
    sheet.getRange(
      startRow, 1, oldRowCount, TOURNAMENT_SHEET_V2_WIDTH_
    ).setValues(oldValues);
    throw writeError;
  }
  return {
    sheet_name: sheet.getName(),
    schedule_count: snapshot.schedules.length,
    entry_count: snapshot.entries.length,
    announcement_count: snapshot.announcements.length,
    email_job_count: snapshot.email_jobs.length,
  };
}

function refreshFiscalYearTournamentSheetsV2_(snapshot) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const results = [];
  (snapshot.tournaments || []).forEach(tournament => {
    (tournament.sheet_names || []).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error('API書戻し対象シートがありません: ' + sheetName);
      results.push(refreshTournamentSheetV2FromApi_(sheet));
    });
  });
  return results;
}

function restoreTournamentSheetMigration(migrationId) {
  try {
    const historySheet = sheetMigrationHistorySheet_(false);
    const history = sheetMigrationHistory_();
    const item = history.find(row => row.migration_id === String(migrationId));
    if (!item || !item.restorable) throw new Error('復元可能な移行履歴が見つかりません。');

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const target = ss.getSheetByName(item.sheet_name);
    const backup = ss.getSheetByName(item.backup_sheet_name);
    if (!target || !backup) throw new Error('対象またはバックアップシートが見つかりません。');
    const plan = inspectTournamentSheetMigration_(target);
    if (plan.status !== 'migrated') {
      throw new Error('対象シートは復元前の列構造ではありません。');
    }
    sheetMigrationRestoreFromBackup_(target, backup);
    const originalRow = historySheet.getRange(
      2, 1, historySheet.getLastRow() - 1, 1
    ).getValues().findIndex(row => String(row[0]) === item.migration_id) + 2;
    historySheet.getRange(originalRow, 7).setValue('restored');
    historySheet.getRange(originalRow, 9).setValue(new Date());
    return JSON.stringify({ ok: true, sheet_name: item.sheet_name, restored: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
