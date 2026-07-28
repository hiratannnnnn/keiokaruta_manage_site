// ============================================================
// 既存大会シートの旧管理列移行
// ============================================================

const SHEET_MIGRATION_HISTORY_SHEET_ = 'シート構造移行履歴';
const SHEET_MIGRATION_BACKUP_PREFIX_ = '__sheet_migration_backup_';
const SHEET_MIGRATION_ARCHIVE_START_ = '__SHEET_MIGRATION_ARCHIVE_START__';
const SHEET_MIGRATION_ARCHIVE_END_ = '__SHEET_MIGRATION_ARCHIVE_END__';

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

// 削除予定領域は、廃止済み制御値を含めて一度すべて左側へ監査保存する。
// 値だけでなく数式・メモ・表示形式・背景色・入力規則も保存する。
function sheetMigrationDeletedCellRecords_(sheet, plan) {
  if (!plan.delete_column_count) return [];
  const rowCount = Math.max(1, sheet.getLastRow());
  const range = sheet.getRange(
    1, plan.delete_start_column, rowCount, plan.delete_column_count
  );
  const values = range.getValues();
  const formulas = range.getFormulas();
  const notes = range.getNotes();
  const numberFormats = range.getNumberFormats();
  const backgrounds = range.getBackgrounds();
  const validations = range.getDataValidations();
  const headers = range.getDisplayValues()[0] || [];
  const records = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = 0; columnIndex < plan.delete_column_count; columnIndex++) {
      const canonical = sheetMigrationCanonicalValue_(values[rowIndex][columnIndex]);
      const formula = String(formulas[rowIndex][columnIndex] || '');
      const note = String(notes[rowIndex][columnIndex] || '');
      const validation = sheetMigrationValidationDescription_(
        validations[rowIndex][columnIndex]
      );
      const numberFormat = String(numberFormats[rowIndex][columnIndex] || '');
      const background = String(backgrounds[rowIndex][columnIndex] || '').toLowerCase();
      const hasNonDefaultFormat =
        (numberFormat && numberFormat !== 'General')
        || (background && background !== '#ffffff');
      if (!canonical.value && !formula && !note && !validation && !hasNonDefaultFormat) {
        continue;
      }
      const column = plan.delete_start_column + columnIndex;
      records.push([
        sheetMigrationColumnLabel_(column) + String(rowIndex + 1),
        JSON.stringify(String(headers[columnIndex] || '')),
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
  return records;
}

function sheetMigrationArchiveRows_(migrationId, records) {
  return [
    [SHEET_MIGRATION_ARCHIVE_START_, migrationId, '', '', '', '', '', '', ''],
    [
      '元セル', '元列見出し', '値の型', '値', '数式', 'メモ',
      '数値表示形式', '背景色', '入力規則',
    ],
  ].concat(records).concat([
    [SHEET_MIGRATION_ARCHIVE_END_, migrationId, '', '', '', '', '', '', ''],
  ]);
}

function sheetMigrationWriteAndVerifyArchive_(sheet, migrationId, records) {
  const rows = sheetMigrationArchiveRows_(migrationId, records);
  const startRow = sheet.getLastRow() + 2;
  const range = sheet.getRange(startRow, 1, rows.length, 9);
  range.setNumberFormat('@');
  range.setValues(rows);
  range.setBackground('#f3f3f3');
  const actual = range.getValues();
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    for (let columnIndex = 0; columnIndex < 9; columnIndex++) {
      if (String(actual[rowIndex][columnIndex] || '')
          !== String(rows[rowIndex][columnIndex] || '')) {
        throw new Error(
          '左側への情報移送後の再読取検証に失敗しました: '
          + sheet.getName() + ' '
          + sheetMigrationColumnLabel_(columnIndex + 1) + (startRow + rowIndex)
        );
      }
    }
  }
  return { start_row: startRow, row_count: rows.length };
}

function sheetMigrationArchiveRange_(sheet, migrationId) {
  const rowCount = Math.max(1, sheet.getLastRow());
  const values = sheet.getRange(1, 1, rowCount, 2).getValues();
  let start = null;
  let end = null;
  values.forEach((row, index) => {
    if (String(row[1] || '') !== String(migrationId)) return;
    if (String(row[0] || '') === SHEET_MIGRATION_ARCHIVE_START_) start = index + 1;
    if (String(row[0] || '') === SHEET_MIGRATION_ARCHIVE_END_) end = index + 1;
  });
  if (!start || !end || end < start) {
    throw new Error('左側に移送した旧管理情報を一意に特定できません。');
  }
  return { start_row: start, row_count: end - start + 1 };
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
    const layout = structure.layout;
    const deleteStart = layout.edit_url_column + 1;
    const deleteCount = Math.max(0, sheet.getLastColumn() - layout.edit_url_column);
    result.edit_url_column = layout.edit_url_column;
    result.form_id_column = layout.form_id_column;
    result.payment_status_column = layout.payment_status_column;
    result.response_count = Math.max(0, structure.response_end_index - 1);
    result.grade_rows = structure.grade_rows;
    result.current_column_count = sheet.getLastColumn();
    result.delete_start_column = deleteStart;
    result.delete_column_count = deleteCount;
    if (deleteCount === 0) {
      result.status = 'migrated';
      result.reason = '編集URLより右側の列はありません。';
      return result;
    }
    result.delete_column_headers = sheet.getRange(
      1, deleteStart, 1, deleteCount
    ).getDisplayValues()[0].map((value, index) =>
      String(value || '').trim() || '（見出しなし・' + (deleteStart + index) + '列目）'
    );
    result.non_empty_cell_count = sheetMigrationDeletedCellRecords_(
      sheet, result
    ).length;
    result.move_target =
      '下部固定領域の左側A:I（値・数式・メモ・表示形式・背景色・入力規則）';
    if (result.non_empty_cell_count) {
      result.warnings.push(
        '削除候補列の情報' + result.non_empty_cell_count
        + 'セルを左側へ移送・検証してから削除します。'
      );
    }
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
  let archive = null;
  let columnsDeleted = false;
  const history = sheetMigrationHistorySheet_(true);
  const historyRow = history.getLastRow() + 1;
  try {
    backup = sheet.copyTo(ss).setName(backupName);
    backup.hideSheet();
    history.getRange(historyRow, 1, 1, 10).setValues([[
      migrationId, sheetName, backupName, plan.edit_url_column,
      plan.delete_start_column, plan.delete_column_count,
      'backed_up', new Date(), '', '',
    ]]);
    const records = sheetMigrationDeletedCellRecords_(sheet, plan);
    if (records.length !== plan.non_empty_cell_count) {
      throw new Error('dry-run後に削除対象の情報が変更されました。再検査してください。');
    }
    archive = sheetMigrationWriteAndVerifyArchive_(sheet, migrationId, records);
    sheet.deleteColumns(plan.delete_start_column, plan.delete_column_count);
    columnsDeleted = true;
    if (sheet.getLastColumn() !== plan.edit_url_column) {
      throw new Error('列削除後の構造検証に失敗しました。バックアップから復元してください。');
    }
    sheetMigrationArchiveRange_(sheet, migrationId);
    history.getRange(historyRow, 7).setValue('success');
    return {
      sheet_name: sheetName,
      migration_id: migrationId,
      status: 'success',
      backup_sheet_name: backupName,
      moved_cell_count: records.length,
      archive_start_row: archive.start_row,
      deleted_column_count: plan.delete_column_count,
      restorable: true,
    };
  } catch (e) {
    let rollbackError = '';
    try {
      if (columnsDeleted && backup) {
        sheet.insertColumnsAfter(plan.edit_url_column, plan.delete_column_count);
        backup.getRange(
          1, plan.delete_start_column, backup.getMaxRows(), plan.delete_column_count
        ).copyTo(sheet.getRange(
          1, plan.delete_start_column, backup.getMaxRows(), plan.delete_column_count
        ));
      }
      if (archive) {
        const currentArchive = sheetMigrationArchiveRange_(sheet, migrationId);
        sheet.deleteRows(currentArchive.start_row, currentArchive.row_count);
      }
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
    const archive = sheetMigrationArchiveRange_(target, item.migration_id);
    const plan = inspectTournamentSheetMigration_(target);
    if (plan.status !== 'migrated') {
      throw new Error('対象シートは復元前の列構造ではありません。');
    }
    target.insertColumnsAfter(item.edit_url_column, item.delete_column_count);
    backup.getRange(
      1, item.delete_start_column, backup.getMaxRows(), item.delete_column_count
    ).copyTo(target.getRange(
      1, item.delete_start_column, backup.getMaxRows(), item.delete_column_count
    ));
    target.deleteRows(archive.start_row, archive.row_count);
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
