// ============================================================
// 既存大会シートの旧管理列移行
// ============================================================

const SHEET_MIGRATION_HISTORY_SHEET_ = 'シート構造移行履歴';
const SHEET_MIGRATION_BACKUP_PREFIX_ = '__sheet_migration_backup_';

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
    const values = sheet.getRange(
      1, deleteStart, sheet.getLastRow(), deleteCount
    ).getDisplayValues();
    result.non_empty_cell_count = values.reduce((total, row) =>
      total + row.filter(value => String(value || '').trim() !== '').length
    , 0);
    if (result.non_empty_cell_count) {
      result.warnings.push(
        '削除候補列に値が' + result.non_empty_cell_count + 'セルあります。'
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
    sheet.deleteColumns(plan.delete_start_column, plan.delete_column_count);
    history.getRange(historyRow, 7).setValue('success');
    return {
      sheet_name: sheetName,
      migration_id: migrationId,
      status: 'success',
      backup_sheet_name: backupName,
      deleted_column_count: plan.delete_column_count,
      restorable: true,
    };
  } catch (e) {
    if (history.getLastRow() < historyRow) {
      history.getRange(historyRow, 1, 1, 10).setValues([[
        migrationId, sheetName, backup ? backup.getName() : '',
        plan.edit_url_column, plan.delete_start_column, plan.delete_column_count,
        'failed', new Date(), '', e.message,
      ]]);
    } else {
      history.getRange(historyRow, 7).setValue('failed');
      history.getRange(historyRow, 10).setValue(e.message);
    }
    return { sheet_name: sheetName, status: 'failed', error: e.message };
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
