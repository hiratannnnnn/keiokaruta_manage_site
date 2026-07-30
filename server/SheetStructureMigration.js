/**
 * 大会シート移行の旧バッチ画面との互換窓口。
 *
 * 現行年度の移行は taikai_manage の migrate_tournament_sheet(sheet_name)
 * で一シートずつ行う。こちらからDB同期やシート再構築は行わない。
 */

function sheetMigrationProgressKey_(operationId) {
  return 'sheet_migration_progress_' + String(operationId || 'default');
}

function sheetMigrationSetProgress_(operationId, progress) {
  CacheService.getScriptCache().put(
    sheetMigrationProgressKey_(operationId),
    JSON.stringify(progress),
    21600
  );
}

function getTournamentSheetMigrationProgress(operationId) {
  var value = CacheService.getScriptCache().get(
    sheetMigrationProgressKey_(operationId)
  );
  return value || JSON.stringify({
    status: 'idle',
    phase: '',
    current: 0,
    total: 0,
    message: ''
  });
}

function requestTournamentSheetMigrationCancellation(operationId) {
  sheetMigrationSetProgress_(operationId, {
    status: 'cancelled',
    phase: 'cancelled',
    current: 0,
    total: 0,
    message: '処理を中断しました。'
  });
  return JSON.stringify({ ok: true });
}

function previewTournamentSheetMigrations(operationId) {
  operationId = operationId || Utilities.getUuid();
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var calendarSheet = spreadsheet.getSheetByName(CALENDAR_SHEET_NAME);
  var sheetNames = [];

  if (calendarSheet && calendarSheet.getLastRow() >= 3) {
    calendarSheet.getRange(3, 1, calendarSheet.getLastRow() - 2, 1)
      .getDisplayValues()
      .forEach(function (row) {
        var name = String(row[0] || '').trim();
        if (name && sheetNames.indexOf(name) < 0) {
          sheetNames.push(name);
        }
      });
  }

  sheetMigrationSetProgress_(operationId, {
    status: 'running',
    phase: 'inspect',
    current: 0,
    total: sheetNames.length,
    message: '大会シートの構造を確認しています。'
  });

  var plans = sheetNames.map(function (sheetName, index) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    var plan;

    if (!sheet) {
      plan = {
        sheet_name: sheetName,
        status: 'blocked',
        reason: 'シートが見つかりません。'
      };
    } else {
      try {
        var structure = tournamentSheetStructure_(sheet);
        if (structure.version === 2) {
          plan = {
            sheet_name: sheetName,
            status: 'migrated',
            reason: '新シート構造へ移行済みです。'
          };
        } else {
          plan = {
            sheet_name: sheetName,
            status: 'blocked',
            reason: 'taikai_manage の migrate_tournament_sheet("' +
              sheetName.replace(/"/g, '\\"') + '") を実行してください。'
          };
        }
      } catch (error) {
        plan = {
          sheet_name: sheetName,
          status: 'blocked',
          reason: '構造を判定できません: ' + error.message
        };
      }
    }

    sheetMigrationSetProgress_(operationId, {
      status: 'running',
      phase: 'inspect',
      current: index + 1,
      total: sheetNames.length,
      sheet_name: sheetName,
      message: sheetName + ' を確認しました。'
    });
    return plan;
  });

  var summary = {
    executable: 0,
    migrated: plans.filter(function (plan) {
      return plan.status === 'migrated';
    }).length,
    blocked: plans.filter(function (plan) {
      return plan.status === 'blocked';
    }).length
  };

  sheetMigrationSetProgress_(operationId, {
    status: 'completed',
    phase: 'completed',
    current: sheetNames.length,
    total: sheetNames.length,
    message: '構造確認が完了しました。'
  });

  return JSON.stringify({
    ok: true,
    operation_id: operationId,
    plans: plans,
    history: [],
    summary: summary
  });
}

function executeTournamentSheetMigrations() {
  return JSON.stringify({
    error: 'この画面からの一括移行は廃止しました。' +
      'taikai_manage の migrate_tournament_sheet(sheet_name) を一シートずつ実行してください。'
  });
}
