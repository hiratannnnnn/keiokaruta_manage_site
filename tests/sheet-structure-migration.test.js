const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const migrationSource = fs.readFileSync(
  path.join(root, 'server/SheetStructureMigration.js'), 'utf8'
);
const v2Source = fs.readFileSync(
  path.join(root, 'server/TournamentSheetV2.js'), 'utf8'
);
const structureSource = fs.readFileSync(
  path.join(root, 'server/TournamentSheetStructure.js'), 'utf8'
);
const page = fs.readFileSync(
  path.join(root, 'pages/sheet-migration.html'), 'utf8'
);
const script = fs.readFileSync(
  path.join(root, 'scripts/sheet-migration.html'), 'utf8'
);
const sandbox = { Date };
vm.runInNewContext(v2Source, sandbox);
vm.runInNewContext(structureSource, sandbox);
vm.runInNewContext(migrationSource, sandbox);

assert.strictEqual(sandbox.sheetMigrationColumnLabel_(1), 'A');
assert.strictEqual(sandbox.sheetMigrationColumnLabel_(26), 'Z');
assert.strictEqual(sandbox.sheetMigrationColumnLabel_(27), 'AA');
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.sheetMigrationCanonicalValue_(2500))),
  { type: 'number', value: '2500' }
);

const snapshot = {
  tournament_name: '第1回テスト大会',
  tournament_id: '10',
  form_id: 'form-1',
  form_public_url: 'https://example.com/form',
  form_edit_url: 'https://docs.google.com/forms/d/form-1/edit',
  registration_completed: true,
  payment_completed: false,
  is_sanctioned: true,
  sync_status: 'synced',
  synced_at: new Date('2026-07-28T00:00:00Z'),
  sync_error: '',
  schedules: [{
    grade: 'A',
    participation_fee_yen: 2500,
    held_on: '2026-08-01',
    id: '20',
    application_deadline: '2026-07-01',
    is_sanctioned: true,
    sync_status: 'synced',
  }],
  entries: [{
    source_row: 2,
    email: 'person@example.com',
    name: '山田 太郎',
    grade: 'A',
    sheet_status: '済',
    player_id: '30',
    entry_id: '40',
    schedule_id: '20',
    participation_fee_yen: 2500,
    paid_yen: 2500,
    balance_yen: 0,
    payment_status: 'paid',
    sync_status: 'synced',
  }],
  announcements: [],
  email_jobs: [],
  legacy_records: [['Q1', '"旧列"', 'string', '"記帳済"', '""', '""', '"@"', '"#fff"', '']],
};
assert.strictEqual(sandbox.tournamentSheetV2ValidateSnapshot_(snapshot, true), true);
const rows = sandbox.tournamentSheetV2Rows_(snapshot);
assert.strictEqual(rows[0][0], '__TAIKAI_MANAGEMENT_V2__');
assert.ok(rows.some(row => row[0] === 'A' && row[3] === '20'));
assert.ok(rows.some(row => row[0] === '申込' && row[7] === '40'));
assert.ok(rows.some(row => row[0] === '旧管理' && row[1] === 'Q1'));
assert.strictEqual(rows[0].length, 17);

assert.match(migrationSource, /admin\/tournament-sheet-snapshot/);
assert.match(migrationSource, /APIが大会シート移行用の全件スナップショット/);
assert.match(migrationSource, /開催日または参加費がAPIと一致しません/);
assert.match(migrationSource, /sheetMigrationRestoreFromBackup_/);
assert.match(migrationSource, /tournamentSheetV2Rows_\(snapshot\)/);
assert.match(migrationSource, /大会管理データv2の再読取検証/);
assert.match(migrationSource, /sheet\.deleteColumns\(plan\.delete_start_column/);
assert.doesNotMatch(migrationSource, /sheetMigrationWriteAndVerifyArchive_/);
assert.match(page, /全情報を[\s\S]*大会管理データv2として記録/);
assert.match(script, /正規化・再読取検証後/);

console.log('Sheet structure migration v2 checks passed.');
