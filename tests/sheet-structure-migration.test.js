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
const formSubmitSource = fs.readFileSync(
  path.join(root, 'server/FormSubmit.js'), 'utf8'
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
vm.runInNewContext(formSubmitSource, sandbox);

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
const scheduleHeader = rows.find(row => row[0] === '[日程]');
assert.strictEqual(scheduleHeader[13], '同期状態');
assert.strictEqual(scheduleHeader[14], '最終同期日時');

const rawStructure = {
  data: [
    ['送信日時', '参加級', '連絡用メールアドレス', '氏名'],
    [new Date('2026-07-01T00:00:00Z'), 'A級', 'FIRST@example.com', '山田 太郎'],
    [new Date('2026-07-02T00:00:00Z'), 'B級', 'first@example.com', '山田 太郎'],
  ],
  response_end_index: 3,
  layout: { raw_response_column_count: 4 },
  version: 2,
};
const responseColumns = sandbox.tournamentSheetResponseColumns_(rawStructure);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(responseColumns)),
  { timestamp: 0, email: 2, name: 3, grade: 1 }
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.sheetMigrationLatestSourceRowsByEmail_(
    rawStructure, responseColumns.email
  ))),
  { 'first@example.com': 3 }
);

const parsed = sandbox.tournamentSheetV2Parse_(
  [['送信日時', 'メールアドレス']].concat(rows), 1
);
assert.strictEqual(parsed.metadata['tournament ID'], '10');
const duplicateMarkers = [['送信日時', 'メールアドレス']]
  .concat(rows)
  .concat([['__TAIKAI_MANAGEMENT_V2__', 2]]);
assert.throws(
  () => sandbox.tournamentSheetV2Parse_(duplicateMarkers, 1),
  /開始マーカーが重複/
);

const writes = [];
const responseSheet = {
  insertRowBefore(rowNumber) {
    writes.push({ method: 'insert', rowNumber });
  },
  getRange(rowNumber, column, rowCount, columnCount) {
    return {
      clearContent() {
        writes.push({ method: 'clear', rowNumber, column, rowCount, columnCount });
        return this;
      },
      setValue(value) {
        writes.push({ method: 'value', rowNumber, column, value });
        return this;
      },
      setValues(values) {
        writes.push({ method: 'values', rowNumber, column, values });
        return this;
      },
    };
  },
};
const formStructure = {
  version: 2,
  data: [
    ['送信日時', '参加級', '連絡用メールアドレス', '氏名'],
    [new Date('2026-07-01T00:00:00Z'), 'A級', 'same@example.com', '山田 太郎'],
    [new Date('2026-07-02T00:00:00Z'), 'A級', 'same@example.com', '山田 太郎'],
    [],
    ['__TAIKAI_MANAGEMENT_V2__', 2],
    ['A', 2500],
    ['申込', 2],
    ['[案内]'],
    ['__TAIKAI_MANAGEMENT_V2_END__', 2],
  ],
  response_end_index: 3,
  grade_rows: { A: 6 },
  layout: { raw_response_column_count: 4 },
  management: {
    start_index: 4,
    end_index: 8,
    entries_by_source_row: { 2: { row_number: 7, row: ['申込', 2] } },
    metadata_rows: { '同期状態': 10, '同期エラー': 11, '最終同期日時': 12 },
  },
};
sandbox.recordFormResponseInTournamentSheetV2_(
  responseSheet,
  formStructure,
  3,
  { email: 'same@example.com', name: '山田 太郎', grade: 'A' },
  {
    player: { id: '30' },
    entry: { id: '40', schedule_id: '20' },
    payment_summary: {
      participation_fee_yen: 2500,
      paid_yen: 1000,
      balance_yen: 1500,
      status: 'partial',
    },
  },
  null
);
assert.ok(writes.some(item =>
  item.method === 'value' && item.rowNumber === 7
  && item.column === 15 && item.value === 'superseded'
));
const insertedEntry = writes.find(item => item.method === 'values');
assert.strictEqual(insertedEntry.values[0][11], 1000);
assert.strictEqual(insertedEntry.values[0][12], 1500);
assert.strictEqual(insertedEntry.values[0][13], 'partial');

assert.match(migrationSource, /admin\/tournament-sheet-snapshot/);
assert.match(migrationSource, /APIが大会シート移行用の全件スナップショット/);
assert.match(migrationSource, /sheetMigrationSnapshot_\(sheet, structure, legacyRecords\)/);
assert.match(migrationSource, /snapshot_signature/);
assert.match(migrationSource, /全セル再読取検証/);
assert.match(migrationSource, /開催日または参加費がAPIと一致しません/);
assert.match(migrationSource, /sheetMigrationRestoreFromBackup_/);
assert.match(migrationSource, /sheetMigrationRestoreProtections_/);
assert.match(migrationSource, /sheetMigrationVerifyBackupRestore_/);
assert.match(migrationSource, /tournamentSheetV2Rows_\(snapshot\)/);
assert.match(migrationSource, /大会管理データv2の再読取検証/);
assert.match(migrationSource, /sheet\.deleteColumns\(plan\.delete_start_column/);
assert.doesNotMatch(migrationSource, /sheetMigrationWriteAndVerifyArchive_/);
assert.match(page, /全情報を[\s\S]*大会管理データv2として記録/);
assert.match(script, /正規化・再読取検証後/);

console.log('Sheet structure migration v2 checks passed.');
