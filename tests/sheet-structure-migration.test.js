const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const v2Source = fs.readFileSync(
  path.join(root, 'server/TournamentSheetV2.js'), 'utf8'
);
const structureSource = fs.readFileSync(
  path.join(root, 'server/TournamentSheetStructure.js'), 'utf8'
);
const migrationSource = fs.readFileSync(
  path.join(root, 'server/SheetStructureMigration.js'), 'utf8'
);
const sandbox = { Date };
vm.runInNewContext(v2Source, sandbox);
vm.runInNewContext(structureSource, sandbox);

const snapshot = {
  form_id: 'form-1',
  form_edit_url: 'https://docs.google.com/forms/d/form-1/edit',
  is_sanctioned: true,
  settings: {
    '申込開始日': new Date('2026-02-26T00:00:00Z'),
    'リマインダー': new Date('2026-03-11T00:00:00Z'),
    '本申込期限': '2026-03-18',
    '抽選日': '2026-03-25',
    '本振込期限': '2026-04-22',
    '後納制': '',
    '振込先': '銀行名\n支店名',
  },
  schedules: [{
    grade: 'D',
    participation_fee_yen: 3000,
    held_on: '2026-04-29',
  }],
};

assert.strictEqual(sandbox.tournamentSheetV2ValidateSnapshot_(snapshot), true);
const rows = sandbox.tournamentSheetV2Rows_(snapshot);
assert.ok(rows.every(row => row.length === 3));
assert.ok(rows.some(row =>
  row[0] === '日程:D' && row[1] === '参加費' && row[2] === 3000
));
assert.ok(!rows.some(row =>
  ['申込', '案内', 'メール', '配信'].includes(String(row[0]))
));
assert.ok(!rows.some(row =>
  ['entry ID', 'player ID', 'schedule ID', '同期状態'].includes(String(row[1]))
));

const data = [
  ['タイムスタンプ', 'メールアドレス', '氏名', '級', '振込み済みか', '', '', ''],
  [new Date('2026-03-01'), 'a@example.com', '山田 太郎', 'D級', '済'],
  [],
].concat(rows);
const parsed = sandbox.tournamentSheetV2Parse_(data);
assert.strictEqual(parsed.metadata['フォームID'], 'form-1');
assert.strictEqual(parsed.schedules.D.row[1], 3000);
assert.strictEqual(parsed.schedules.D.row[2], '2026-04-29');
assert.strictEqual(parsed.entries_by_source_row, undefined);

const sheet = {
  getName: () => '鳳玉大会D級',
  getDataRange: () => ({ getValues: () => data }),
};
const structure = sandbox.tournamentSheetStructure_(sheet, true);
assert.strictEqual(structure.layout.raw_response_column_count, 4);
assert.strictEqual(structure.layout.payment_status_column, 5);
assert.strictEqual(
  sandbox.tournamentSheetRawSheetStatus_(structure, 2), '済'
);
assert.strictEqual(sandbox.tournamentSheetGradeFee_(structure, 'D'), 3000);

// 空行がなくても開始マーカーを境界としてV2を判定する。
const contiguousData = [
  ['タイムスタンプ', 'メールアドレス', '氏名', '級', '振込み済みか'],
  [new Date('2026-03-01'), 'a@example.com', '山田 太郎', 'D級', '済'],
].concat(rows);
const contiguousSheet = {
  getName: () => '鳳玉大会D級',
  getDataRange: () => ({ getValues: () => contiguousData }),
};
const contiguousStructure = sandbox.tournamentSheetStructure_(
  contiguousSheet, true
);
assert.strictEqual(contiguousStructure.version, 2);
assert.strictEqual(contiguousStructure.response_end_index, 2);
assert.strictEqual(
  sandbox.tournamentSheetFormEditUrl_(contiguousStructure),
  'https://docs.google.com/forms/d/form-1/edit'
);

assert.match(
  migrationSource,
  /taikai_manage の migrate_tournament_sheet\(sheet_name\)/
);
assert.doesNotMatch(migrationSource, /refreshTournamentSheetV2FromApi_/);

console.log('Tournament sheet v2 regression checks passed.');
