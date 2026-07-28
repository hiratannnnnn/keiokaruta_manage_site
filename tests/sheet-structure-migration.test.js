const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'server/SheetStructureMigration.js'), 'utf8'
);
const page = fs.readFileSync(
  path.join(root, 'pages/sheet-migration.html'), 'utf8'
);
const script = fs.readFileSync(
  path.join(root, 'scripts/sheet-migration.html'), 'utf8'
);
const sandbox = { Date };
vm.runInNewContext(source, sandbox);

assert.strictEqual(sandbox.sheetMigrationColumnLabel_(1), 'A');
assert.strictEqual(sandbox.sheetMigrationColumnLabel_(26), 'Z');
assert.strictEqual(sandbox.sheetMigrationColumnLabel_(27), 'AA');
assert.strictEqual(sandbox.sheetMigrationColumnLabel_(703), 'AAA');
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.sheetMigrationCanonicalValue_(2500))),
  { type: 'number', value: '2500' }
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.sheetMigrationCanonicalValue_(true))),
  { type: 'boolean', value: 'true' }
);

const sourceRange = {
  getValues: () => [[123, ''], ['記帳完了', false]],
  getFormulas: () => [['=1+1', ''], ['', '']],
  getNotes: () => [['確認用メモ', ''], ['', '']],
  getNumberFormats: () => [['0', '@'], ['@', '@']],
  getBackgrounds: () => [['#ffff00', '#ffffff'], ['#ffffff', '#ffffff']],
  getDataValidations: () => [[null, null], [null, null]],
  getDisplayValues: () => [['count', '旧操作'], ['記帳完了', 'FALSE']],
};
const records = sandbox.sheetMigrationDeletedCellRecords_({
  getLastRow: () => 2,
  getRange: () => sourceRange,
}, {
  delete_start_column: 4,
  delete_column_count: 2,
});
assert.deepStrictEqual(JSON.parse(JSON.stringify(records)), [
  [
    'D1', '"count"', 'number', '"123"', '"=1+1"', '"確認用メモ"',
    '"0"', '"#ffff00"', '',
  ],
  [
    'E1', '"旧操作"', 'string', '""', '""', '""',
    '"@"', '"#ffffff"', '',
  ],
  [
    'D2', '"count"', 'string', '"記帳完了"', '""', '""',
    '"@"', '"#ffffff"', '',
  ],
  [
    'E2', '"旧操作"', 'boolean', '"false"', '""', '""',
    '"@"', '"#ffffff"', '',
  ],
]);

let writtenRows = null;
const archiveRange = {
  setNumberFormat: () => archiveRange,
  setValues(rows) {
    writtenRows = rows;
    return archiveRange;
  },
  setBackground: () => archiveRange,
  getValues: () => writtenRows,
};
const archive = sandbox.sheetMigrationWriteAndVerifyArchive_({
  getLastRow: () => 10,
  getRange(row, column, rowCount, columnCount) {
    assert.deepStrictEqual([row, column, rowCount, columnCount], [12, 1, 7, 9]);
    return archiveRange;
  },
  getName: () => 'テスト大会A級',
}, 'migration-1', records);
assert.deepStrictEqual(JSON.parse(JSON.stringify(archive)), {
  start_row: 12,
  row_count: 7,
});
assert.strictEqual(writtenRows[0][0], '__SHEET_MIGRATION_ARCHIVE_START__');
assert.strictEqual(writtenRows[6][0], '__SHEET_MIGRATION_ARCHIVE_END__');

const archiveIndex = source.indexOf('sheetMigrationWriteAndVerifyArchive_(');
const deleteIndex = source.indexOf(
  'sheet.deleteColumns(plan.delete_start_column, plan.delete_column_count)'
);
assert.ok(archiveIndex >= 0 && deleteIndex > archiveIndex);
assert.match(source, /getFormulas\(\)/);
assert.match(source, /getNotes\(\)/);
assert.match(source, /getDataValidations\(\)/);
assert.match(source, /dry-run後に削除対象の情報が変更されました/);
assert.match(source, /列削除後の構造検証に失敗しました/);
assert.match(source, /sheetMigrationArchiveRange_\(sheet, migrationId\)/);
assert.match(source, /sheet\.insertColumnsAfter\(plan\.edit_url_column/);
assert.match(source, /target\.deleteRows\(archive\.start_row, archive\.row_count\)/);
assert.match(page, /全情報を[\s\S]*左側へ移送して再読取検証/);
assert.match(script, /移送・再読取検証後/);
assert.match(script, /左側へ移送.*セル/);

console.log('Sheet structure migration safety checks passed.');
