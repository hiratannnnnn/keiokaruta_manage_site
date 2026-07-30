const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const sandbox = { Date };
vm.runInNewContext(read('server/TournamentSheetV2.js'), sandbox);
vm.runInNewContext(read('server/TournamentSheetStructure.js'), sandbox);
vm.runInNewContext(read('server/Calendar.js'), sandbox);

assert.strictEqual(
  sandbox.tournamentSheetBaseName_('第1回大会AB級'), '第1回大会'
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    sandbox.tournamentSheetDeclaredGrades_('第1回大会AB級')
  )),
  ['A', 'B']
);
assert.throws(
  () => sandbox.tournamentSheetDeclaredGrades_('第1回大会AA級'),
  /担当級が重複/
);
assert.throws(
  () => sandbox.tournamentSheetValidateGradeOwnership_([
    '第1回大会AB級', '第1回大会BC級',
  ]),
  /B級が複数フォームに重複/
);
assert.doesNotThrow(
  () => sandbox.tournamentSheetValidateGradeOwnership_([
    '第1回大会AB級', '第1回大会CD級', '第2回大会AB級',
  ])
);

const calendarRows = [
  [],
  [],
  ['第1回大会AB級', '', '', '', '', '', '済', '', '', '', '', '', ''],
  ['第1回大会CD級', '', '', '', '', '', '', '', '', '', '', '済', ''],
  ['別大会A級', '', '', '', '', '', '済', '', '', '', '', '済', ''],
];
const siblings = sandbox.calendarTournamentSiblingRows_(
  calendarRows, '第1回大会AB級'
);
assert.strictEqual(siblings.length, 2);
assert.strictEqual(
  sandbox.calendarAggregatedCompletion_(siblings, '第1回大会CD級', 7, '済'),
  true
);
assert.strictEqual(
  sandbox.calendarAggregatedCompletion_(siblings, '第1回大会AB級', 12, '済'),
  true
);

const latest = sandbox.tournamentLatestByEmail_([
  {
    email: 'same@example.com',
    registered_at_ms: new Date('2026-07-01').getTime(),
    source_order: 0,
    sheet_name: '第1回大会AB級',
  },
  {
    email: 'SAME@example.com',
    registered_at_ms: new Date('2026-07-02').getTime(),
    source_order: 1,
    sheet_name: '第1回大会CD級',
  },
]);
assert.strictEqual(
  latest['same@example.com'].sheet_name,
  '第1回大会CD級'
);

console.log('Tournament multi-sheet regression checks passed.');
