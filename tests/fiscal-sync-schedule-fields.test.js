const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const structureSource = fs.readFileSync(
  path.join(root, 'server/TournamentSheetStructure.js'), 'utf8'
);
const fiscalSource = fs.readFileSync(
  path.join(root, 'server/FiscalYearSync.js'), 'utf8'
);
const taikaiSource = fs.readFileSync(
  path.join(root, 'server/TaikaiApi.js'), 'utf8'
);
const databaseAdminSource = fs.readFileSync(
  path.join(root, 'scripts/database-admin.html'), 'utf8'
);

const sandbox = {
  Date,
  Utilities: {
    formatDate(value) {
      return [
        value.getFullYear(),
        String(value.getMonth() + 1).padStart(2, '0'),
        String(value.getDate()).padStart(2, '0'),
      ].join('-');
    },
  },
};
vm.runInNewContext(structureSource, sandbox);
vm.runInNewContext(fiscalSource, sandbox);

const data = [
  ['タイムスタンプ'],
  ['2026-04-01'],
  ['', '', '↓振込先'],
  ['', '', '銀行名'],
  ['', '', '支店名'],
  ['', '', ''],
  ['', '', '口座番号'],
  ['', '', '名義'],
];
assert.strictEqual(
  sandbox.tournamentSheetPaymentInstructions_(data, 2),
  '銀行名\n支店名\n口座番号\n名義'
);

const blankInstructions = [
  ['header'],
  ['', '', '↓振込先'],
  ['', '', ''],
  ['', '', ''],
  ['', '', ''],
  ['', '', ''],
  ['', '', ''],
];
assert.strictEqual(
  sandbox.tournamentSheetPaymentInstructions_(blankInstructions, 1),
  null
);

assert.throws(
  () => sandbox.tournamentSheetPaymentInstructions_([
    ['header'],
    ['', '', '↓振込先'],
    ['', '', '↓振込先'],
  ], 1),
  /候補2件/
);
assert.throws(
  () => sandbox.tournamentSheetPaymentInstructions_([['header']], 1),
  /候補0件/
);
assert.throws(
  () => sandbox.tournamentSheetPaymentInstructions_([
    ['header'],
    ['', '', '↓振込先'],
    ['', '', 'x'.repeat(10001)],
  ], 1),
  /10000文字/
);

assert.strictEqual(
  sandbox.fiscalSyncCalendarColumn_([
    ['', ''],
    ['大会名', '振込 開始'],
  ], '振込開始'),
  1
);
assert.strictEqual(
  sandbox.fiscalSyncCalendarColumn_([
    ['振込開始'],
    ['振込開始'],
  ], '振込開始'),
  0,
  '同じ列に同じ見出しがあっても一つの列として扱う'
);
assert.strictEqual(
  sandbox.fiscalSyncOptionalDate_('').value,
  null
);
assert.strictEqual(
  sandbox.fiscalSyncOptionalDate_(new Date(2026, 6, 31)).value,
  '2026-07-31'
);
const fiscalStructure = {
  version: 2,
  grade_rows: { A: 1, B: 2, C: 3 },
  management: {
    schedules: {
      A: { row: ['A', 2500, new Date(2026, 3, 1)] },
      B: { row: ['B', 2500, new Date(2027, 2, 31)] },
      C: { row: ['C', 2500, new Date(2027, 3, 1)] },
    },
  },
};
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.fiscalSyncCurrentGradesForSheet_(
    '年度境界大会ABC級', fiscalStructure, 2026
  ))),
  ['A', 'B']
);
assert.strictEqual(
  sandbox.fiscalSyncOptionalDate_('2026/07/31').valid,
  false
);
assert.strictEqual(
  sandbox.tournamentSheetPaymentTimingFromStructure_({
    version: 2,
    management: {
      metadata: { '後納制': '後納', '本振込期限': '' },
    },
  }),
  'after_tournament'
);
assert.strictEqual(
  sandbox.tournamentSheetPaymentTimingFromStructure_({
    version: 2,
    management: {
      metadata: { '後納制': '', '本振込期限': new Date(2026, 6, 31) },
    },
  }),
  'before_tournament'
);

assert.match(fiscalSource, /internal_payment_deadline/);
assert.match(fiscalSource, /payment_instructions/);
assert.match(fiscalSource, /payment_method: isCarriedOver \? 'carried_over' : 'bank_transfer'/);
assert.match(fiscalSource, /sync_payment: structure\.version !== 2/);
assert.match(fiscalSource, /if \(structure\.version === 2\)[\s\S]*schedule\.payment_instructions/);
assert.match(fiscalSource, /else \{[\s\S]*schedule\.venue = null/);
assert.doesNotMatch(fiscalSource, /refreshFiscalYearTournamentSheetsV2_/);
assert.doesNotMatch(fiscalSource, /markFiscalTournamentSheetsV2SyncState_/);
assert.match(taikaiSource, /internal_payment_deadline/);
assert.match(taikaiSource, /payment_instructions/);
assert.match(taikaiSource, /if \(structure\.version !== 2\) \{[\s\S]*body\.venue = null/);
assert.match(databaseAdminSource, /internal_payment_deadline/);
assert.match(databaseAdminSource, /payment_instructions/);

console.log('Fiscal sync schedule field regression checks passed.');
