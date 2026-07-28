const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const sandbox = {
  Date,
  FormApp: {
    openById() {
      return { getPublishedUrl: () => 'https://example.com/public-form' };
    },
  },
};
vm.runInNewContext(read('server/TournamentSheetV2.js'), sandbox);
vm.runInNewContext(read('server/TournamentSheetStructure.js'), sandbox);
vm.runInNewContext(read('server/SheetStructureMigration.js'), sandbox);
vm.runInNewContext(read('server/Calendar.js'), sandbox);
vm.runInNewContext(read('server/TaikaiApi.js'), sandbox);

assert.strictEqual(sandbox.tournamentSheetBaseName_('第1回大会AB級'), '第1回大会');
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
assert.strictEqual(
  sandbox.calendarAggregatedCompletion_(siblings, '第1回大会CD級', 12, ''),
  false
);

const schedule = (id, grade, heldOn) => ({
  id,
  tournament_id: '1',
  grade,
  held_on: heldOn,
  participation_fee_yen: 2500,
  payment_instructions: 'テスト口座',
  is_sanctioned: true,
});
const entry = (id, scheduleId, grade, email) => ({
  entry_id: id,
  player_id: 'p' + id,
  player_email: email,
  schedule_id: scheduleId,
  grade,
  canceled_at: '',
  participation_fee_yen: 2500,
  paid_yen: 0,
  balance_yen: 2500,
  payment_status: 'unpaid',
});
const apiSnapshot = {
  complete: true,
  tournament: {
    id: '1',
    name: '第1回大会',
    registration_completed: true,
    payment_completed: false,
  },
  schedules: [
    schedule('sA', 'A', '2026-08-01'),
    schedule('sC', 'C', '2026-08-02'),
  ],
  entries: [
    entry('eA', 'sA', 'A', 'pseudo-a@example.com'),
    entry('eC', 'sC', 'C', 'pseudo-c@example.com'),
  ],
  announcements: [
    { id: 'aA', schedule_ids: ['sA'] },
    { id: 'aC', schedule_ids: ['sC'] },
  ],
  email_jobs: [{
    id: 'jA',
    announcement_id: 'aA',
    schedule_ids: ['sA'],
    deliveries: [
      { id: 'dA', entry_id: 'eA' },
      { id: 'dC', entry_id: 'eC' },
    ],
  }],
};
let currentApiSnapshot = apiSnapshot;
sandbox.taikaiFindTournament_ = () => ({ id: '1', name: '第1回大会' });
sandbox.taikaiApiRequest_ = () => currentApiSnapshot;
sandbox.pseudonymousEmailFor_ = email => 'pseudo-' + email;
sandbox.taikaiCompareIds_ = (left, right) => String(left).localeCompare(String(right));

const managementEntry = selectionStatus => [
  '申込', 2, 'a@example.com', '山田 太郎', 'A', selectionStatus,
  'peA', 'eA', 'sA', '', 2500, 0, 2500, 'unpaid', 'synced', '', '',
];
const structure = selectionStatus => ({
  version: 2,
  data: [
    ['送信日時', '参加級', 'メールアドレス', '氏名'],
    [new Date('2026-07-01T00:00:00Z'), 'A級', 'a@example.com', '山田 太郎'],
    [], [], [], [], [], [], [],
    ['A', 2500, '2026-08-01', 'sA', '', '', '', '', '', '', '', true,
      'テスト口座'],
  ],
  response_end_index: 2,
  layout: { raw_response_column_count: 4 },
  grade_rows: { A: 10 },
  management: {
    metadata: {
      '大会名': '第1回大会',
      'tournament ID': '1',
      'フォームID': 'form-A',
      'フォーム公開URL': 'https://example.com/public-form',
      'フォーム編集URL': 'https://example.com/edit-form',
      '申込処理完了': true,
      '大会振込完了': false,
      '公認': true,
      '同期状態': 'synced',
      '最終同期日時': '',
      '同期エラー': '',
    },
    schedules: {
      A: {
        row_number: 10,
        row: ['A', 2500, '2026-08-01', 'sA', '', '', '', '', '', '', '', true,
          'テスト口座'],
      },
    },
    entries_by_source_row: {
      2: { row_number: 12, row: managementEntry(selectionStatus) },
    },
  },
});
const siblingStructure = {
  version: 2,
  data: [
    ['送信日時', '参加級', 'メールアドレス', '氏名'],
    [new Date('2026-07-02T00:00:00Z'), 'C級', 'c@example.com', '佐藤 花子'],
    [],
  ],
  response_end_index: 2,
  layout: { raw_response_column_count: 4 },
  management: {
    schedules: {
      C: {
        row_number: 10,
        row: ['C', 2500, '2026-08-02', 'sC'],
      },
    },
    entries_by_source_row: {},
  },
};
const sheet = {
  getName: () => '第1回大会A級',
  structure: structure(''),
};
const siblingSheet = {
  getName: () => '第1回大会C級',
  structure: siblingStructure,
};
const calendarSheet = {
  getLastRow: () => 4,
  getRange: () => ({
    getValues: () => [['第1回大会A級'], ['第1回大会C級']],
  }),
};
sandbox.CONFIG = {
  SPREADSHEET_ID: 'main',
  SHEET_NAMES: { CALENDAR: 'カレンダー' },
};
sandbox.SpreadsheetApp = {
  openById: () => ({
    getSheetByName(name) {
      if (name === 'カレンダー') return calendarSheet;
      if (name === '第1回大会A級') return sheet;
      if (name === '第1回大会C級') return siblingSheet;
      return null;
    },
  }),
};
sandbox.tournamentSheetStructure_ = target => target.structure;
const projected = sandbox.sheetMigrationSnapshot_(sheet, structure(''), []);
assert.strictEqual(projected.schedules.length, 1);
assert.strictEqual(projected.schedules[0].grade, 'A');
assert.strictEqual(projected.entries.length, 1);
assert.strictEqual(projected.entries[0].entry_id, 'eA');
assert.strictEqual(projected.announcements.length, 1);
assert.strictEqual(projected.announcements[0].id, 'aA');
assert.strictEqual(projected.email_jobs[0].deliveries.length, 1);
assert.strictEqual(projected.email_jobs[0].deliveries[0].entry_id, 'eA');
assert.throws(
  () => {
    sheet.structure = structure('キャンセル待ち1番');
    return sandbox.sheetMigrationSnapshot_(sheet, sheet.structure, []);
  },
  /選考対象外ですが、API申込が有効/
);

// 同じ人が兄弟フォームへ新しく回答した場合、旧フォーム行をsupersededにする。
sheet.structure = structure('');
siblingSheet.structure = Object.assign({}, siblingStructure, {
  data: [
    ['送信日時', '参加級', 'メールアドレス', '氏名'],
    [new Date('2026-07-03T00:00:00Z'), 'C級', 'a@example.com', '山田 太郎'],
    [],
  ],
});
currentApiSnapshot = Object.assign({}, apiSnapshot, {
  entries: [
    entry('eC', 'sC', 'C', 'pseudo-a@example.com'),
  ],
});
const correctedProjection = sandbox.sheetMigrationSnapshot_(
  sheet, sheet.structure, []
);
assert.strictEqual(correctedProjection.entries[0].sync_status, 'superseded');
assert.strictEqual(correctedProjection.entries[0].entry_id, '');

const sanctionRequests = [];
sandbox.taikaiFindTournament_ = () => ({ id: '1' });
sandbox.taikaiApiRequest_ = (method, apiPath, body) => {
  if (method === 'GET') {
    return [
      { id: 'sA', grade: 'A' },
      { id: 'sB', grade: 'B' },
      { id: 'sC', grade: 'C' },
    ];
  }
  sanctionRequests.push({ method, apiPath, body });
  return {};
};
assert.strictEqual(
  sandbox.taikaiSetTournamentSanctioned_('第1回大会', false, ['A', 'B']),
  2
);
assert.deepStrictEqual(
  sanctionRequests.map(item => item.apiPath),
  ['/schedules/sA', '/schedules/sB']
);

const formCreateSource = read('server/FormCreate.js');
const deleteSource = read('server/DeleteTournament.js');
assert.match(
  formCreateSource,
  /const grades2 = tournamentSheetDeclaredGrades_\(sheetName\)/
);
assert.match(formCreateSource, /tournamentSheetValidateGradeOwnership_/);
assert.match(deleteSource, /taikaiDeleteTournamentSchedules_/);
assert.match(deleteSource, /先に大会シートv2へ移行してください/);

console.log('Tournament multi-sheet ownership checks passed.');
