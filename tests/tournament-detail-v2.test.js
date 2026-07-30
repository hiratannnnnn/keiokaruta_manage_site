const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const rows = [
  ['タイムスタンプ', 'メールアドレス', '氏名', '級', '振込み済みか'],
  [new Date('2026-07-01'), 'a@example.com', '山田 太郎', 'B級', ''],
  // 空行がなくても、開始マーカーからV2境界を判定できること。
  ['__TAIKAI_MANAGEMENT_V2__', 2, ''],
  ['[大会]', '項目', '値'],
  ['大会', 'フォームID', 'form-1'],
  ['大会', 'フォーム編集URL', 'https://docs.google.com/forms/d/form-1/edit'],
  ['大会', '公認', true],
  ['大会', '申込開始日', '2026-07-01'],
  ['大会', 'リマインダー', '2026-07-31'],
  ['大会', '本申込期限', '2026-08-07'],
  ['大会', '抽選日', '2026-08-14'],
  ['大会', '本振込期限', '2026-08-28'],
  ['大会', '大会の日時', '2026-09-01'],
  ['大会', 'メモ', '管理メモ'],
  ['大会', '後納制', 'before_tournament'],
  ['大会', '振込先', 'テスト銀行'],
  ['[日程]', '項目', '値'],
  ['日程:B', '級', 'B'],
  ['日程:B', '参加費', 2500],
  ['日程:B', '開催日', '2026-09-01'],
  ['__TAIKAI_MANAGEMENT_V2_END__', 2, ''],
];
const actualAceRows = [
  [
    'タイムスタンプ', 'メールアドレス',
    '氏名（名字と名前の間に空白を含めてください。）',
    'ふりがな', '級', '段位',
    '2026年度公認大会出場回数（2026年4月1日～2026年7月29日）',
    '出場大会を全てお書きください。（略称等で構いません）',
    'その他', '振込み済みか',
  ],
  ['', '', '', '', '', '', '', '', '', ''],
  ['__TAIKAI_MANAGEMENT_V2__', 2, ''],
  ['[大会]', '項目', '値'],
  ['大会', 'フォームID', '1PriGP29DjzmKXKIXenm8JGOHsXo_Xr-EI-_Q4Sh_9io'],
  [
    '大会', 'フォーム編集URL',
    'https://docs.google.com/forms/d/1PriGP29DjzmKXKIXenm8JGOHsXo_Xr-EI-_Q4Sh_9io/edit',
  ],
  ['大会', '公認', true],
  ['大会', '申込開始日', '2026/07/29'],
  ['大会', 'リマインダー', '2026/09/04'],
  ['大会', '本申込期限', '2026/09/10'],
  ['大会', '抽選日', '2026/09/25'],
  ['大会', '本振込期限', '2026/10/08'],
  ['大会', '大会の日時', ''],
  ['大会', 'メモ', ''],
  ['大会', '後納制', 'before_tournament'],
  ['大会', '振込先', ''],
  ['[日程]', '項目', '値'],
  ['日程:A', '級', 'A'],
  ['日程:A', '参加費', 2500],
  ['日程:A', '開催日', '2026/10/01'],
  ['日程:C', '級', 'C'],
  ['日程:C', '参加費', 2000],
  ['日程:C', '開催日', '2026/10/02'],
  ['日程:E', '級', 'E'],
  ['日程:E', '参加費', 1500],
  ['日程:E', '開催日', '2026/10/03'],
  ['__TAIKAI_MANAGEMENT_V2_END__', 2, ''],
];
const sheet = {
  getName: () => '第9回大阪なにはえ大会B級',
  getDataRange: () => ({ getValues: () => rows }),
};
const sandbox = {
  Date,
  JSON,
  String,
  Number,
  Boolean,
  Object,
  Array,
  Utilities: {
    formatDate: date => {
      const jst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
      return jst.toISOString().slice(0, 19) + '+09:00';
    },
  },
  CONFIG: { SPREADSHEET_ID: 'spreadsheet-1' },
  SpreadsheetApp: {
    openById: () => ({
      getSheetByName: name =>
        name === '第9回大阪なにはえ大会B級' ? sheet : null,
    }),
  },
  formatCell: value =>
    value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? ''),
  taikaiFindTournament_: () => ({ id: 9 }),
  taikaiApiRequest_: () => ({
    tournament: { registration_completed: 1 },
    entries: [],
  }),
  emailMapRows_: () => [],
  isPseudonymousEmail_: () => false,
  taikaiCompareIds_: (left, right) => Number(left) - Number(right),
};
vm.createContext(sandbox);
vm.runInContext(read('server/TournamentSheetV2.js'), sandbox);
vm.runInContext(read('server/TournamentSheetStructure.js'), sandbox);
vm.runInContext(read('server/TournamentDetail.js'), sandbox);

const detail = JSON.parse(
  sandbox.getTournamentDetail('第9回大阪なにはえ大会B級')
);
assert.strictEqual(detail.error, undefined);
assert.strictEqual(detail.sheetVersion, 2);
assert.strictEqual(detail.personRows.length, 1);
assert.strictEqual(detail.gradeSummary.length, 1);
assert.strictEqual(detail.gradeSummary[0].grade, 'B');
assert.strictEqual(detail.gradeSummary[0].fee, 2500);
assert.ok(detail.bottomRight.some(item =>
  item.key === 'メモ' && item.value === '管理メモ'
));
assert.ok(detail.bottomRight.some(item =>
  item.key === '公認' && item.value === '公認大会'
));
assert.ok(detail.bottomRight.some(item =>
  item.key === '支払時期' && item.value === '大会前'
));
assert.ok(!detail.bottomRight.some(item =>
  item.key === 'フォームID' || item.key === 'フォーム編集URL'
));

// 実際に作成されたACE級シート（回答0件・行2が空行）の配置を固定する。
const actualAceSheet = {
  getName: () => '第21回テストひらたん大会ACE級',
  getDataRange: () => ({ getValues: () => actualAceRows }),
};
sandbox.SpreadsheetApp.openById = () => ({
  getSheetByName: name =>
    name === '第21回テストひらたん大会ACE級' ? actualAceSheet : null,
});
sandbox.taikaiFindTournament_ = () => ({ id: 21 });
const actualAceDetail = JSON.parse(
  sandbox.getTournamentDetail('第21回テストひらたん大会ACE級')
);
assert.strictEqual(actualAceDetail.error, undefined);
assert.strictEqual(actualAceDetail.sheetVersion, 2);
assert.strictEqual(actualAceDetail.personRows.length, 0);
assert.deepStrictEqual(
  actualAceDetail.gradeSummary.map(item => item.grade),
  ['A', 'C', 'E']
);
assert.strictEqual(
  sandbox.tournamentDetailCancellationTiming_(
    '2026-09-24 23:59:59', '2026-09-25'
  ),
  'before'
);
assert.strictEqual(
  sandbox.tournamentDetailCancellationTiming_(
    '2026-09-25 00:00:00', '2026-09-25'
  ),
  'after'
);
assert.strictEqual(
  sandbox.tournamentDetailCountsAsParticipant_({
    cancellation_timing: 'before',
  }),
  false
);
assert.strictEqual(
  sandbox.tournamentDetailCountsAsParticipant_({
    cancellation_timing: 'after',
  }),
  true
);
assert.strictEqual(
  sandbox.tournamentDetailCancellationAt_('before', '2026-09-25'),
  '2026-09-24T12:00:00+09:00'
);

console.log('Tournament detail V2 regression checks passed.');
