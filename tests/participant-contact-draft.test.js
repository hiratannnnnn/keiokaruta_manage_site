const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'server', 'GmailDraft.js'),
  'utf8'
);
const pageSource = fs.readFileSync(
  path.join(root, 'pages', 'make-email.html'),
  'utf8'
);
const scriptSource = fs.readFileSync(
  path.join(root, 'scripts', 'make-email.html'),
  'utf8'
);
const calendarSource = fs.readFileSync(
  path.join(root, 'scripts', 'calendar.html'),
  'utf8'
);

let createdDraft = null;
const sandbox = {
  JSON,
  String,
  Number,
  Boolean,
  Object,
  Array,
  GmailApp: {
    createDraft: (to, subject, body, options) => {
      createdDraft = { to, subject, body, options };
    },
  },
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

sandbox.getDefaultRecipients_ = () => ({
  to: 'office@example.com',
  bcc: 'default-bcc@example.com',
});
sandbox.normalizePrivateEmail_ = email => String(email || '').trim().toLowerCase();
sandbox.participantContactCandidateData_ = () => [
  {
    candidateId: '11:2', name: '山田 花子', grade: 'A',
    state: '出場可能', canceled: false, selectable: true,
    _realEmail: 'alice@example.com',
  },
  {
    candidateId: '11:3', name: '佐藤 太郎', grade: 'C',
    state: 'キャンセル待ち1番', canceled: false, selectable: true,
    _realEmail: 'bob@example.com',
  },
  {
    candidateId: '11:4', name: '取消 三郎', grade: 'A',
    state: 'キャンセル', canceled: true, selectable: true,
    _realEmail: 'cancelled@example.com',
  },
  {
    candidateId: '11:5', name: '待機 四郎', grade: 'C',
    state: 'キャンセル待ち2番', canceled: false, selectable: true,
    _realEmail: 'waitlist@example.com',
  },
];

const created = JSON.parse(sandbox.createParticipantContactDraft(JSON.stringify({
  tournamentName: '第1回テスト大会',
  grades: ['A', 'C'],
  selectedCandidateIds: ['11:2', '11:3'],
  body: '',
})));
assert.strictEqual(created.ok, true);
assert.strictEqual(created.recipientCount, 2);
assert.strictEqual(created.subject, '第1回テスト大会　参加者への連絡');
assert.strictEqual(JSON.stringify(created).includes('alice@example.com'), false);
assert.deepStrictEqual(JSON.parse(JSON.stringify(createdDraft)), {
  to: 'office@example.com',
  subject: '第1回テスト大会　参加者への連絡',
  body: '',
  options: {
    bcc: 'alice@example.com,bob@example.com',
    name: '慶應かるた会',
  },
});
assert.strictEqual(
  createdDraft.options.bcc.includes('cancelled@example.com'),
  false,
  '個別選択されていないキャンセル済み申込者をBCCに含めない'
);
assert.strictEqual(
  createdDraft.options.bcc.includes('waitlist@example.com'),
  false,
  '個別選択されていない申込者をBCCに含めない'
);
assert.strictEqual(
  createdDraft.options.bcc.includes('default-bcc@example.com'),
  false,
  '出場者向けBCCへ既定BCCを混在させない'
);

createdDraft = null;
const staleSelection = JSON.parse(
  sandbox.createParticipantContactDraft(JSON.stringify({
    tournamentName: '第1回テスト大会',
    grades: ['A'],
    selectedCandidateIds: ['99:99'],
    body: '本文',
  }))
);
assert.match(staleSelection.error, /再読み込み/);
assert.strictEqual(createdDraft, null);
assert.strictEqual(
  JSON.stringify(staleSelection).includes('alice@example.com'),
  false,
  'エラー応答にメールアドレスを露出しない'
);
assert.strictEqual(
  sandbox.participantContactSafeError_(
    'invalid recipient: private.person@example.com'
  ),
  'invalid recipient: ***@example.com'
);

const candidates = JSON.parse(sandbox.getParticipantContactCandidates(
  JSON.stringify({
    tournamentName: '第1回テスト大会',
    grades: ['A', 'C'],
  })
));
assert.strictEqual(candidates.candidates.length, 4);
assert.strictEqual(candidates.candidates[1].state, 'キャンセル待ち1番');
assert.strictEqual(
  JSON.stringify(candidates).includes('@example.com'),
  false,
  '候補者一覧へ実メールを含めない'
);

createdDraft = null;
const canceledSelected = JSON.parse(
  sandbox.createParticipantContactDraft(JSON.stringify({
    tournamentName: '第1回テスト大会',
    grades: ['A'],
    selectedCandidateIds: ['11:4'],
    body: '申込者全員向け本文',
  }))
);
assert.strictEqual(canceledSelected.ok, true);
assert.strictEqual(createdDraft.options.bcc, 'cancelled@example.com');

assert.match(pageSource, /出場者への連絡/);
assert.match(pageSource, /data-contact-grade="A"/);
assert.doesNotMatch(pageSource, /id="me-contact-body"/);
assert.match(scriptSource, /createParticipantContactDraft/);
assert.match(scriptSource, /meGetContactGrades/);
assert.match(scriptSource, /meApplyContactScope/);
assert.match(scriptSource, /_meActiveTab === 4[\s\S]*\? ''/);
assert.match(scriptSource, /meScheduleParticipantContactCandidates/);
assert.match(
  scriptSource,
  /_meContactCandidatesByGrade\[grade\]/,
  '級ごとに取得結果を保持する'
);
assert.match(
  scriptSource,
  /selectedGrades\.filter\(grade =>[\s\S]*_meContactLoadingGrades\[grade\]/,
  '未取得の級だけを差分取得する'
);
const toggleGradeSource = scriptSource.match(
  /function meToggleContactGrade\(button\) \{[\s\S]*?\n\}/
)[0];
assert.doesNotMatch(
  toggleGradeSource,
  /meResetParticipantContactCandidates/,
  '級の切替では候補者一覧全体を破棄しない'
);
assert.match(
  scriptSource,
  /scope === 'all' \|\| !candidate\.canceled/,
  '未キャンセル／全申込者の初期選択を切り替える'
);
assert.match(
  calendarSource,
  /onCalendarLoaded[\s\S]*mePopulateParticipantContactTournaments\(\)/,
  '非同期取得後に出場者連絡用の大会プルダウンも更新する'
);

console.log('Participant contact draft regression checks passed.');
