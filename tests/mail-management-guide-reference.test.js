const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const page = fs.readFileSync(
  path.join(root, 'pages/mail-management.html'), 'utf8'
);
const script = fs.readFileSync(
  path.join(root, 'scripts/mail-management.html'), 'utf8'
);
const makeEmailScript = fs.readFileSync(
  path.join(root, 'scripts/make-email.html'), 'utf8'
);
const style = fs.readFileSync(path.join(root, 'style.html'), 'utf8');

assert.match(page, /onclick="mmShowAddFromMakeEmail\(\)"/);
assert.match(page, /案内作成の内容から登録/);
assert.match(page, /リマインダー行をクリックすると/);
assert.match(page, /id="mm-create-draft-btn"/);
assert.match(page, /id="mm-add-modal-prompt"/);
assert.match(page, /この内容をメール管理に登録しますか/);
assert.match(script, /function mmShowAddFromMakeEmail\(\)/);
assert.match(script, /meGetActiveGrades\(\)/);
assert.match(script, /_meActiveTab !== 0 && _meActiveTab !== 1/);
assert.match(script, /document\.getElementById\('me-form-url'\)/);
assert.match(script, /document\.getElementById\('me-pay-real'\)/);
assert.match(script, /function mmDatetimeBeforeDate\(/);
assert.match(script, /function mmScheduledDateTimeFromInternalDeadline\(/);
assert.match(script, /isPayment \? 'me-pay-kaInai' : 'me-apply-kaInai'/);
assert.match(
  script,
  /mmScheduledDateTimeFromInternalDeadline\(\s*internalDeadline, mailType/
);
assert.match(script, /案内作成タブの入力内容を参照しました/);
assert.match(script, /function mmSelectReminderAfterNavigation\(/);
assert.match(script, /function mmOpenPendingReminderSelection\(/);
assert.match(script, /mm-previewable-row/);
assert.match(script, /function mmCreateReminderDraft\(/);
assert.match(script, /createReminderDraftFromPreview\(/);
assert.match(
  makeEmailScript,
  /mmSelectReminderAfterNavigation\(title, grades, sendDateTime\)/
);
assert.doesNotMatch(
  makeEmailScript,
  /mmSelectReminderAfterNavigation\(title, grades, sendDateTime\);\s*initMailManagement\(\)/
);
assert.match(page, /案内は会内締切の前日、振込確認は会内期限の当日/);
assert.match(
  makeEmailScript,
  /mmScheduledDateTimeFromInternalDeadline\(\s*internalDeadline, '振込確認'/
);
assert.match(makeEmailScript, /振込会内期限が未入力/);
assert.match(makeEmailScript, /参加表明締切（会内）が未入力/);
assert.match(
  makeEmailScript,
  /mmScheduledDateTimeFromInternalDeadline\(\s*internalDeadline, 'リマインダー'/
);
assert.match(
  script,
  /selection\.preferredSendDateTime \|\| row\.sendDateTime/
);
assert.match(
  makeEmailScript,
  /'', true, sendDateTime/
);
assert.match(style, /\.mm-guide-register-btn\s*\{/);
assert.match(style, /\.mm-previewable-row\s*\{/);
assert.match(style, /min-height:\s*76px/);
assert.ok(
  page.indexOf('id="mail-management-table"')
  < page.indexOf('<details class="mm-maintenance-details">'),
  '補助・同期機能はメール管理表より下に配置する'
);
assert.ok(
  page.indexOf('id="mail-management-table"')
  < page.indexOf('<div class="panel-header">案内Gmail ID紐付け</div>'),
  '案内Gmail紐付けはメール管理表より下に配置する'
);
assert.match(page, /<details class="mm-maintenance-details">\s*<summary>補助・同期機能<\/summary>/);
assert.strictEqual(
  (page.match(/<div class="panel-header">メールジョブDB同期<\/div>/g) || []).length,
  1,
  'メールジョブDB同期パネルを重複させない'
);

const sandbox = { Date, Number, String };
vm.createContext(sandbox);
vm.runInContext(
  script.replace(/^\s*<script>\s*/, '').replace(/\s*<\/script>\s*$/, ''),
  sandbox
);
assert.strictEqual(
  sandbox.mmDatetimeBeforeDate('2026-08-14', 1, 8, 0),
  '2026-08-13T08:00'
);
assert.strictEqual(
  sandbox.mmDatetimeBeforeDate('2026-03-01', 1, 7, 50),
  '2026-02-28T07:50'
);
assert.strictEqual(
  sandbox.mmScheduledDateTimeFromInternalDeadline(
    '2026-08-14', '振込確認'
  ),
  '2026-08-14T07:50'
);
assert.strictEqual(
  sandbox.mmScheduledDateTimeFromInternalDeadline(
    '2026-08-14', 'リマインダー'
  ),
  '2026-08-13T08:00'
);
assert.strictEqual(sandbox.mmDatetimeBeforeDate('未入力', 1, 8, 0), '');
assert.strictEqual(
  sandbox.mmFindReminderRow([
    {
      rowNum: 8,
      tournamentName: '第1回大会',
      grades: 'AB級',
      mailType: 'リマインダー',
      sent: '済',
    },
    {
      rowNum: 6,
      tournamentName: '第1回大会',
      grades: 'AB級',
      mailType: 'リマインダー',
      sent: '',
    },
    {
      rowNum: 9,
      tournamentName: '第1回大会',
      grades: 'AB級',
      mailType: '振込確認',
      sent: '',
    },
  ], '第1回大会', 'AB級').rowNum,
  6
);

console.log('Mail management guide-reference checks passed.');
