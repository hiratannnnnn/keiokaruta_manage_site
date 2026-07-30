const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'server/MailManagement.js'),
  'utf8'
);
let mailRow = [
  '第1回大会',
  'AB級',
  '2099-08-13 08:00:00',
  'リマインダー',
  '第1回大会AB級　案内',
  'https://example.com/form',
  '',
  '',
];
let draft = null;
const sandbox = {
  Date,
  JSON,
  String,
  Number,
  Boolean,
  Object,
  Array,
  Set,
  CONFIG: {
    SPREADSHEET_ID: 'spreadsheet-1',
    SHEET_NAMES: { MAIL: 'メール管理' },
  },
  SpreadsheetApp: {
    openById: () => ({
      getSheetByName: () => ({
        getRange: () => ({
          getValues: () => [mailRow],
        }),
      }),
    }),
  },
  Utilities: {
    formatDate: () => '2099年08月13日 08時00分',
  },
  GmailApp: {
    createDraft: (to, subject, body, options) => {
      draft = { to, subject, body, options };
    },
  },
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox);
sandbox.makeParticipantList_ = () => 'A級：山田';
sandbox.quoteEmail_ = () => '> 元の案内';
sandbox.createEmailBody_ = (
  tournamentName, grades, participantList
) => tournamentName + grades + '\n' + participantList;
sandbox.getDefaultRecipients_ = () => ({
  to: 'to@example.com',
  bcc: 'bcc@example.com',
});

const input = JSON.stringify({
  rowNum: 6,
  includeNotPaid: false,
  sendDateTime: '2099-08-13 08:00:00',
});
const preview = JSON.parse(sandbox.getReminderPreview(input));
assert.strictEqual(preview.error, undefined);
assert.strictEqual(preview.body, '第1回大会AB級\nA級：山田');
assert.strictEqual(draft, null, 'dry-runではGmail下書きを作らない');

const created = JSON.parse(sandbox.createReminderDraftFromPreview(input));
assert.strictEqual(created.error, undefined);
assert.deepStrictEqual(JSON.parse(JSON.stringify(draft)), {
  to: 'to@example.com',
  subject: '第1回大会AB級　リマインダー',
  body: '第1回大会AB級\nA級：山田',
  options: {
    bcc: 'bcc@example.com',
    name: '慶應かるた会',
  },
});

draft = null;
mailRow = mailRow.slice();
mailRow[3] = '振込確認';
const rejected = JSON.parse(sandbox.createReminderDraftFromPreview(input));
assert.match(rejected.error, /リマインダー行だけ/);
assert.strictEqual(draft, null);

console.log('Reminder draft regression checks passed.');
