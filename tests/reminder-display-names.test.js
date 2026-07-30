const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const source = fs.readFileSync(
  path.join(__dirname, '..', 'server', 'MailManagement.js'),
  'utf8'
);
const sandbox = {
  Date,
  String,
  Number,
  Boolean,
  Object,
  Array,
  JSON,
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

const displayNames = sandbox.mailManagementReminderDisplayNames_([
  { name: '平田　智也' },
  { name: '平田 智恵' },
  { name: '平田 太郎' },
  { name: '佐藤　花子' },
  { name: '山田 智' },
  { name: '山田 智也' },
]);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(displayNames)),
  ['平田智也', '平田智恵', '平田太', '佐藤', '山田智', '山田智也']
);

assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    sandbox.mailManagementReminderDisplayNames_([
      { name: '鈴木　花子', email: 'first@keio.jp' },
      { name: '鈴木 花子', email: 'second@example.com' },
    ])
  )),
  ['鈴木花子（***@keio.jp）', '鈴木花子（***@example.com）']
);

const sameDomain = sandbox.mailManagementReminderDisplayNames_([
  { name: '伊藤　直子', email: 'aoki.secret@keio.jp' },
  { name: '伊藤 直子', email: 'arai.private@keio.jp' },
]);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sameDomain)),
  ['伊藤直子（ao***@keio.jp）', '伊藤直子（ar***@keio.jp）']
);
assert.doesNotMatch(sameDomain.join(''), /aoki\.secret|arai\.private/);

assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    sandbox.mailManagementReminderDisplayNames_([
      { name: '加藤　隆', email: 'abc@keio.jp' },
      { name: '加藤 隆', email: 'abd@keio.jp' },
    ])
  )),
  ['加藤隆（***@keio.jp・同名1）', '加藤隆（***@keio.jp・同名2）']
);

assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    sandbox.mailManagementReminderDisplayNames_([
      { name: '渡辺　健', email: 'invalid-address' },
      { name: '渡辺 健', email: 'valid@example.com' },
    ])
  )),
  ['渡辺健（同名1）', '渡辺健（同名2）']
);

assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    sandbox.mailManagementReminderDisplayNames_([
      { name: '高橋一郎' },
      { name: '森　次郎' },
    ])
  )),
  ['高橋一郎', '森']
);

assert.match(
  source,
  /mailType === 'リマインダー'/
);

console.log('Reminder display-name regression checks passed.');
