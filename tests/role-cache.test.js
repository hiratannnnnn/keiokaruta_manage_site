const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const calendar = fs.readFileSync(path.join(root, 'scripts/calendar.html'), 'utf8');
const utils = fs.readFileSync(path.join(root, 'scripts/utils.html'), 'utf8');

assert.match(calendar, /const SELECTABLE_ROLES = \['申込係', '振込係', '案内係', '会長副会長'\]/);
assert.match(calendar, /function restorePreviousRoleAndStart/);
assert.match(calendar, /SELECTABLE_ROLES\.includes\(savedRole\)/);
assert.match(calendar, /selectLoginRole\(roleBtn\);\s*submitRoleSelection\(\);/);
assert.match(calendar, /localStorage\.removeItem\('karuta_role'\)/);
assert.match(utils, /restorePreviousRoleAndStart\(\)/);
assert.doesNotMatch(
  utils,
  /if \(savedRole\) \{\s*const roleBtn/,
  '役職を選択状態へ戻すだけで、開始処理を忘れてはいけません'
);

console.log('Role cache regression checks passed.');
