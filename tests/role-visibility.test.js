const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const utils = read('scripts/utils.html');
const calendar = read('scripts/calendar.html');
const board = read('scripts/board.html');
const settingsPage = read('pages/settings.html');
const settingsServer = read('server/Settings.js');

assert.doesNotMatch(utils, /hasPermission|currentPermissions|getPermissions/);
assert.doesNotMatch(calendar, /hasPermission|currentPermissions|権限エラー/);
assert.doesNotMatch(utils, /currentRole === '会長副会長'/);
assert.doesNotMatch(utils, /const isAnnouncer =/);
assert.doesNotMatch(utils, /const canSeeSuitou =/);

assert.match(calendar, /detail-ops-section'\)\.style\.display = 'grid'/);
assert.match(calendar, /detail-delete-section'\)\.style\.display = 'block'/);
assert.match(calendar, /btn-pay-done'\)\.style\.display = 'block'/);
assert.match(calendar, /detail-complete-section'\)\.style\.display = 'block'/);
assert.match(board, /const canDelete = true/);

assert.doesNotMatch(settingsPage, /権限設定|perm-table-wrap|perm-save-btn/);
assert.doesNotMatch(settingsServer, /DEFAULT_PERMISSIONS_|getPermissions|savePermissions|PERMISSIONS/);

console.log('Role-independent visibility regression checks passed.');
