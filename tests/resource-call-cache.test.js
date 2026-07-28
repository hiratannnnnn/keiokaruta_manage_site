const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const board = read('scripts/board.html');
const mail = read('scripts/mail-management.html');
const suitou = read('scripts/suitou.html');
const migration = read('scripts/sheet-migration.html');
const settings = read('scripts/settings.html');
const utils = read('scripts/utils.html');
const calendarPage = read('pages/calendar.html');
const matrixPage = read('pages/participation-matrix.html');
const index = read('index.html');

assert.match(board, /if \(_boardLoaded && !force\)/);
assert.match(board, /_boardLoaded = true/);
assert.match(mail, /if \(_mmLoaded\)/);
assert.match(mail, /_mmLoaded = true/);
assert.match(suitou, /if \(_suAllData\)/);
assert.doesNotMatch(
  suitou.match(/function suShowAddModal[\s\S]*?function suCloseAddModal/)[0],
  /\.getTournamentList\(\)/,
  '出納追加を開くたびに大会一覧を再取得してはいけません'
);
assert.doesNotMatch(
  suitou.match(/function suShowAddModal[\s\S]*?function suCloseAddModal/)[0],
  /ensureMembersLoaded/,
  'API台帳に含まれるplayer IDを利用し、名簿APIを追加取得してはいけません'
);
assert.match(migration, /if \(!sheetMigrationLoaded\) sheetMigrationPreview\(\)/);
assert.match(settings, /if \(settingsPageLoaded\) return/);
assert.match(settings, /function reloadSettingsPage/);
assert.match(utils, /if \(page === 'settings'\) initSettingsPage\(\)/);
assert.doesNotMatch(utils, /\.getPermissions\(\)/);
assert.match(utils, /if \(calendarLoaded\) \{\s*renderCalendar\(allTournaments\)/);
assert.match(calendarPage, /page-title-row[\s\S]*?onclick="loadCalendar\(\)">再読込/);
assert.match(matrixPage, /page-title-row[\s\S]*?loadParticipationMatrix\(true\)/);
assert.match(index, /掲示板や案内作成で使用する役職を選択してください/);
[
  'nav-form-create-btn',
  'nav-make-email-btn',
  'nav-mail-management-btn',
  'nav-suitou-btn',
  'nav-database-admin-btn',
  'nav-settings-btn',
].forEach(id => {
  assert.doesNotMatch(
    index,
    new RegExp(`id="${id}"[^>]*style="display:none;"`),
    `${id}を役職選択前から非表示にしてはいけません`
  );
});
assert.doesNotMatch(settings, /PERM_LABELS|renderPermTable|savePermSettings/);

console.log('Resource call cache regression checks passed.');
