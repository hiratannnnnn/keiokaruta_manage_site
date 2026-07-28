const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');

const utils = read('scripts/utils.html');
const calendar = read('scripts/calendar.html');
const suitou = read('scripts/suitou.html');
const board = read('scripts/board.html');
const mail = read('scripts/mail-management.html');

assert.match(utils, /\.replace\(\/"\/g, '&quot;'\)/);
assert.match(utils, /\.replace\(\/'\/g, '&#39;'\)/);
assert.match(utils, /function setSanitizedRichText/);
assert.match(utils, /url\.protocol === 'https:' \|\| url\.protocol === 'http:'/);
assert.match(utils, /if \(!\/\^https\?:\\\/\\\/\/i\.test\(text\)\) return false/);

assert.doesNotMatch(
  calendar,
  /innerHTML\s*=\s*\(t && t\.announcementHtml\)/,
  '案内HTMLを無加工でinnerHTMLへ代入してはいけません'
);
assert.match(calendar, /setSanitizedRichText\(content, t\.announcementHtml\)/);
assert.doesNotMatch(calendar, /goToPlayerSearch\('\$\{name\}'\)/);
assert.doesNotMatch(suitou, /suSelectName\('\$\{/);
assert.doesNotMatch(board, /boardInsertTournament\('\$\{/);
assert.match(mail, /isSafeHttpUrl\(row\.formLink\)/);

console.log('UI security regression checks passed.');
