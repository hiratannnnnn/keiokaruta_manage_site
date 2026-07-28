const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const server = fs.readFileSync(path.join(root, 'server/ParticipationMatrix.js'), 'utf8');
const script = fs.readFileSync(path.join(root, 'scripts/participation-matrix.html'), 'utf8');
const page = fs.readFileSync(path.join(root, 'pages/participation-matrix.html'), 'utf8');
const config = fs.readFileSync(path.join(root, 'server/config.js'), 'utf8');

assert.match(config, /PLAYER_NOTES:\s*'選手管理メモ'/);
assert.match(server, /function saveParticipationPlayerNote/);
assert.match(server, /getScriptLock\(\)/);
assert.match(server, /normalizedMemo\.length > 1000/);
assert.match(server, /sheet\.hideSheet\(\)/);
assert.match(server, /visibleNotes/);
assert.match(page, /id="participation-matrix-memo-panel"/);
assert.match(page, /maxlength="1000"/);
assert.match(script, /participationMatrixSelectedPlayerId/);
assert.match(script, /\.saveParticipationPlayerNote\(playerId, memo\)/);
assert.match(script, /textContent =\s*player\.name/);
assert.doesNotMatch(script, /innerHTML =\s*result\.memo/);

const sandbox = {};
vm.runInNewContext(server, sandbox);
const values = [
  ['101', '要確認', '2026-07-28'],
  ['102', '', '2026-07-28'],
  ['', 'キーなし', '2026-07-28'],
];
const sheet = {
  getLastRow: () => 4,
  getRange: () => ({ getValues: () => values }),
};
const ss = {
  getSheetByName: () => sheet,
};
sandbox.CONFIG = { SHEET_NAMES: { PLAYER_NOTES: '選手管理メモ' } };
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.participationMatrixNotes_(ss))),
  { 101: '要確認' }
);

console.log('Participation note regression checks passed.');
