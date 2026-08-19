const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
let selectionWrites = 0;
const sheet = {};
const sandbox = {
  CONFIG: { SPREADSHEET_ID: 'main' },
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => sheet }),
  },
  tournamentSheetStructure_: () => ({ version: 1 }),
  tournamentSheetResponseRecord_: () => ({
    source_row: 2,
    entry_id: '',
    name: '山田 太郎',
  }),
  tournamentSheetSelectionStatusRange_: () => ({
    setValue: () => { selectionWrites++; },
  }),
};
vm.runInNewContext(read('server/TournamentDetail.js'), sandbox);

const result = JSON.parse(
  sandbox.setDetailPayStatus('第1回大会A級', 2, '', '済', false)
);
assert.match(result.error, /先に大会シートv2へ移行/);
assert.strictEqual(selectionWrites, 0);

const serverSource = read('server/TournamentDetail.js');
const uiSource = read('scripts/calendar.html');
assert.match(serverSource, /const isLegacySheet = structure\.version === 1/);
assert.match(serverSource, /slice\(0, legacyDisplayColumnCount\)/);
assert.match(uiSource, /v2移行後に利用できます/);
assert.match(uiSource, /detailSheetVersion !== 2/);

console.log('Legacy selection/payment separation checks passed.');
