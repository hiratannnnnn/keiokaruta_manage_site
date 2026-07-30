const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const utils = read('scripts/utils.html');
const utilityBlock = utils.match(
  /function csvDownloadCell[\s\S]+?(?=function isSafeHttpUrl)/
);
assert.ok(utilityBlock, 'CSV utility functions must exist');

let downloaded = null;
let blobParts = null;
const sandbox = {
  Blob: function Blob(parts, options) {
    blobParts = parts;
    this.options = options;
  },
  URL: {
    createObjectURL: () => 'blob:test',
    revokeObjectURL: () => {},
  },
  document: {
    body: { appendChild: () => {} },
    createElement: () => ({
      style: {},
      click() { downloaded = this.download; },
      remove: () => {},
    }),
  },
  setTimeout: callback => callback(),
};
vm.createContext(sandbox);
vm.runInContext(utilityBlock[0], sandbox);
assert.strictEqual(
  sandbox.csvDownloadDateStamp(new Date('2026-07-28T15:30:00Z')),
  '2026-07-29'
);
sandbox.downloadCsvFile(
  '2026/出場一覧.csv',
  ['氏名', 'メモ'],
  [['平田　智也', '=HYPERLINK("https://example.com")']]
);
assert.strictEqual(downloaded, '2026_出場一覧.csv');
assert.ok(blobParts[0].startsWith('\uFEFF'));
assert.ok(blobParts[0].includes(
  '"平田　智也","\'=HYPERLINK(""https://example.com"")"'
));

[
  ['pages/player.html', 'results-download'],
  ['pages/participation-matrix.html', 'participation-matrix-download'],
  ['pages/suitou.html', 'su-tx-download'],
].forEach(([file, id]) => {
  assert.match(read(file), new RegExp('id="' + id + '"'));
});
assert.match(read('scripts/results.html'), /downloadTournamentResultsCsv/);
assert.match(
  read('scripts/participation-matrix.html'),
  /downloadParticipationMatrixCsv/
);
assert.match(read('scripts/suitou.html'), /suDownloadTransactionsCsv/);

console.log('CSV download regression checks passed.');
