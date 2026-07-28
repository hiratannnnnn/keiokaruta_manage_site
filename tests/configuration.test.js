const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const config = fs.readFileSync(path.join(root, 'server/config.js'), 'utf8');
const board = fs.readFileSync(path.join(root, 'server/Board.js'), 'utf8');
const gitignore = fs.readFileSync(path.join(root, '.gitignore'), 'utf8');
const claspignore = fs.readFileSync(path.join(root, '.claspignore'), 'utf8');

[
  'MAIN_SPREADSHEET_ID',
  'FORM_FOLDER_ID',
  'FORM_TEMPLATE_ID',
  'TRASH_SPREADSHEET_ID',
  'BOARD_SPREADSHEET_ID',
].forEach(key => assert.match(config, new RegExp("'" + key + "'")));

assert.doesNotMatch(config, /configValue_\('[A-Z_]+'\s*,/);
assert.doesNotMatch(config, /['"][A-Za-z0-9_-]{30,}['"]/);
assert.match(board, /CONFIG\.BOARD_SPREADSHEET_ID/);
assert.doesNotMatch(board, /BOARD_SPREADSHEET_ID_\s*=/);
assert.match(gitignore, /^\.env$/m);
assert.match(claspignore, /^\.env$/m);

console.log('Configuration regression checks passed.');
