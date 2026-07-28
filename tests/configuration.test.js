const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const config = fs.readFileSync(path.join(root, 'server/config.js'), 'utf8');
const board = fs.readFileSync(path.join(root, 'server/Board.js'), 'utf8');
const settings = fs.readFileSync(path.join(root, 'server/Settings.js'), 'utf8');
const gitignore = fs.readFileSync(path.join(root, '.gitignore'), 'utf8');
const claspignore = fs.readFileSync(path.join(root, '.claspignore'), 'utf8');
const envExample = fs.readFileSync(path.join(root, '.env.example'), 'utf8');

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
assert.match(settings, /'LINE_LINK_WEBHOOK_SECRET'/);
assert.match(settings, /'LINE_LINK_BINDING_SECRET'/);
assert.match(settings, /'TAIKAI_API_TOKEN'/);
assert.match(settings, /'FORM_RESPONSE_NOTIFICATION_TO'/);
assert.match(envExample, /^TAIKAI_API_TOKEN=$/m);
assert.match(envExample, /^FORM_RESPONSE_NOTIFICATION_TO=$/m);
assert.doesNotMatch(board, /BOARD_SPREADSHEET_ID_\s*=/);
assert.match(gitignore, /^\.env$/m);
assert.match(claspignore, /^\.env$/m);

console.log('Configuration regression checks passed.');
