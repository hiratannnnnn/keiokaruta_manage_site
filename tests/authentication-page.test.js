const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = relativePath =>
  fs.readFileSync(path.join(root, relativePath), 'utf8');
const index = read('index.html');
const page = read('pages/authentication.html');
const script = read('scripts/authentication.html');
const settingsPage = read('pages/settings.html');
const settingsScript = read('scripts/settings.html');
const utils = read('scripts/utils.html');
const server = read('server/Authentication.js');

assert.match(index, /id="nav-authentication-btn" style="display:none;"/);
assert.match(index, /include\('pages\/authentication'\)/);
assert.match(index, /include\('scripts\/authentication'\)/);
assert.match(settingsPage, /id="settings-show-authentication"/);
assert.match(settingsScript, /const AUTHENTICATION_NAV_KEY = 'karuta_show_authentication'/);
assert.match(settingsScript, /function consumeAuthenticationNavigationActivation/);
assert.match(utils, /initAuthenticationPage\(\);[\s\S]*?consumeAuthenticationNavigationActivation\(\)/);
assert.match(page, /同姓同名・異メールアドレス/);
assert.doesNotMatch(script, /innerHTML/);

const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(server, sandbox);
const groups = sandbox.authenticationDuplicateGroups_([
  ['first@example.com', '山田 太郎'],
  ['FIRST@example.com', '山田　太郎'],
  ['second@example.com', '山田太郎'],
  ['hanako@example.com', '山田 花子'],
  ['', '山田 太郎'],
], 1, 0, 2);

assert.deepStrictEqual(JSON.parse(JSON.stringify(groups)), [{
  name: '山田 太郎',
  emails: ['first@example.com', 'second@example.com'],
  rows: [2, 3, 4],
}]);

console.log('Authentication page regression checks passed.');
