const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'server/TaikaiApi.js'),
  'utf8'
);

function loadClient(properties, fetch) {
  const sandbox = {
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => properties[key] || '',
      }),
    },
    UrlFetchApp: { fetch },
    Utilities: {},
    CacheService: {
      getScriptCache: () => ({ get: () => null, put: () => {} }),
    },
    MailApp: { sendEmail: () => {} },
    encodeURIComponent,
    Object,
    String,
    Boolean,
    JSON,
    Date,
  };

  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  return sandbox;
}

let request = null;
const client = loadClient({
  TAIKAI_API_BASE_URL: 'http://api.example.com/api/v1/',
  TAIKAI_API_TOKEN: 'test-token',
}, (url, options) => {
  request = { url, options };
  return {
    getResponseCode: () => 200,
    getContentText: () => JSON.stringify({ data: { status: 'ok' } }),
  };
});

const result = client.taikaiApiRequest_('GET', '/health');
assert.deepStrictEqual(JSON.parse(JSON.stringify(result)), { status: 'ok' });
assert.strictEqual(request.url, 'http://api.example.com/api/v1/health');
assert.strictEqual(
  request.options.headers.Authorization,
  'Bearer test-token'
);

const missingToken = loadClient({
  TAIKAI_API_BASE_URL: 'http://api.example.com/api/v1',
}, () => {
  assert.fail('トークン未設定時は外部通信してはいけません');
});
assert.throws(
  () => missingToken.taikaiApiRequest_('GET', '/health'),
  /TAIKAI_API_TOKEN/
);

const emptyToken = loadClient({
  TAIKAI_API_BASE_URL: 'http://api.example.com/api/v1',
  TAIKAI_API_TOKEN: '   ',
}, () => {
  assert.fail('空白トークンでは外部通信してはいけません');
});
assert.throws(
  () => emptyToken.taikaiApiRequest_('GET', '/health'),
  /TAIKAI_API_TOKEN/
);

console.log('Taikai API Bearer authentication checks passed.');
