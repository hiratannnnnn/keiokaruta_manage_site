const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'server/TaikaiApi.js'),
  'utf8'
);
const databaseAdminSource = fs.readFileSync(
  path.join(root, 'scripts/database-admin.html'),
  'utf8'
);
assert.doesNotMatch(source, /recipient_group_id/);
assert.doesNotMatch(databaseAdminSource, /recipient_group_id/);

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

const unauthorized = loadClient({
  TAIKAI_API_BASE_URL: 'http://api.example.com/api/v1',
  TAIKAI_API_TOKEN: 'wrong-token',
}, () => ({
  getResponseCode: () => 401,
  getContentText: () => JSON.stringify({
    error: {
      code: 'unauthorized',
      message: '認証に失敗しました。',
      fields: {},
      request_id: 'req_0123456789abcdef',
    },
  }),
}));
assert.throws(
  () => unauthorized.taikaiApiRequest_('POST', '/tournaments', {}),
  error => {
    assert.match(error.message, /TAIKAI_API_TOKENとPHPのAPI_TOKEN/);
    assert.match(error.message, /HTTP 401/);
    assert.match(error.message, /code: unauthorized/);
    assert.match(error.message, /POST \/tournaments/);
    assert.match(error.message, /request ID: req_0123456789abcdef/);
    assert.strictEqual(error.api_error_code, 'unauthorized');
    assert.strictEqual(error.taikai_request_id, 'req_0123456789abcdef');
    return true;
  }
);

const validationFailure = loadClient({
  TAIKAI_API_BASE_URL: 'http://api.example.com/api/v1',
  TAIKAI_API_TOKEN: 'test-token',
}, () => ({
  getResponseCode: () => 400,
  getContentText: () => JSON.stringify({
    error: {
      code: 'validation_error',
      message: '入力内容を確認してください。',
      fields: {
        grade: 'A〜Eで指定してください。',
        'unsafe key': '表示しない',
      },
      request_id: 'req_fedcba9876543210',
    },
  }),
}));
assert.throws(
  () => validationFailure.taikaiApiRequest_('POST', '/schedules', {}),
  error => {
    assert.match(error.message, /grade: A〜Eで指定してください。/);
    assert.doesNotMatch(error.message, /unsafe key/);
    assert.deepStrictEqual(
      JSON.parse(JSON.stringify(error.validation_fields)),
      { grade: 'A〜Eで指定してください。' }
    );
    return true;
  }
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

const finder = loadClient({}, () => {
  assert.fail('大会検索の単体確認では外部通信してはいけません');
});
finder.taikaiApiRequest_ = () => [];
assert.throws(
  () => finder.taikaiFindTournament_('未登録大会'),
  /DBに大会が登録されていません.*年度のシート→DB完全同期/
);
finder.taikaiApiRequest_ = () => [
  { id: '1', name: '重複大会' },
  { id: '2', name: '重複大会' },
];
assert.throws(
  () => finder.taikaiFindTournament_('重複大会'),
  /DBに同名の大会が2件/
);
finder.taikaiApiRequest_ = () => [
  { id: '3', name: '対象大会記念' },
  { id: '4', name: '対象大会' },
];
assert.strictEqual(finder.taikaiFindTournament_('対象大会').id, '4');
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    finder.taikaiEnsureTournamentWithState_('対象大会')
  )),
  {
    tournament: { id: '4', name: '対象大会' },
    created: false,
  }
);
finder.taikaiApiRequest_ = (method) => method === 'GET'
  ? [] : { id: '5', name: '新規大会' };
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(
    finder.taikaiEnsureTournamentWithState_('新規大会')
  )),
  {
    tournament: { id: '5', name: '新規大会' },
    created: true,
  }
);

finder.taikaiApiRequest_ = method => {
  if (method === 'GET') return [];
  const error = new Error('timeout');
  error.taikai_api_error = true;
  error.transient = true;
  throw error;
};
assert.throws(
  () => finder.taikaiEnsureTournamentWithState_('応答不明大会'),
  error => error.tournament_creation_uncertain === true
);

const paymentJobClient = loadClient({}, () => {
  assert.fail('メールジョブpayload確認では外部通信してはいけません');
});
paymentJobClient.taikaiFormatDate_ = value => value;
paymentJobClient.taikaiFindTournament_ = () => ({ id: '12' });
let paymentJobPayload = null;
paymentJobClient.taikaiApiRequest_ = (method, path, body) => {
  if (method === 'GET') {
    return [{ id: '31', grade: 'A' }, { id: '32', grade: 'B' }];
  }
  if (method === 'POST' && path === '/email-jobs') {
    paymentJobPayload = body;
    return { id: '41' };
  }
  return { id: path };
};
paymentJobClient.taikaiRegisterPaymentEmailJob_(
  '対象大会', 'AB級', '2026-08-14 07:50:00', '2026-08-21'
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(paymentJobPayload)),
  {
    scheduled_at: '2026-08-14T07:50:00+09:00',
    mail_type: 'payment_confirmation',
    announcement_id: null,
    schedule_ids: ['31', '32'],
  }
);

console.log('Taikai API Bearer authentication checks passed.');
