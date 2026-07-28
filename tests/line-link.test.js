const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'server/LineLink.js'), 'utf8');
const codeSource = fs.readFileSync(path.join(root, 'server/Code.js'), 'utf8');
const envExample = fs.readFileSync(path.join(root, '.env.example'), 'utf8');

const properties = {
  LINE_LINK_WEBHOOK_SECRET: 'webhook-test-secret',
  LINE_LINK_BINDING_SECRET: 'binding-test-secret',
};
const sentEmails = [];
const sandbox = {
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: key => properties[key] || '',
    }),
  },
  Utilities: {
    Charset: { UTF_8: 'UTF-8' },
    DigestAlgorithm: { SHA_256: 'SHA-256' },
    computeDigest: (algorithm, value) =>
      Array.from(crypto.createHash('sha256').update(String(value)).digest())
        .map(value => value > 127 ? value - 256 : value),
    computeHmacSha256Signature: (value, secret) =>
      Array.from(crypto.createHmac('sha256', secret).update(String(value)).digest())
        .map(value => value > 127 ? value - 256 : value),
  },
  MailApp: {
    sendEmail: message => sentEmails.push(message),
  },
  LockService: {
    getScriptLock: () => ({
      waitLock: () => {},
      releaseLock: () => {},
    }),
  },
  encodeURIComponent,
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

assert.match(codeSource, /function doPost\(e\)/);
assert.match(codeSource, /lineLinkWebhookResponse_\(e\)/);
assert.match(envExample, /^LINE_LINK_WEBHOOK_SECRET=$/m);
assert.match(envExample, /^LINE_LINK_BINDING_SECRET=$/m);
assert.doesNotMatch(source, /Logger\./);
assert.doesNotMatch(
  source.match(/function lineLinkWebhookResponse_[\s\S]*$/)[0],
  /message:\s*String\(e\.message/,
  'Webhookへ内部例外文をそのまま返してはいけません'
);

const binding = sandbox.lineLinkBindingHash_('U-line-user');
assert.match(binding, /^[0-9a-f]{64}$/);
assert.strictEqual(binding, sandbox.lineLinkBindingHash_('U-line-user'));
assert.notStrictEqual(binding, sandbox.lineLinkBindingHash_('U-other-user'));
assert.strictEqual(sandbox.lineLinkSafeEquals_('same', 'same'), true);
assert.strictEqual(sandbox.lineLinkSafeEquals_('same', 'different'), false);

const rows = [
  ['real@example.com', 'v1-a@example.invalid', '', '', ''],
  ['other@example.com', 'v1-b@example.invalid', '2', 'U-existing', ''],
];
const map = {
  indexes: {
    '実メールアドレス': 0,
    'DB用疑似メールアドレス': 1,
    player_id: 2,
    'LINEユーザーID': 3,
    'LINE連携日時': 4,
  },
  sheet: {
    getLastRow: () => rows.length + 1,
    getRange: () => ({ getValues: () => rows }),
  },
  headers: ['実メールアドレス', 'DB用疑似メールアドレス', 'player_id', 'LINEユーザーID', 'LINE連携日時'],
};
assert.throws(
  () => sandbox.lineLinkAssertAvailable_(map, 2, '1', 'U-existing'),
  /既に連携/
);
assert.doesNotThrow(
  () => sandbox.lineLinkAssertAvailable_(map, 2, '1', 'U-new')
);
rows[0][2] = '1';
rows[0][3] = 'U-same';
assert.doesNotThrow(
  () => sandbox.lineLinkAssertAvailable_(map, 2, '1', 'U-same'),
  '同一連携の再実行は冪等に扱う'
);
assert.throws(
  () => sandbox.lineLinkAssertAvailable_(map, 2, '1', 'U-new'),
  /別のLINE/
);

sandbox.normalizePrivateEmail_ = email => String(email).trim().toLowerCase();
sandbox.lineLinkPlayerByEmail_ = () => ({
  player: { id: '10' },
  pseudo_email: 'v1-test@example.invalid',
});
sandbox.lineLinkEmailMap_ = () => map;
sandbox.lineLinkFindEmailRow_ = () => ({ row_number: 2, values: rows[0] });
sandbox.lineLinkAssertAvailable_ = () => {};
sandbox.taikaiApiRequest_ = (method, apiPath) => {
  assert.strictEqual(method, 'POST');
  assert.strictEqual(apiPath, '/line-link-verifications');
  return {
    verification_id: '55',
    code: '123456',
    expires_at: '2026-07-28 12:00:00',
  };
};
const started = sandbox.lineLinkStart_({
  email: 'REAL@example.com',
  line_user_id: 'U-line-user',
});
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(started)),
  {
    ok: true,
    status: 'otp_sent',
    verification_id: '55',
    expires_at: '2026-07-28 12:00:00',
  }
);
assert.strictEqual(sentEmails.length, 1);
assert.strictEqual(sentEmails[0].to, 'real@example.com');
assert.match(sentEmails[0].body, /123456/);
assert.doesNotMatch(JSON.stringify(started), /123456|real@example|U-line-user/i);

let saved = null;
sandbox.taikaiApiRequest_ = () => ({
  verified: true,
  player_id: '10',
});
sandbox.lineLinkSave_ = (playerId, lineUserId) => {
  saved = { playerId, lineUserId };
};
const verified = sandbox.lineLinkVerify_({
  verification_id: '55',
  code: '123456',
  line_user_id: 'U-line-user',
});
assert.deepStrictEqual(JSON.parse(JSON.stringify(verified)), {
  ok: true,
  status: 'linked',
});
assert.deepStrictEqual(saved, {
  playerId: '10',
  lineUserId: 'U-line-user',
});
assert.doesNotMatch(JSON.stringify(verified), /123456|U-line-user/);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.lineLinkPublicError_({
    message: 'mail delivery failed for real@example.com',
  }))),
  { status: 'failed', message: '処理に失敗しました。' },
  'メール送信基盤の例外文を外部へ返してはいけません'
);

console.log('LINE link regression checks passed.');
