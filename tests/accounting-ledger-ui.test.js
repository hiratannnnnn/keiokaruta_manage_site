const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const server = fs.readFileSync(path.join(root, 'server/Suitou.js'), 'utf8');
const script = fs.readFileSync(path.join(root, 'scripts/suitou.html'), 'utf8');
const page = fs.readFileSync(path.join(root, 'pages/suitou.html'), 'utf8');
const detail = fs.readFileSync(path.join(root, 'server/TournamentDetail.js'), 'utf8');
const taikai = fs.readFileSync(path.join(root, 'server/TaikaiApi.js'), 'utf8');
const calendar = fs.readFileSync(path.join(root, 'server/Calendar.js'), 'utf8');

assert.match(server, /'GET', '\/admin\/accounting-ledger'/);
assert.match(server, /limit:\s*500/);
assert.match(server, /next_cursor/);
assert.match(server, /function addSuitouApiTransaction/);
assert.match(server, /idempotency_key/);
assert.match(server, /function reverseSuitouApiTransaction/);
assert.doesNotMatch(server, /refreshTournamentSheetV2ByEntryId_/);
assert.doesNotMatch(server, /大会シートの書戻しに失敗/);
assert.match(server, /payment\.method \|\| ''\) !== 'deposit'/);
assert.match(server, /deposit\.transaction_type \|\| ''\) !== 'deposit_applied'/);
assert.match(script, /function suIdempotencyKey/);
assert.match(script, /crypto\.randomUUID/);
assert.match(script, /\.addSuitouApiTransaction/);
assert.match(script, /\.reverseSuitouApiTransaction/);
assert.doesNotMatch(script, /\.updateSuitou\(\)/);
assert.doesNotMatch(page, /シートから再計算/);
assert.doesNotMatch(page, /onclick="suUpdate\(\)"/);
assert.doesNotMatch(page, /carried_over/);
assert.match(page, /DBの支払い・デポジット履歴/);
assert.match(page, /デポジットを参加費へ充当/);
assert.doesNotMatch(detail, /appendSuitouTx_/);
assert.doesNotMatch(detail, /taikaiRecordFullPaymentByPlayer_/);
assert.doesNotMatch(taikai, /function taikaiRecordFullPaymentByPlayer_/);
assert.match(detail, /taikaiRecordFullPaymentByEntry_/);
assert.doesNotMatch(taikai, /\/entries\/.*\/payment['"]/);
assert.doesNotMatch(calendar, /appendSuitouTx_/);
assert.doesNotMatch(calendar, /removeSuitouNegTxByReason_/);
assert.doesNotMatch(calendar, /getSheetByName\('出納管理'\)/);
assert.match(calendar, /'PATCH',[\s\S]*'\/tournaments\/'/);
assert.match(detail, /taikaiRecordFullPaymentByEntry_/);

const sandbox = {};
vm.runInNewContext(server, sandbox);
assert.strictEqual(sandbox.suitouDepositSignedAmount_({
  amount_yen: 3000,
  transaction_type: 'deposit_received',
}), 3000);
assert.strictEqual(sandbox.suitouDepositSignedAmount_({
  amount_yen: 1000,
  transaction_type: 'deposit_applied',
}), -1000);
assert.strictEqual(sandbox.suitouDepositSignedAmount_({
  amount_yen: 500,
  transaction_type: 'adjustment',
  direction: 'credit',
}), 500);
assert.strictEqual(sandbox.suitouDepositSignedAmount_({
  amount_yen: 500,
  transaction_type: 'adjustment',
  direction: 'debit',
}), -500);

console.log('Accounting ledger UI regression checks passed.');
