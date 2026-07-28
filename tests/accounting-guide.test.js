const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const page = fs.readFileSync(path.join(root, 'pages/accounting-guide.html'), 'utf8');
const suitouPage = fs.readFileSync(path.join(root, 'pages/suitou.html'), 'utf8');
const suitouScript = fs.readFileSync(path.join(root, 'scripts/suitou.html'), 'utf8');
const suitou = fs.readFileSync(path.join(root, 'server/Suitou.js'), 'utf8');
const detail = fs.readFileSync(path.join(root, 'server/TournamentDetail.js'), 'utf8');

assert.doesNotMatch(index, /data-page="accounting-guide"/);
assert.match(index, /include\('pages\/accounting-guide'\)/);
assert.doesNotMatch(index, /id="help-btn"/);
assert.doesNotMatch(index, /include\('scripts\/help'\)/);
assert.match(page, /id="page-accounting-guide"/);
assert.match(page, /onclick="returnToSuitou\(\)"/);
assert.match(suitouPage, /onclick="showAccountingGuide\(\)"/);
assert.match(suitouScript, /function showAccountingGuide/);
assert.match(suitouScript, /function returnToSuitou/);
assert.match(page, /大会シートの文字列や再計算結果は会計の正本にしません/);
assert.match(page, /請求額－有効な支払い額/);
assert.match(page, /期限前キャンセル/);
assert.match(page, /期限後キャンセル/);
assert.match(page, /年度をまたいでも保持/);
assert.match(page, /利用可能なデポジット残高の全額を使用します/);
assert.match(page, /デポジット額と充当したい参加費が異なる場合は実行せず/);
assert.match(page, /選手ごとの支払い履歴やデポジット履歴は追加・削除しません/);
assert.match(page, /キャンセル状態の変更だけでは支払い履歴は作成されません/);
assert.match(suitou, /deposit: balance > 0/);
assert.match(detail, /taikaiRecordFullPaymentByPlayer_/);
assert.doesNotMatch(detail, /appendSuitouTx_/);
assert.match(suitou, /\/players\/.*\/deposits/);
assert.doesNotMatch(
  suitou,
  /function (addSuitouRow|deleteSuitouRow|convertToDeposit|updateSuitou|getPlayerDepositBalance_)/
);

console.log('Accounting guide regression checks passed.');
