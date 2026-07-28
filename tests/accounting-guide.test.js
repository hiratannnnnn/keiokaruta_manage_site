const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

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
assert.match(page, /通常の振込済み/);
assert.match(page, /デポジットとして完了/);
assert.match(page, /期限前キャンセル/);
assert.match(page, /期限後キャンセル/);
assert.match(page, /年度をまたいでも保持/);
assert.match(page, /利用可能なデポジット残高の全額を使用します/);
assert.match(page, /デポジット額と充当したい参加費が異なる場合は実行せず/);
assert.match(page, /選手個人のデポジットを自動で探して消費する操作ではありません/);
assert.match(page, /期限後キャンセルへの変更だけでは、参加費のプラス記録は自動作成されません/);
assert.match(suitou, /function getPlayerDepositBalance_/);
assert.match(suitou, /deposit: balance > 0/);
assert.match(detail, /getPlayerDepositBalance_\(txRows, normalizedName\)/);
assert.doesNotMatch(detail, /const depositAmt = Number\(txRows\[j\]\[1\]\)/);

const sandbox = {};
vm.runInNewContext(suitou, sandbox);
assert.strictEqual(sandbox.getPlayerDepositBalance_([
  ['山田　太郎', 3000, 'デポジット'],
  ['山田 太郎', -3000, 'デポジット'],
], '山田 太郎'), 0, '使用済みデポジットを再利用可能にしてはいけません');
assert.strictEqual(sandbox.getPlayerDepositBalance_([
  ['山田 太郎', 2000, 'デポジット'],
  ['山田　太郎', 1000, 'デポジット'],
  ['山田 太郎', -500, 'デポジット'],
  ['山田 太郎', 9999, '別の事由'],
], '山田　太郎'), 2500, 'デポジットは正負全行の純残高で判定します');

console.log('Accounting guide regression checks passed.');
