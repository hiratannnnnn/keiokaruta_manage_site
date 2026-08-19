const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'server', 'Cancellation.js'), 'utf8'
);
const sandbox = { Date, String, Number, Object, Array, isNaN };
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

const status = (canceledAt, lotteryResultDate, extra) =>
  sandbox.cancellationStatusFromRecord_(Object.assign({
    canceled_at: canceledAt,
    lottery_result_date: lotteryResultDate,
  }, extra || {}));

assert.strictEqual(status(null, '2026-09-25'), 'none');
assert.strictEqual(status('', '2026-09-25'), 'none');
assert.strictEqual(
  status('2026-09-24T23:59:59+09:00', '2026-09-25'),
  'before'
);
// 抽選公開日当日は公開後キャンセル。
assert.strictEqual(
  status('2026-09-25T00:00:00+09:00', '2026-09-25'),
  'after'
);
assert.strictEqual(
  status('2026-09-25T23:59:59+09:00', '2026-09-25'),
  'after'
);
// タイムゾーン付き日時は日本時間の日付へ変換して比較する。
assert.strictEqual(
  status('2026-09-24T15:30:00Z', '2026-09-25'),
  'after'
);
assert.strictEqual(
  status('2026-09-24T14:30:00Z', '2026-09-25'),
  'before'
);
assert.strictEqual(status('2026-09-24 23:59:59', ''), 'unknown');
assert.strictEqual(status('2026-09-24 23:59:59', null), 'unknown');
// 開催日・申込期限などから推測せず、旧汎用ラベルも根拠にしない。
assert.strictEqual(status('2026-09-24 23:59:59', '', {
  cancellation_timing: 'after',
  held_on: '2026-09-30',
  application_deadline: '2026-09-01',
}), 'unknown');
assert.strictEqual(
  sandbox.cancellationLabelFromRecord_({
    canceled_at: '2026-09-25', lottery_result_date: '2026-09-25',
  }),
  '公開後キャンセル'
);
assert.strictEqual(
  sandbox.cancellationCountsAsParticipant_({
    canceled_at: '2026-09-24', lottery_result_date: '2026-09-25',
  }), false
);
assert.strictEqual(
  sandbox.cancellationCountsAsParticipant_({
    canceled_at: '2026-09-25', lottery_result_date: '2026-09-25',
  }), true
);
assert.strictEqual(
  sandbox.cancellationCountsAsParticipant_({
    canceled_at: '2026-09-24', lottery_result_date: null,
  }), false
);

const detail = fs.readFileSync(
  path.join(root, 'server', 'TournamentDetail.js'), 'utf8'
);
const calendar = fs.readFileSync(
  path.join(root, 'scripts', 'calendar.html'), 'utf8'
);
const style = fs.readFileSync(path.join(root, 'style.html'), 'utf8');
const results = fs.readFileSync(
  path.join(root, 'scripts', 'results.html'), 'utf8'
);
const playerPage = fs.readFileSync(
  path.join(root, 'pages', 'player.html'), 'utf8'
);
assert.match(detail, /function tournamentDetailCountsAsParticipant_\(record\)/);
assert.match(detail, /return cancellationCountsAsParticipant_\(record\)/);
assert.match(calendar, /record\.cancellation_status === 'before'/);
assert.match(calendar, /cancel-row-before/);
assert.match(calendar, /cancel-row-after/);
assert.match(style, /tr\.cancel-row-after > td/);
assert.match(calendar, /detail-cancel-after/);
assert.match(
  style,
  /#detail-person-table > tbody > tr\.detail-cancel-after:nth-child\(even\) > td/
);
assert.match(calendar, /record\.cancellation_label/);
assert.match(
  calendar,
  /cancellationStatus !== 'none' && cancellationStatus !== 'after'/
);
assert.match(calendar, /const actionHtml = cancellationStatus === 'after'/);
assert.doesNotMatch(calendar, /支払詳細・訂正/);
assert.doesNotMatch(calendar, /右欄からキャンセルを解除してください/);
const actionStart = calendar.indexOf('function renderPayButtonsHtml(');
const actionEnd = calendar.indexOf('function renderEntryActionCells(', actionStart);
assert.ok(actionStart >= 0 && actionEnd > actionStart);
const actionSandbox = {
  String,
  detailSheetVersion: 2,
  escAttr: value => String(value),
};
vm.createContext(actionSandbox);
vm.runInContext(
  calendar.slice(actionStart, actionEnd)
    + '\nthis.renderPayButtonsHtml = renderPayButtonsHtml;',
  actionSandbox
);
const afterCancellationActions = actionSandbox.renderPayButtonsHtml(
  '', '未払い', '山田 太郎', 2, '10', '20', 'after'
);
assert.match(afterCancellationActions, /振込済み/);
assert.doesNotMatch(afterCancellationActions, /キャンセル待ち/);
const beforeCancellationActions = actionSandbox.renderPayButtonsHtml(
  '', '未払い', '山田 太郎', 2, '10', '20', 'before'
);
assert.doesNotMatch(beforeCancellationActions, /振込済み/);
assert.match(results, /result\.cancellation_label/);
assert.match(playerPage, /<th>キャンセル状態<\/th>/);

console.log('Cancellation status regression checks passed.');
