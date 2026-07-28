const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const page = fs.readFileSync(path.join(root, 'pages/calendar.html'), 'utf8');
const script = fs.readFileSync(path.join(root, 'scripts/calendar.html'), 'utf8');

[
  'detail-ops-section',
  'btn-toggle-official',
  'btn-register',
  'btn-count',
  'btn-pay-done',
  'btn-complete-detail',
  'btn-delete',
].forEach(id => {
  const matches = page.match(new RegExp('id="' + id + '"', 'g')) || [];
  assert.strictEqual(matches.length, 1, id + ' must appear exactly once');
});

assert.match(script, /function setDetailOperationState/);
assert.match(script, /完了処理中/);
assert.doesNotMatch(script, /楽観的UI: 即座に一覧から除いて戻る/);
assert.match(script, /tournament\.applyDone && tournament\.payDone/);
assert.match(script, /function updateDetailCompletionAvailability/);
assert.match(script, /function isCurrentTournamentDetail/);
assert.match(script, /let detailRequestSequence = 0/);
assert.match(script, /requestId !== detailRequestSequence/);
assert.match(script, /detailPendingNotice\.tournamentName/);
assert.match(page, /管理操作はPCから行うことを推奨します/);

console.log('Tournament detail UI regression checks passed.');
