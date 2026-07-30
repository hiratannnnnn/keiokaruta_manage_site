const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const page = fs.readFileSync(path.join(root, 'pages/calendar.html'), 'utf8');
const script = fs.readFileSync(path.join(root, 'scripts/calendar.html'), 'utf8');
const detailServer = fs.readFileSync(
  path.join(root, 'server/TournamentDetail.js'), 'utf8'
);
const mailServer = fs.readFileSync(
  path.join(root, 'server/MailManagement.js'), 'utf8'
);

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
assert.match(script, /data-source-row=/);
assert.match(script, /data-entry-id=/);
assert.match(script, /data-player-id=/);
assert.match(script, /detailSelectionStatusIndex/);
assert.match(script, /function renderEntryActionCells/);
assert.match(script, /function detailIdentityFromButton/);
assert.match(script, /setDetailPayStatus\(\s*tournamentName, identity\.sourceRow, identity\.entryId/);
assert.match(script, /setEntryCancellationStatus\(/);
assert.match(script, /function openEntryPaymentDetail/);
assert.match(script, /reverseEntryPaymentFromDetail\(/);
assert.doesNotMatch(script, /function updateDetailPayCell/);
assert.doesNotMatch(script, /String\(row\[2\] \|\| ''\)/);
assert.doesNotMatch(script, /入金取消は出納管理から/);
assert.doesNotMatch(script, /DB申込が未同期です/);
assert.match(script, /DB未登録/);
assert.match(script, />公開前へ</);
assert.match(script, />公開後へ</);
assert.match(script, />繰り越し</);
assert.match(script, /\.getPlayerDeposit\(identity\.playerId, playerName\)/);
assert.match(page, /id="entry-cancellation-modal"/);
assert.match(page, /現在の入金をデポジットへ戻す/);
assert.match(page, /id="entry-payment-detail-modal"/);
assert.match(page, /管理操作はPCから行うことを推奨します/);
assert.match(detailServer, /'POST', '\/admin\/tournament-sheet-snapshot'/);
assert.match(detailServer, /emailMapRows_\(\)/);
assert.doesNotMatch(detailServer, /pseudonymousEmailFor_\(/);
assert.strictEqual(
  (detailServer.match(/taikaiRecordFullPaymentByEntry_\(/g) || []).length,
  1,
  '一回の支払操作で入金APIを一度だけ呼ぶ'
);
assert.match(
  detailServer,
  /tournamentDetailEnrichRecords_\(name, records\)/
);
assert.match(
  mailServer,
  /tournamentDetailEnrichRecords_\(tournamentName \+ grades, records\)/
);
assert.match(detailServer, /const detailSettings = \[/);
assert.doesNotMatch(
  detailServer,
  /Object\.keys\(structure\.management\.metadata\)\.map/
);
assert.match(detailServer, /function tournamentDetailPaymentTimingLabel_/);
assert.match(detailServer, /\['支払時期', tournamentDetailPaymentTimingLabel_/);
assert.match(detailServer, /function tournamentDetailCancellationTiming_/);
assert.match(detailServer, /function tournamentDetailCountsAsParticipant_/);
assert.match(detailServer, /function setEntryCancellationStatus/);
assert.match(detailServer, /function getEntryPaymentDetail/);
assert.match(detailServer, /function reverseEntryPaymentFromDetail/);
assert.doesNotMatch(
  detailServer,
  /taikaiRecordFullPaymentByEntry_[\s\S]{0,250}setValue\('済'\)/
);

console.log('Tournament detail UI regression checks passed.');
