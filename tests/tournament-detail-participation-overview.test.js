const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const page = read('pages/calendar.html');
const script = read('scripts/calendar.html');
const matrixScript = read('scripts/participation-matrix.html');
const style = read('style.html');
const server = read('server/TournamentDetail.js');
const matrixServer = read('server/ParticipationMatrix.js');

[
  'detail-participation-btn',
  'detail-participation-modal',
  'detail-participation-status',
  'detail-participation-table',
  'detail-participation-tbody',
].forEach(id => assert.match(page, new RegExp('id="' + id + '"')));

assert.ok(
  page.indexOf('id="toggle-cols-btn"')
    < page.indexOf('id="detail-participation-btn"'),
  '出場大会ボタンを詳細列ボタンの後に配置する'
);
assert.match(page, /出場回数/);
assert.match(page, /出場大会/);
assert.match(script, /prefetchDetailParticipationOverview\(\s*name, requestId/);
assert.match(script, /function detailParticipationFiscalYear\(tournament\)/);
assert.ok(
  script.indexOf('prefetchDetailParticipationOverview(name, requestId)')
    < script.indexOf('.getTournamentDetail(name)'),
  '詳細画面を開いた時点で出場一覧を先読みする'
);
assert.match(script, /\.getParticipationMatrix\(fiscalYear\)/);
assert.match(script, /participationMatrixGetCachedData\(fiscalYear\)/);
assert.match(script, /participationMatrixCacheData\(data\)/);
assert.doesNotMatch(script, /detailParticipationMatrixCache/);
assert.match(matrixScript, /let participationMatrixDataByFiscalYear = \{\}/);
assert.match(matrixScript, /function participationMatrixCacheData\(data\)/);
assert.match(matrixScript, /function participationMatrixGetCachedData\(fiscalYear\)/);
assert.match(matrixScript, /function participationMatrixInvalidateCache\(\)/);
assert.match(matrixScript, /participationMatrixCacheGeneration\+\+/);
assert.match(matrixScript, /cacheGeneration !== participationMatrixCacheGeneration/);
assert.match(matrixScript, /participationMatrixData = cached \|\| participationMatrixData/);
assert.match(script, /const cacheGeneration = participationMatrixCacheGeneration/);
assert.match(script, /state\.applicants\.forEach\(applicant =>/);
assert.match(script, /detailParticipationResolvePlayer/);
assert.match(script, /participation\.mark === 'unsanctioned'/);
assert.match(script, /return '（' \+ name \+ '）'/);
assert.match(script, /\.join\('、'\)/);
assert.match(script, /function detailParticipationSanctionedCount\(player, items\)/);
assert.match(script, /sanctioned_schedule_ids/);
assert.match(script, /function detailParticipationTournamentDate\(tournament\)/);
assert.match(script, /function detailParticipationCompareChronologically\(/);
assert.match(script, /detailParticipationTournamentDate\(leftTournament\)/);
assert.doesNotMatch(script, /tournamentOrder/);
assert.match(script, /matrix\.truncated === true/);
assert.match(script, /APIの取得上限に達したため/);
assert.match(server, /grade: record\.grade/);
assert.match(matrixServer, /function getParticipationMatrix\(fiscalYearInput\)/);
assert.match(style, /\.detail-participation-table-wrap[\s\S]*?overflow: auto/);
assert.match(style, /\.detail-participation-history-cell[\s\S]*?overflow: auto/);
assert.match(style, /\.detail-section-header[\s\S]*?flex-wrap: wrap/);

console.log('Tournament detail participation overview checks passed.');
