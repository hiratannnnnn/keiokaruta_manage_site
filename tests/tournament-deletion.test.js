const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const apiCalls = [];
const deletedSheets = [];
const deletedCalendarRows = [];
const deletedMailRows = [];
let formTrashed = false;

const targetSheet = { getName: () => '第1回大会AB級' };
const siblingSheet = { getName: () => '第1回大会CD級' };
const calendarSheet = {
  getLastRow: () => 4,
  getRange: () => ({
    getValues: () => [[], [], ['第1回大会AB級'], ['第1回大会CD級']],
  }),
  deleteRow: row => deletedCalendarRows.push(row),
};
const mailSheet = {
  getLastRow: () => 4,
  getRange: () => ({
    getValues: () => [
      ['第1回大会', 'AB級'],
      ['第1回大会', 'CD級'],
    ],
  }),
  deleteRow: row => deletedMailRows.push(row),
};
const spreadsheet = {
  getSheetByName(name) {
    if (name === '第1回大会AB級') return targetSheet;
    if (name === 'カレンダー') return calendarSheet;
    if (name === 'メール管理') return mailSheet;
    return null;
  },
  getSheets: () => [targetSheet, siblingSheet],
  deleteSheet: sheet => deletedSheets.push(sheet.getName()),
};
const sandbox = {
  CONFIG: {
    SPREADSHEET_ID: 'main',
    TRASH_SPREADSHEET_ID: 'trash',
    SHEET_NAMES: {
      CALENDAR: 'カレンダー',
      MEMBERS: '名簿',
      MAIL: 'メール管理',
    },
  },
  SpreadsheetApp: {
    openById: id => id === 'main' ? spreadsheet : { getId: () => 'trash' },
  },
  FormApp: {
    DestinationType: { SPREADSHEET: 'spreadsheet' },
    openById: () => ({ setDestination() {} }),
  },
  DriveApp: {
    getFileById: () => ({
      isTrashed: () => formTrashed,
      setTrashed: value => { formTrashed = value; },
    }),
  },
  tournamentSheetStructure_: () => ({
    version: 2,
    management: {
      metadata: { 'tournament ID': '10' },
      schedules: {
        A: { row: ['A', '', '', '101'] },
        B: { row: ['B', '', '', '102'] },
      },
    },
  }),
  tournamentSheetFormId_: () => 'form-1',
  tournamentSheetBaseName_: () => '第1回大会',
  tournamentSheetDeclaredGrades_: () => ['A', 'B'],
  taikaiDeleteTournamentSchedules_: (tournamentId, scheduleIds) => {
    apiCalls.push({ tournamentId, scheduleIds });
    return {
      deleted_schedule_ids: scheduleIds,
      remaining_schedule_ids: ['103', '104'],
      tournament_deleted: false,
    };
  },
};
vm.runInNewContext(read('server/DeleteTournament.js'), sandbox);

const result = JSON.parse(sandbox.deleteTournament('第1回大会AB級'));
assert.strictEqual(result.ok, true);
assert.strictEqual(result.tournament_deleted, false);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(apiCalls)),
  [{ tournamentId: '10', scheduleIds: ['101', '102'] }]
);
assert.deepStrictEqual(deletedCalendarRows, [3]);
assert.deepStrictEqual(deletedMailRows, [3]);
assert.deepStrictEqual(deletedSheets, ['第1回大会AB級']);
assert.strictEqual(formTrashed, true);
assert.strictEqual(siblingSheet.getName(), '第1回大会CD級');

const apiSource = read('server/TaikaiApi.js');
const uiSource = read('scripts/calendar.html');
assert.match(apiSource, /DELETE[\s\S]*schedule_ids: ids/);
assert.match(uiSource, /兄弟フォームの大会データは維持されています/);

let googleWritesAfterApiFailure = 0;
const apiFailureSandbox = Object.assign({}, sandbox, {
  taikaiDeleteTournamentSchedules_: () => {
    throw new Error('API unavailable');
  },
  FormApp: {
    DestinationType: { SPREADSHEET: 'spreadsheet' },
    openById: () => {
      googleWritesAfterApiFailure++;
      return { setDestination() {} };
    },
  },
  DriveApp: {
    getFileById: () => ({
      isTrashed: () => false,
      setTrashed: () => { googleWritesAfterApiFailure++; },
    }),
  },
});
vm.runInNewContext(read('server/DeleteTournament.js'), apiFailureSandbox);
const apiFailure = JSON.parse(
  apiFailureSandbox.deleteTournament('第1回大会AB級')
);
assert.strictEqual(apiFailure.partial, false);
assert.match(apiFailure.error, /API unavailable/);
assert.strictEqual(googleWritesAfterApiFailure, 0);

let sheetDeletesAfterPartial = 0;
const partialSpreadsheet = Object.assign({}, spreadsheet, {
  deleteSheet: () => { sheetDeletesAfterPartial++; },
});
const partialSandbox = Object.assign({}, sandbox, {
  SpreadsheetApp: {
    openById: id => id === 'main'
      ? partialSpreadsheet : { getId: () => 'trash' },
  },
  FormApp: {
    DestinationType: { SPREADSHEET: 'spreadsheet' },
    openById: () => ({
      setDestination: () => { throw new Error('destination failed'); },
    }),
  },
  DriveApp: {
    getFileById: () => ({
      isTrashed: () => false,
      setTrashed() {},
    }),
  },
});
vm.runInNewContext(read('server/DeleteTournament.js'), partialSandbox);
const partial = JSON.parse(
  partialSandbox.deleteTournament('第1回大会AB級')
);
assert.strictEqual(partial.partial, true);
assert.match(partial.error, /同じ削除操作を再実行/);
assert.strictEqual(sheetDeletesAfterPartial, 0);

console.log('Tournament schedule deletion checks passed.');
