const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const serverSource = fs.readFileSync(
  path.join(root, 'server/FormCreate.js'), 'utf8'
);
const browserSource = fs.readFileSync(
  path.join(root, 'scripts/form-create.html'), 'utf8'
);

const calls = [];
const responseSheet = { id: 88 };
const ranges = {
  B3: [['old title'], ['old grades']],
  B17: [['old lottery']],
  B28: [['old start'], ['old reminder'], ['old url']],
};
const sandbox = {
  Date,
  JSON,
  String,
  Boolean,
  Object,
  Array,
  encodeURIComponent,
  FormApp: {
    openById: () => ({
      removeDestination: () => calls.push('remove destination'),
    }),
  },
  DriveApp: {
    getFileById: () => ({
      isTrashed: () => false,
      setTrashed: () => calls.push('trash form'),
    }),
  },
  taikaiApiRequest_: (method, apiPath) => {
    calls.push(method + ' ' + apiPath);
  },
};
vm.createContext(sandbox);
vm.runInContext(serverSource, sandbox);

const state = {
  title: '第1回大会',
  grades: 'A級',
  sheetName: '第1回大会A級',
  formUrl: 'https://example.com/form',
  announceWritten: true,
  announceSheet: {
    getRange: a1 => ({
      setValues: values => calls.push(
        'restore ' + a1 + ' ' + JSON.stringify(values)
      ),
    }),
  },
  announceSnapshots: Object.keys(ranges).map(a1 => ({
    a1,
    values: ranges[a1],
  })),
  calendarWritten: true,
  calendarSheet: {
    getRange: (row, column) => ({
      setValue: value => calls.push(
        'restore calendar ' + row + ':' + column + ' ' + String(value)
      ),
    }),
  },
  calendarMutation: {
    row: 5,
    snapshots: [
      { column: 1, value: '' },
      { column: 3, value: '' },
      { column: 6, value: '' },
      { column: 8, value: '' },
      { column: 11, value: '' },
    ],
  },
  mailWritten: true,
  mailSheet: {
    getRange(row, column) {
      if (row === 3 && column === 1) {
        return { setValues: () => calls.push('restore mail row') };
      }
      return { setValue: () => calls.push('restore mail pointer') };
    },
  },
  mailMutation: {
    row: 3,
    previousValues: [['', '', '', '', '', '']],
  },
  mailPointerBefore: 2,
  destinationRequested: true,
  formId: 'form-1',
  spreadsheet: {
    getSheetById: () => responseSheet,
    deleteSheet: sheet => {
      assert.strictEqual(sheet, responseSheet);
      calls.push('delete response sheet');
    },
  },
  existingSheetIds: {},
  responseSheetId: '88',
  formCopyRequested: true,
  formFileId: 'form-1',
  dbTournamentCreated: true,
  dbTournamentId: '9',
};

assert.deepStrictEqual(
  JSON.parse(JSON.stringify(sandbox.formCreateCreatedResources_(state))),
  [
    'DB大会',
    'Googleフォーム',
    'フォーム回答シート',
    'メール管理行',
    'カレンダー行',
    '案内メール作成設定',
  ]
);
const rollback = sandbox.formCreateRollback_(state);
assert.strictEqual(rollback.complete, true);
assert.strictEqual(rollback.manual_action_required, false);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(rollback.unreverted_resources)),
  []
);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(rollback.actions.map(action => action.target))),
  [
    '案内メール作成設定',
    'カレンダー行',
    'メール管理行',
    'フォーム回答先の解除',
    'フォーム回答シート',
    'Googleフォーム',
    'DB大会',
  ]
);
assert.deepStrictEqual(calls.slice(-4), [
  'remove destination',
  'delete response sheet',
  'trash form',
  'DELETE /tournaments/9',
]);

const failedRollbackState = Object.assign({}, state, {
  announceWritten: false,
  calendarWritten: false,
  mailWritten: false,
  destinationRequested: false,
  responseSheetId: '',
  formFileId: 'unremovable-form',
  dbTournamentCreated: false,
});
const originalGetFileById = sandbox.DriveApp.getFileById;
sandbox.DriveApp.getFileById = () => ({
  isTrashed: () => false,
  setTrashed: () => {
    throw new Error('権限不足');
  },
});
const failedRollback = sandbox.formCreateRollback_(failedRollbackState);
assert.strictEqual(failedRollback.complete, false);
assert.strictEqual(failedRollback.manual_action_required, true);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(failedRollback.unreverted_resources)),
  ['Googleフォーム']
);
assert.match(failedRollback.actions[0].error, /権限不足/);
sandbox.DriveApp.getFileById = originalGetFileById;

const uncertainFormState = Object.assign({}, failedRollbackState, {
  formFileId: '',
  formCopyRequested: true,
});
const uncertainResources = sandbox.formCreateCreatedResources_(
  uncertainFormState
);
assert.ok(
  uncertainResources.includes('Googleフォーム（作成された可能性あり）')
);
const uncertainRollback = sandbox.formCreateRollback_(uncertainFormState);
assert.strictEqual(uncertainRollback.complete, false);
assert.strictEqual(uncertainRollback.manual_action_required, true);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(uncertainRollback.unreverted_resources)),
  ['Googleフォーム（作成確認）']
);

const uncertainDbState = Object.assign({}, uncertainFormState, {
  formCopyRequested: false,
  dbTournamentCreationUncertain: true,
});
assert.ok(
  sandbox.formCreateCreatedResources_(uncertainDbState)
    .includes('DB大会（作成された可能性あり）')
);
const uncertainDbRollback = sandbox.formCreateRollback_(uncertainDbState);
assert.strictEqual(uncertainDbRollback.manual_action_required, true);
assert.deepStrictEqual(
  JSON.parse(JSON.stringify(uncertainDbRollback.unreverted_resources)),
  ['DB大会（作成確認）']
);

assert.match(browserSource, /function fcFormatCreateFailure/);
assert.match(browserSource, /失敗した処理:/);
assert.match(browserSource, /作成状況: Googleフォーム・回答シート等は作成されていません/);
assert.match(browserSource, /作成途中の変更はすべて巻き戻しました/);
assert.match(browserSource, /一部を自動で巻き戻せませんでした/);
assert.doesNotMatch(
  serverSource,
  /一時障害の場合はフォーム作成を継続/
);
assert.doesNotMatch(
  serverSource,
  /tournamentSheetValidateGradeOwnership_/
);
assert.doesNotMatch(serverSource, /calendarSheet\.deleteRow/);
assert.doesNotMatch(serverSource, /mailSheet\.deleteRow/);
assert.ok(
  serverSource.indexOf('state.mailWritten = true;')
  < serverSource.indexOf(
    'mailSheet.getRange(nextRow, 1, 1, 6).setValues'
  )
);
assert.ok(
  serverSource.indexOf('state.announceWritten = true;')
  < serverSource.indexOf(
    'announceSheet.getRange(3, 2, 2, 1).setValues'
  )
);

console.log('Form creation rollback regression checks passed.');
