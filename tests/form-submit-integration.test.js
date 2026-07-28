const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');

function freshJournal() {
  return {
    record: {
      'イベントキー': 'event-1',
      'DB状態': 'pending',
      'v2書戻し状態': 'pending',
      '名簿状態': 'pending',
      '追加申込通知状態': 'pending',
      'player ID': '',
      'entry ID': '',
      'schedule ID': '',
    },
  };
}

function handlerHarness() {
  const journal = freshJournal();
  const calls = {
    register: 0,
    record: 0,
    refresh: 0,
    roster: 0,
    notification: 0,
  };
  let registrationFailure = false;
  let structureFailure = false;
  const sheet = {
    getName: () => '第1回大会A級',
    getParent: () => ({ getId: () => 'main' }),
  };
  const sandbox = {
    CONFIG: { SPREADSHEET_ID: 'main' },
    SpreadsheetApp: { openById: () => ({ getId: () => 'main' }) },
    formSubmitEventKey_: () => 'event-1',
    formSubmitJournalLoadOrCreate_: () => ({ record: journal.record }),
    formSubmitJournalSave_: (target, changes) => {
      Object.assign(target.record, changes);
      return target.record;
    },
    formSubmitNormalizeName_: value => String(value).trim(),
    tournamentSheetStructure_: () => {
      if (structureFailure) throw new Error('broken v2 structure');
      return { version: 2 };
    },
    taikaiRegisterEntry_: () => {
      calls.register++;
      if (registrationFailure) throw new Error('temporary API failure');
      return {
        player: { id: 'p1' },
        entry: { id: 'e1', schedule_id: 's1' },
      };
    },
    taikaiApiRequest_: () => ({
      participation_fee_yen: 2500,
      paid_yen: 0,
      balance_yen: 2500,
      status: 'unpaid',
    }),
    LockService: {
      getScriptLock: () => ({ waitLock() {}, releaseLock() {} }),
    },
  };
  vm.runInNewContext(read('server/FormSubmit.js'), sandbox);
  sandbox.recordFormResponseInTournamentSheetV2_ = (
    target, structure, sourceRow, response, result, apiError
  ) => {
    calls.record++;
    assert.strictEqual(target, sheet);
    assert.strictEqual(sourceRow, 2);
    if (registrationFailure) assert(apiError);
    else assert.strictEqual(result.entry.id, 'e1');
  };
  sandbox.refreshSiblingTournamentSheetsV2AfterResponse_ = () => {
    calls.refresh++;
  };
  sandbox.formSubmitMaintainRoster_ = () => {
    calls.roster++;
  };
  sandbox.formSubmitSendLateRegistration_ = target => {
    calls.notification++;
    sandbox.formSubmitJournalSave_(target, {
      '追加申込通知状態': 'not_required',
    });
  };
  const event = {
    range: {
      getSheet: () => sheet,
      getRow: () => 2,
    },
    values: ['2026/07/29', 'A級', 'user@example.com', '山田 太郎'],
    namedValues: {
      '参加級': ['A級'],
      'メールアドレス': ['user@example.com'],
      '氏名': ['山田 太郎'],
    },
  };
  return {
    sandbox,
    journal,
    calls,
    event,
    setRegistrationFailure: value => { registrationFailure = value; },
    setStructureFailure: value => { structureFailure = value; },
  };
}

{
  const harness = handlerHarness();
  const first = harness.sandbox.registerFormResponseToDatabaseUnlocked_(
    harness.event
  );
  const second = harness.sandbox.registerFormResponseToDatabaseUnlocked_(
    harness.event
  );
  assert.strictEqual(first.ok, true);
  assert.strictEqual(second.ok, true);
  assert.deepStrictEqual(harness.calls, {
    register: 1,
    record: 1,
    refresh: 1,
    roster: 1,
    notification: 1,
  });
  assert.strictEqual(harness.journal.record['DB状態'], 'done');
  assert.strictEqual(harness.journal.record['v2書戻し状態'], 'done');
  assert.strictEqual(harness.journal.record['名簿状態'], 'done');
  assert.throws(
    () => harness.sandbox.formAnswerValue_({
      '本人氏名': ['山田 太郎'],
      '振込名義人氏名': ['山田 花子'],
    }, title => title.includes('氏名'), '氏名', true),
    /一意に特定/
  );
}

{
  const harness = handlerHarness();
  harness.setRegistrationFailure(true);
  assert.throws(
    () => harness.sandbox.registerFormResponseToDatabaseUnlocked_(harness.event),
    /一部未完了/
  );
  assert.strictEqual(harness.journal.record['DB状態'], 'error');
  assert.strictEqual(
    harness.journal.record['v2書戻し状態'], 'waiting_for_db'
  );
  assert.strictEqual(harness.journal.record['名簿状態'], 'done');
  harness.setRegistrationFailure(false);
  const retried = harness.sandbox.registerFormResponseToDatabaseUnlocked_(
    harness.event
  );
  assert.strictEqual(retried.ok, true);
  assert.deepStrictEqual(harness.calls, {
    register: 2,
    record: 2,
    refresh: 1,
    roster: 1,
    notification: 1,
  });
  assert.strictEqual(harness.journal.record['DB状態'], 'done');
  assert.strictEqual(harness.journal.record['v2書戻し状態'], 'done');
}

{
  const harness = handlerHarness();
  harness.setStructureFailure(true);
  assert.throws(
    () => harness.sandbox.registerFormResponseToDatabaseUnlocked_(harness.event),
    /一部未完了/
  );
  assert.deepStrictEqual(harness.calls, {
    register: 1,
    record: 0,
    refresh: 0,
    roster: 1,
    notification: 1,
  });
  assert.strictEqual(harness.journal.record['DB状態'], 'done');
  assert.strictEqual(harness.journal.record['v2書戻し状態'], 'error');
  assert.strictEqual(harness.journal.record['名簿状態'], 'done');
}

{
  const saved = [];
  let drafts = 0;
  let sends = 0;
  const message = {
    getId: () => 'message-1',
    getPlainBody: () => 'body TAIKAI_FORM_EVENT_event-1',
  };
  const sandbox = {
    CONFIG: {
      SPREADSHEET_ID: 'main',
      SHEET_NAMES: {
        MEMBERS: '名簿',
        FORM_SUBMIT_JOURNAL: 'フォーム送信処理',
      },
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA-256' },
      Charset: { UTF_8: 'UTF-8' },
      computeDigest: (algorithm, value) =>
        Array.from(crypto.createHash('sha256').update(value).digest())
          .map(value => value > 127 ? value - 256 : value),
      formatDate: () => '2026-07-29',
    },
    configValue_: () => 'admin@example.com',
    tournamentSheetBaseName_: () => '第1回大会',
    taikaiFindTournament_: () => ({
      id: 't1',
      registration_completed: true,
    }),
    taikaiApiRequest_: () => [{
      id: 's1',
      grade: 'A',
      application_deadline: '2026-07-29',
    }],
    GmailApp: {
      search: () => [],
      createDraft: (to, subject, body) => {
        drafts++;
        assert.strictEqual(to, 'admin@example.com');
        assert.match(subject, /追加申込/);
        assert.match(body, /TAIKAI_FORM_EVENT_event-1/);
        return {
          getId: () => 'draft-1',
          send: () => {
            sends++;
            return message;
          },
        };
      },
    },
    ScriptApp: {
      EventType: { ON_FORM_SUBMIT: 'ON_FORM_SUBMIT' },
      getProjectTriggers: () => [{
        getHandlerFunction: () => 'registerFormResponseToDatabase',
        getTriggerSource: () => 'SPREADSHEETS',
        getEventType: () => 'ON_FORM_SUBMIT',
        getTriggerSourceId: () => 'main',
      }],
    },
  };
  vm.runInNewContext(read('server/FormSubmitEffects.js'), sandbox);
  const keySheet = {
    getParent: () => ({ getId: () => 'main' }),
    getSheetId: () => 10,
    getRange: () => ({
      getValue: () => new Date('2026-07-29T00:00:00Z'),
    }),
  };
  assert.strictEqual(
    sandbox.formSubmitEventKey_(keySheet, 2),
    sandbox.formSubmitEventKey_(keySheet, 2)
  );
  const rosterRows = [
    ['', '姓', '', '氏名'],
  ];
  const rosterSheet = {
    getDataRange: () => ({ getValues: () => rosterRows.map(row => row.slice()) }),
    getLastRow: () => rosterRows.length,
    insertRowAfter: rowNumber => {
      rosterRows.splice(rowNumber, 0, ['', '', '', '']);
    },
    getRange: rowNumber => ({
      setValues: values => { rosterRows[rowNumber - 1] = values[0].slice(); },
    }),
  };
  const rosterSpreadsheet = {
    getSheetByName: name => name === '名簿' ? rosterSheet : null,
  };
  sandbox.formSubmitMaintainRoster_(rosterSpreadsheet, ' 山田　太郎 ');
  sandbox.formSubmitMaintainRoster_(rosterSpreadsheet, '山田 太郎');
  assert.strictEqual(rosterRows.length, 2);
  assert.deepStrictEqual(
    JSON.parse(JSON.stringify(rosterRows[1])),
    ['', '山田', '', '山田 太郎']
  );
  const journal = {
    record: freshJournal().record,
    sheet: {
      getRange: () => ({
        setValues: values => saved.push(values[0]),
      }),
    },
  };
  journal.record.row_number = 2;
  sandbox.formSubmitSendLateRegistration_(
    journal, '第1回大会A級', 'A', ['回答'], new Date()
  );
  sandbox.formSubmitSendLateRegistration_(
    journal, '第1回大会A級', 'A', ['回答'], new Date()
  );
  assert.strictEqual(drafts, 1);
  assert.strictEqual(sends, 1);
  assert.strictEqual(journal.record['追加申込通知状態'], 'sent');
  assert.strictEqual(journal.record['Gmail message ID'], 'message-1');
  assert(saved.length >= 4);

  const recoveryJournal = {
    record: Object.assign(freshJournal().record, {
      row_number: 3,
      '追加申込通知状態': 'sending',
      '通知トークン': 'TAIKAI_FORM_EVENT_event-1',
      'Gmail draft ID': 'draft-1',
    }),
    sheet: journal.sheet,
  };
  sandbox.GmailApp.search = () => [{
    getMessages: () => [message],
  }];
  const recovered = sandbox.formSubmitSendLateRegistration_(
    recoveryJournal, '第1回大会A級', 'A', ['回答'], new Date()
  );
  assert.strictEqual(recovered.recovered, true);
  assert.strictEqual(recoveryJournal.record['追加申込通知状態'], 'sent');
  assert.strictEqual(sends, 1);
  const triggerDiagnosis = JSON.parse(sandbox.diagnoseFormSubmitTrigger());
  assert.strictEqual(triggerDiagnosis.ok, true);
  assert.strictEqual(triggerDiagnosis.matching_count, 1);
}

{
  const source = read('server/FormSubmitEffects.js');
  const handlerSource = read('server/FormSubmit.js');
  const documentation = read('FORM_SUBMIT_TRIGGER.md');
  assert.match(source, /ScriptApp\.getProjectTriggers\(\)/);
  assert.doesNotMatch(source, /ScriptApp\.newTrigger|ScriptApp\.deleteTrigger/);
  assert.match(documentation, /手動で1件だけ設定/);
  assert.match(documentation, /memoForm.*停止/);
  assert.match(handlerSource, /function retryFormResponseProcessing/);
  assert.match(documentation, /retryFormResponseProcessing/);
}

console.log('Unified form submit processing checks passed.');
