// ============================================================
// フォーム送信後処理の永続ジャーナル・名簿・追加申込通知
// ============================================================

const FORM_SUBMIT_JOURNAL_HEADERS_ = [
  'イベントキー', '回答シート', '回答行',
  'DB状態', 'v2書戻し状態', '名簿状態', '追加申込通知状態',
  'player ID', 'entry ID', 'schedule ID',
  '通知トークン', 'Gmail draft ID', 'Gmail message ID',
  'DBエラー', 'v2書戻しエラー', '名簿エラー', '追加申込通知エラー',
  '更新日時',
];

function formSubmitDigestHex_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  ).map(byte => {
    const unsigned = byte < 0 ? byte + 256 : byte;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function formSubmitEventKey_(sheet, sourceRow) {
  return formSubmitDigestHex_([
    sheet.getParent().getId(),
    sheet.getSheetId(),
    Number(sourceRow),
  ].join('|'));
}

function formSubmitJournalSheet_(spreadsheet) {
  const name = CONFIG.SHEET_NAMES.FORM_SUBMIT_JOURNAL;
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.getRange(1, 1, 1, FORM_SUBMIT_JOURNAL_HEADERS_.length)
      .setValues([FORM_SUBMIT_JOURNAL_HEADERS_]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
    return sheet;
  }
  const actual = sheet.getRange(
    1, 1, 1, FORM_SUBMIT_JOURNAL_HEADERS_.length
  ).getValues()[0].map(String);
  if (actual.join('\n') !== FORM_SUBMIT_JOURNAL_HEADERS_.join('\n')) {
    throw new Error(
      'フォーム送信処理ジャーナルの見出しが不正です。'
      + '手動で列を変更しないでください。'
    );
  }
  return sheet;
}

function formSubmitJournalObject_(rowNumber, values) {
  const record = { row_number: rowNumber };
  FORM_SUBMIT_JOURNAL_HEADERS_.forEach((header, index) => {
    record[header] = (values || [])[index] === undefined
      ? '' : (values || [])[index];
  });
  return record;
}

function formSubmitJournalLoadOrCreate_(
  spreadsheet, eventKey, sheetName, sourceRow
) {
  const sheet = formSubmitJournalSheet_(spreadsheet);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const matches = [];
    keys.forEach((row, index) => {
      if (String(row[0] || '') === eventKey) matches.push(index + 2);
    });
    if (matches.length > 1) {
      throw new Error('フォーム送信処理ジャーナルのイベントキーが重複しています。');
    }
    if (matches.length === 1) {
      const values = sheet.getRange(
        matches[0], 1, 1, FORM_SUBMIT_JOURNAL_HEADERS_.length
      ).getValues()[0];
      if (String(values[1] || '') !== sheetName
          || Number(values[2]) !== Number(sourceRow)) {
        throw new Error('フォーム送信処理ジャーナルの回答識別情報が競合しています。');
      }
      return { sheet: sheet, record: formSubmitJournalObject_(matches[0], values) };
    }
  }
  const rowNumber = lastRow + 1;
  const values = new Array(FORM_SUBMIT_JOURNAL_HEADERS_.length).fill('');
  values[0] = eventKey;
  values[1] = sheetName;
  values[2] = Number(sourceRow);
  values[3] = 'pending';
  values[4] = 'pending';
  values[5] = 'pending';
  values[6] = 'pending';
  values[17] = new Date();
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  return { sheet: sheet, record: formSubmitJournalObject_(rowNumber, values) };
}

function formSubmitJournalSave_(journal, changes) {
  Object.keys(changes || {}).forEach(key => {
    if (!FORM_SUBMIT_JOURNAL_HEADERS_.includes(key)) {
      throw new Error('フォーム送信処理ジャーナルの更新項目が不正です: ' + key);
    }
    journal.record[key] = changes[key];
  });
  journal.record['更新日時'] = new Date();
  const values = FORM_SUBMIT_JOURNAL_HEADERS_.map(
    header => journal.record[header] === undefined ? '' : journal.record[header]
  );
  journal.sheet.getRange(
    journal.record.row_number, 1, 1, values.length
  ).setValues([values]);
  return journal.record;
}

function formSubmitNormalizeName_(value) {
  return String(value || '').trim().split(/[ 　]+/).filter(Boolean).join(' ');
}

function formSubmitMaintainRoster_(spreadsheet, name) {
  const normalized = formSubmitNormalizeName_(name);
  if (!normalized) throw new Error('名簿へ追加する氏名が空です。');
  const lastName = normalized.split(' ')[0];
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.MEMBERS);
  if (!sheet) throw new Error('名簿シートが見つかりません。');
  const roster = sheet.getDataRange().getValues();
  let sameSurnameIndex = null;
  for (let index = 1; index < roster.length; index++) {
    if (formSubmitNormalizeName_(roster[index][3]) === normalized) {
      return { changed: false, row: index + 1 };
    }
    if (String(roster[index][1] || '').trim() === lastName) {
      sameSurnameIndex = index;
    }
  }
  if (sameSurnameIndex !== null) {
    const rowNumber = sameSurnameIndex + 2;
    sheet.insertRowAfter(sameSurnameIndex + 1);
    sheet.getRange(rowNumber, 1, 1, 4)
      .setValues([['', lastName, '重複', normalized]]);
    return { changed: true, row: rowNumber, duplicate_surname: true };
  }
  const rowNumber = sheet.getLastRow() + 1;
  sheet.insertRowAfter(sheet.getLastRow());
  sheet.getRange(rowNumber, 1, 1, 4)
    .setValues([['', lastName, '', normalized]]);
  return { changed: true, row: rowNumber, duplicate_surname: false };
}

function formSubmitLateRegistrationTarget_(sheetName, grade, now) {
  const tournamentName = tournamentSheetBaseName_(sheetName);
  const tournament = taikaiFindTournament_(tournamentName);
  if (!Boolean(Number(tournament.registration_completed))) return null;
  const schedules = taikaiApiRequest_(
    'GET',
    '/tournaments/' + encodeURIComponent(String(tournament.id)) + '/schedules'
  ) || [];
  const targets = schedules.filter(schedule =>
    String(schedule.grade || '').toUpperCase() === String(grade || '').toUpperCase()
  );
  if (targets.length !== 1) {
    throw new Error(
      '追加申込通知の日程を一意に特定できません: '
      + tournamentName + ' ' + grade + '級'
    );
  }
  const today = Utilities.formatDate(now || new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(targets[0].application_deadline || '') >= today
    ? { tournament: tournament, schedule: targets[0] } : null;
}

function formSubmitSentMessageByToken_(token) {
  const threads = GmailApp.search('in:sent "' + token + '"', 0, 20) || [];
  for (let threadIndex = 0; threadIndex < threads.length; threadIndex++) {
    const messages = threads[threadIndex].getMessages() || [];
    for (let messageIndex = 0; messageIndex < messages.length; messageIndex++) {
      const message = messages[messageIndex];
      if (String(message.getPlainBody() || '').includes(token)) return message;
    }
  }
  return null;
}

function formSubmitNotificationRecipient_() {
  const recipients = configValue_('FORM_RESPONSE_NOTIFICATION_TO')
    .split(',').map(value => value.trim()).filter(Boolean);
  if (!recipients.length || recipients.some(value =>
    !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(value)
  )) {
    throw new Error('FORM_RESPONSE_NOTIFICATION_TO のメール形式が不正です。');
  }
  return recipients.join(',');
}

function formSubmitSendLateRegistration_(
  journal, sheetName, grade, eventValues, now
) {
  const currentStatus = String(
    journal.record['追加申込通知状態'] || ''
  );
  if (currentStatus === 'sent' || currentStatus === 'not_required') {
    return { status: currentStatus, skipped: true };
  }
  const target = formSubmitLateRegistrationTarget_(sheetName, grade, now);
  if (!target) {
    formSubmitJournalSave_(journal, {
      '追加申込通知状態': 'not_required',
      '追加申込通知エラー': '',
    });
    return { status: 'not_required', skipped: true };
  }

  const token = String(journal.record['通知トークン'] || '')
    || 'TAIKAI_FORM_EVENT_' + String(journal.record['イベントキー']);
  formSubmitJournalSave_(journal, { '通知トークン': token });
  const sent = formSubmitSentMessageByToken_(token);
  if (sent) {
    formSubmitJournalSave_(journal, {
      '追加申込通知状態': 'sent',
      'Gmail message ID': sent.getId(),
      '追加申込通知エラー': '',
    });
    return { status: 'sent', recovered: true, message_id: sent.getId() };
  }
  if (currentStatus === 'sending') {
    formSubmitJournalSave_(journal, {
      '追加申込通知状態': 'delivery_unknown',
      '追加申込通知エラー':
        '送信開始後の結果を確認できません。トークンで送信済みメールを再確認してください。',
    });
    throw new Error('追加申込通知の送信結果を安全に確認できません。');
  }
  if (currentStatus === 'delivery_unknown') {
    throw new Error('追加申込通知の送信結果が未確認です。');
  }

  let draft = null;
  const draftId = String(journal.record['Gmail draft ID'] || '');
  if (draftId) {
    try {
      draft = GmailApp.getDraft(draftId);
    } catch (error) {
      if (currentStatus === 'draft_ready') {
        formSubmitJournalSave_(journal, {
          '追加申込通知状態': 'pending',
          'Gmail draft ID': '',
          '追加申込通知エラー': '',
        });
      } else {
        formSubmitJournalSave_(journal, {
          '追加申込通知状態': 'delivery_unknown',
          '追加申込通知エラー':
            '保存済み下書きが見つからず、送信済みメールも確認できません。',
        });
        throw new Error('追加申込通知の保存済み下書きを確認できません。');
      }
    }
  }
  if (!draft) {
    const recipient = formSubmitNotificationRecipient_();
    const subject = sheetName + '　追加申込';
    const body = (eventValues || []).map(value => String(value || '')).join('\n')
      + '\n\n' + token;
    draft = GmailApp.createDraft(recipient, subject, body, {
      name: '慶應かるた会',
    });
    formSubmitJournalSave_(journal, {
      '追加申込通知状態': 'draft_ready',
      'Gmail draft ID': draft.getId(),
      '追加申込通知エラー': '',
    });
  }
  formSubmitJournalSave_(journal, {
    '追加申込通知状態': 'sending',
    '追加申込通知エラー': '',
  });
  const message = draft.send();
  formSubmitJournalSave_(journal, {
    '追加申込通知状態': 'sent',
    'Gmail message ID': message.getId(),
    '追加申込通知エラー': '',
  });
  return { status: 'sent', message_id: message.getId() };
}

function diagnoseFormSubmitTrigger() {
  try {
    const spreadsheetId = CONFIG.SPREADSHEET_ID;
    const triggers = ScriptApp.getProjectTriggers().filter(trigger =>
      trigger.getHandlerFunction() === 'registerFormResponseToDatabase'
    ).map(trigger => ({
      handler: trigger.getHandlerFunction(),
      source: String(trigger.getTriggerSource()),
      event_type: String(trigger.getEventType()),
      source_id: typeof trigger.getTriggerSourceId === 'function'
        ? String(trigger.getTriggerSourceId() || '') : '',
    }));
    const matching = triggers.filter(trigger =>
      trigger.source_id === spreadsheetId
      && trigger.event_type === String(ScriptApp.EventType.ON_FORM_SUBMIT)
    );
    return JSON.stringify({
      ok: matching.length === 1 && triggers.length === 1,
      expected_handler: 'registerFormResponseToDatabase',
      expected_spreadsheet_id: spreadsheetId,
      matching_count: matching.length,
      handler_count: triggers.length,
      triggers: triggers,
      warning: matching.length === 1 && triggers.length === 1
        ? '' : 'フォーム送信トリガーは対象スプレッドシートに1件だけ手動設定してください。',
    });
  } catch (error) {
    return JSON.stringify({ error: error.message });
  }
}
