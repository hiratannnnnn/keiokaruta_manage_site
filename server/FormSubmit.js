// ============================================================
// Googleフォーム回答の新DB登録
// ============================================================

function formAnswerValue_(namedValues, predicate, label, required) {
  const matches = Object.keys(namedValues || {}).filter(predicate);
  if (matches.length > 1) {
    throw new Error(
      'フォーム回答の「' + label + '」を一意に特定できません'
      + '（候補' + matches.length + '件）。'
    );
  }
  if (!matches.length) {
    if (required) throw new Error('フォーム回答に「' + label + '」がありません。');
    return '';
  }
  return String((namedValues[matches[0]] || [''])[0] || '').trim();
}

function formDateValue_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/(\d{4})[年\/-](\d{1,2})[月\/-](\d{1,2})日?/);
  if (!match) return '';
  return match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
}

function recordFormResponseInTournamentSheetV2_(
  sheet, structure, sourceRow, response, apiResult, apiError
) {
  if (!structure || structure.version !== 2) return;
  const existingManagement =
    structure.management.entries_by_source_row[sourceRow];
  const responseColumns = tournamentSheetResponseColumns_(structure);
  const fee = tournamentSheetGradeFee_(structure, response.grade);
  const player = apiResult && apiResult.player ? apiResult.player : {};
  const entry = apiResult && apiResult.entry ? apiResult.entry : {};
  const summary = apiResult && apiResult.payment_summary
    ? apiResult.payment_summary : {};
  const participationFee = summary.participation_fee_yen !== undefined
    ? summary.participation_fee_yen : fee;
  const paid = summary.paid_yen !== undefined ? summary.paid_yen : 0;
  const balance = summary.balance_yen !== undefined
    ? summary.balance_yen
    : (participationFee === null ? '' : Number(participationFee) - Number(paid));
  const row = new Array(TOURNAMENT_SHEET_V2_WIDTH_).fill('');
  [
    '申込', sourceRow, response.email, response.name, response.grade, '',
    player.id || '', entry.id || '', entry.schedule_id || '', entry.canceled_at || '',
    participationFee === null ? '' : participationFee,
    paid,
    balance,
    summary.status || (participationFee === null ? 'unpriced' : 'unpaid'),
    apiError ? 'pending_api' : 'synced',
    apiError ? '' : new Date(),
    apiError ? String(apiError.message || apiError).slice(0, 5000) : '',
  ].forEach((value, index) => { row[index] = value; });

  if (existingManagement) {
    const existingEntryId = String(existingManagement.row[7] || '');
    const expectedEntryId = String(entry.id || '');
    if (apiError && existingEntryId) {
      throw new Error(
        '回答行' + sourceRow
        + 'は同期済みのため、API失敗値で管理行を上書きしません。'
      );
    }
    if (existingEntryId && expectedEntryId
        && existingEntryId !== expectedEntryId) {
      throw new Error(
        '回答行' + sourceRow + 'の申込管理行が別のentry IDを保持しています。'
      );
    }
    if (!apiError && !expectedEntryId) {
      throw new Error('回答行' + sourceRow + 'のentry IDが取得できません。');
    }
    // pending_api行は同じ行を成功値で更新する。成功済みの同一entryも冪等に上書きする。
    const dbValues = row.slice(6, 17);
    sheet.getRange(existingManagement.row_number, 7, 1, 11)
      .setValues([dbValues]);
    const refreshed = tournamentSheetStructure_(sheet, false);
    const metadataRows = refreshed.management.metadata_rows;
    if (apiError) {
      sheet.getRange(metadataRows['同期状態'], 3).setValue('pending_api');
      sheet.getRange(metadataRows['同期エラー'], 3)
        .setValue(String(apiError.message || apiError).slice(0, 5000));
    } else {
      const hasPendingEntry = Object.keys(
        refreshed.management.entries_by_source_row
      ).some(key => {
        const status = String(
          refreshed.management.entries_by_source_row[key].row[14] || ''
        );
        return status === 'pending_api' || status === 'pending_sheet';
      });
      if (!hasPendingEntry) {
        sheet.getRange(metadataRows['同期状態'], 3).setValue('synced');
        sheet.getRange(metadataRows['同期エラー'], 3).setValue('');
      }
      sheet.getRange(metadataRows['最終同期日時'], 3).setValue(new Date());
    }
    return;
  }

  let insertRow = null;
  for (let index = structure.management.start_index + 1;
       index < structure.management.end_index; index++) {
    if (String((structure.data[index] || [])[0] || '') === '[案内]') {
      insertRow = index + 1;
      break;
    }
  }
  if (!insertRow) throw new Error('大会管理データv2の案内セクションがありません。');
  const normalizedEmail = String(response.email || '').trim().toLowerCase();
  const previousEntries = [];
  for (let previousRow = 2; previousRow < sourceRow; previousRow++) {
    const raw = structure.data[previousRow - 1] || [];
    if (String(raw[responseColumns.email] || '').trim().toLowerCase()
        !== normalizedEmail) continue;
    const previousEntry = structure.management.entries_by_source_row[previousRow];
    if (!previousEntry) continue;
    previousEntries.push({
      row_number: previousEntry.row_number,
      values: sheet.getRange(previousEntry.row_number, 7, 1, 11).getValues(),
    });
  }

  let inserted = false;
  const metadataSnapshots = [];
  try {
    sheet.insertRowBefore(insertRow);
    inserted = true;
    sheet.getRange(insertRow, 1, 1, TOURNAMENT_SHEET_V2_WIDTH_).setValues([row]);

    previousEntries.forEach(previous => {
      const superseded = previous.values[0].slice();
      for (let index = 0; index < 7; index++) superseded[index] = '';
      superseded[7] = 'superseded';
      superseded[8] = 'superseded';
      superseded[9] = '';
      superseded[10] = '';
      sheet.getRange(previous.row_number, 7, 1, 11).setValues([superseded]);
    });

    const refreshed = tournamentSheetStructure_(sheet, false);
    const metadataRows = refreshed.management.metadata_rows;
    const metadataKeys = apiError
      ? ['同期状態', '同期エラー']
      : ['最終同期日時'];
    metadataKeys.forEach(key => {
      const rowNumber = metadataRows[key];
      metadataSnapshots.push({
        row_number: rowNumber,
        value: sheet.getRange(rowNumber, 3).getValue(),
      });
    });
    if (apiError) {
      sheet.getRange(metadataRows['同期状態'], 3).setValue('pending_api');
      sheet.getRange(metadataRows['同期エラー'], 3)
        .setValue(String(apiError.message || apiError).slice(0, 5000));
    } else {
      sheet.getRange(metadataRows['最終同期日時'], 3).setValue(new Date());
    }
  } catch (error) {
    const rollbackErrors = [];
    previousEntries.forEach(previous => {
      try {
        sheet.getRange(previous.row_number, 7, 1, 11).setValues(previous.values);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError.message || String(rollbackError));
      }
    });
    metadataSnapshots.forEach(snapshot => {
      try {
        sheet.getRange(snapshot.row_number, 3).setValue(snapshot.value);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError.message || String(rollbackError));
      }
    });
    if (inserted) {
      try {
        sheet.deleteRow(insertRow);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError.message || String(rollbackError));
      }
    }
    throw new Error(
      String(error.message || error)
      + (rollbackErrors.length
        ? ' ロールバックにも失敗しました: ' + rollbackErrors.join(' / ')
        : '')
    );
  }
}

function refreshSiblingTournamentSheetsV2AfterResponse_(sourceSheet) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const calendar = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
  if (!calendar || calendar.getLastRow() < 3) return [];
  const sourceName = sourceSheet.getName();
  const baseName = tournamentSheetBaseName_(sourceName);
  const names = calendar.getRange(
    3, 1, calendar.getLastRow() - 2, 1
  ).getValues().map(row => String(row[0] || '').trim());
  const seen = {};
  const results = [];
  const errors = [];
  names.forEach(name => {
    if (!name || name === sourceName || seen[name]
        || tournamentSheetBaseName_(name) !== baseName
        || !/[A-E]+級$/.test(name)) {
      return;
    }
    seen[name] = true;
    const sibling = ss.getSheetByName(name);
    if (!sibling) {
      errors.push(name + ': 大会シートが見つかりません。');
      return;
    }
    try {
      const structure = tournamentSheetStructure_(sibling, false);
      if (structure.version === 2) {
        results.push(refreshTournamentSheetV2FromApi_(sibling));
      }
    } catch (error) {
      try {
        markTournamentSheetV2SyncState_(
          sibling, 'pending_sheet', error.message || error
        );
      } catch (markError) {}
      errors.push(name + ': ' + error.message);
    }
  });
  if (errors.length) {
    throw new Error(
      '兄弟フォームへの書戻しに失敗しました: ' + errors.join(' / ')
    );
  }
  return results;
}

function formSubmitRegistrationResultFromJournal_(journal) {
  const playerId = String(journal.record['player ID'] || '');
  const entryId = String(journal.record['entry ID'] || '');
  if (!playerId || !entryId) {
    throw new Error('フォーム送信処理ジャーナルのDB識別情報が不足しています。');
  }
  return {
    player: taikaiApiRequest_(
      'GET', '/players/' + encodeURIComponent(playerId)
    ),
    entry: taikaiApiRequest_(
      'GET', '/entries/' + encodeURIComponent(entryId)
    ),
    payment_summary: taikaiApiRequest_(
      'GET',
      '/entries/' + encodeURIComponent(entryId) + '/payment-summary'
    ),
  };
}

function registerFormResponseToDatabaseUnlocked_(e) {
  if (!e || !e.range) throw new Error('フォーム回答イベントがありません。');
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const sourceRow = e.range.getRow();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (String(sheet.getParent().getId()) !== String(ss.getId())) {
    throw new Error('設定外のスプレッドシートからのフォーム回答は処理しません。');
  }
  const eventKey = formSubmitEventKey_(sheet, sourceRow);
  const journal = formSubmitJournalLoadOrCreate_(
    ss, eventKey, sheetName, sourceRow
  );
  const match = sheetName.match(/^(.*?)([A-E]+)級$/);
  if (!match) {
    const message = '回答シート名から大会名を取得できません: ' + sheetName;
    formSubmitJournalSave_(journal, {
      'DB状態': 'error',
      'v2書戻し状態': 'error',
      '名簿状態': 'error',
      '追加申込通知状態': 'error',
      'DBエラー': message,
      'v2書戻しエラー': message,
      '名簿エラー': message,
      '追加申込通知エラー': message,
    });
    throw new Error(message);
  }
  const alreadyComplete =
    String(journal.record['DB状態']) === 'done'
    && ['done', 'not_applicable'].includes(
      String(journal.record['v2書戻し状態'])
    )
    && String(journal.record['名簿状態']) === 'done'
    && ['sent', 'not_required'].includes(
      String(journal.record['追加申込通知状態'])
    );
  if (alreadyComplete) {
    return {
      ok: true,
      event_key: eventKey,
      player_id: String(journal.record['player ID'] || ''),
      entry_id: String(journal.record['entry ID'] || ''),
      roster_status: 'done',
      notification_status: String(
        journal.record['追加申込通知状態'] || ''
      ),
    };
  }

  let response;
  try {
    const named = e.namedValues || {};
    const name = formAnswerValue_(
      named, title => title.includes('氏名'), '氏名', true
    );
    const email = formAnswerValue_(
      named, title => /メールアドレス/i.test(title), 'メールアドレス', true
    );
    const grade = formAnswerValue_(
      named,
      title => title === '級' || /参加.*級|出場.*級/.test(title),
      '級',
      true
    ).replace(/級/g, '').trim();
    const heldOn = formDateValue_(
      formAnswerValue_(named, title => title.includes('希望日'), '希望日', false)
    );
    if (!name || !email || !/^[A-E]$/.test(grade)) {
      throw new Error('回答から氏名・メールアドレス・級を取得できません。');
    }
    response = {
      name: formSubmitNormalizeName_(name),
      email: email,
      grade: grade,
      held_on: heldOn,
    };
  } catch (error) {
    const message = String(error.message || error);
    formSubmitJournalSave_(journal, {
      'DB状態': 'error',
      'v2書戻し状態': 'error',
      '名簿状態': 'error',
      '追加申込通知状態': 'error',
      'DBエラー': message,
      'v2書戻しエラー': message,
      '名簿エラー': message,
      '追加申込通知エラー': message,
    });
    throw error;
  }

  const errors = [];
  let structure = null;
  let structureError = null;
  if (String(journal.record['v2書戻し状態']) !== 'done'
      && String(journal.record['v2書戻し状態']) !== 'not_applicable') {
    try {
      structure = tournamentSheetStructure_(sheet, false);
    } catch (error) {
      structureError = error;
      formSubmitJournalSave_(journal, {
        'v2書戻し状態': 'error',
        'v2書戻しエラー': String(error.message || error),
      });
      errors.push('v2構造確認: ' + String(error.message || error));
    }
  }

  let result = null;
  let apiError = null;
  if (String(journal.record['DB状態']) === 'done') {
    if (!structureError
        && String(journal.record['v2書戻し状態']) !== 'done'
        && String(journal.record['v2書戻し状態']) !== 'not_applicable') {
      try {
        result = formSubmitRegistrationResultFromJournal_(journal);
      } catch (error) {
        apiError = error;
        errors.push('DB結果の再取得: ' + String(error.message || error));
      }
    }
  } else {
    try {
      result = taikaiRegisterEntry_(
        match[1], response.grade, response.held_on,
        response.name, response.email
      );
      if (!result || !result.player || !result.player.id
          || !result.entry || !result.entry.id) {
        throw new Error('大会申込APIの識別情報が不足しています。');
      }
      result.payment_summary = taikaiApiRequest_(
        'GET',
        '/entries/' + encodeURIComponent(String(result.entry.id))
          + '/payment-summary'
      );
      formSubmitJournalSave_(journal, {
        'DB状態': 'done',
        'player ID': String(result.player.id),
        'entry ID': String(result.entry.id),
        'schedule ID': String(result.entry.schedule_id || ''),
        'DBエラー': '',
      });
    } catch (error) {
      apiError = error;
      formSubmitJournalSave_(journal, {
        'DB状態': 'error',
        'DBエラー': String(error.message || error),
      });
      errors.push('DB登録: ' + String(error.message || error));
    }
  }

  const v2Status = String(journal.record['v2書戻し状態']);
  if (!structureError
      && v2Status !== 'done' && v2Status !== 'not_applicable') {
    if (apiError && String(journal.record['DB状態']) === 'done') {
      formSubmitJournalSave_(journal, {
        'v2書戻し状態': 'error',
        'v2書戻しエラー':
          'DB登録結果を再取得できないため、既存のv2管理行を変更していません。',
      });
    } else {
      try {
        recordFormResponseInTournamentSheetV2_(
          sheet, structure, sourceRow, response, result, apiError
        );
        if (structure.version !== 2) {
          formSubmitJournalSave_(journal, {
            'v2書戻し状態': 'not_applicable',
            'v2書戻しエラー': '',
          });
        } else if (apiError) {
          formSubmitJournalSave_(journal, {
            'v2書戻し状態': 'waiting_for_db',
            'v2書戻しエラー': '',
          });
        } else {
          refreshSiblingTournamentSheetsV2AfterResponse_(sheet);
          formSubmitJournalSave_(journal, {
            'v2書戻し状態': 'done',
            'v2書戻しエラー': '',
          });
        }
      } catch (error) {
        formSubmitJournalSave_(journal, {
          'v2書戻し状態': 'error',
          'v2書戻しエラー': String(error.message || error),
        });
        errors.push('v2書戻し: ' + String(error.message || error));
      }
    }
  }

  if (String(journal.record['名簿状態']) !== 'done') {
    try {
      formSubmitMaintainRoster_(ss, response.name);
      formSubmitJournalSave_(journal, {
        '名簿状態': 'done',
        '名簿エラー': '',
      });
    } catch (error) {
      formSubmitJournalSave_(journal, {
        '名簿状態': 'error',
        '名簿エラー': String(error.message || error),
      });
      errors.push('名簿更新: ' + String(error.message || error));
    }
  }

  const notificationStatus = String(journal.record['追加申込通知状態']);
  if (notificationStatus !== 'sent' && notificationStatus !== 'not_required') {
    try {
      formSubmitSendLateRegistration_(
        journal, sheetName, response.grade, e.values || [], new Date()
      );
    } catch (error) {
      if (String(journal.record['追加申込通知状態']) === 'pending'
          || String(journal.record['追加申込通知状態']) === 'error') {
        formSubmitJournalSave_(journal, {
          '追加申込通知状態': 'error',
          '追加申込通知エラー': String(error.message || error),
        });
      } else {
        formSubmitJournalSave_(journal, {
          '追加申込通知エラー': String(error.message || error),
        });
      }
      errors.push('追加申込通知: ' + String(error.message || error));
    }
  }

  if (errors.length) {
    throw new Error(
      'フォーム送信後処理が一部未完了です。再実行できます: '
      + errors.join(' / ')
    );
  }
  return {
    ok: true,
    event_key: eventKey,
    player_id: String(journal.record['player ID'] || ''),
    entry_id: String(journal.record['entry ID'] || ''),
    roster_status: String(journal.record['名簿状態'] || ''),
    notification_status: String(journal.record['追加申込通知状態'] || ''),
  };
}

function registerFormResponseToDatabase(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return registerFormResponseToDatabaseUnlocked_(e);
  } finally {
    lock.releaseLock();
  }
}

function retryFormResponseProcessing(sheetName, sourceRow) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(String(sheetName || ''));
  if (!sheet || !/[A-E]+級$/.test(sheet.getName())) {
    throw new Error('再処理する大会回答シートが見つかりません。');
  }
  const rowNumber = Number(sourceRow);
  const structure = tournamentSheetStructure_(sheet, false);
  if (!Number.isInteger(rowNumber) || rowNumber < 2
      || rowNumber >= structure.response_end_index + 1) {
    throw new Error('再処理するフォーム回答行が不正です。');
  }
  const width = tournamentSheetRawResponseColumnCount_(structure);
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, width).getValues()[0];
  const namedValues = {};
  headers.forEach((header, index) => {
    const key = String(header || '').trim();
    if (!key) return;
    if (Object.prototype.hasOwnProperty.call(namedValues, key)) {
      throw new Error('フォーム回答見出しが重複しています: ' + key);
    }
    namedValues[key] = [values[index]];
  });
  return registerFormResponseToDatabase({
    range: sheet.getRange(rowNumber, 1, 1, width),
    values: values,
    namedValues: namedValues,
  });
}
