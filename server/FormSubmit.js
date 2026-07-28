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
  if (existingManagement) {
    const expectedEntryId = apiResult && apiResult.entry
      ? String(apiResult.entry.id || '') : '';
    if (!apiError && expectedEntryId
        && String(existingManagement.row[7] || '') === expectedEntryId) {
      return;
    }
    throw new Error('回答行' + sourceRow + 'の申込管理行が既に存在します。');
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
  const responseColumns = tournamentSheetResponseColumns_(structure);
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

function registerFormResponseToDatabaseUnlocked_(e) {
  if (!e || !e.range) throw new Error('フォーム回答イベントがありません。');
  const sheetName = e.range.getSheet().getName();
  const match = sheetName.match(/^(.*?)([A-E]+)級$/);
  if (!match) throw new Error('回答シート名から大会名を取得できません: ' + sheetName);

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
  )
    .replace(/級/g, '').trim();
  const heldOn = formDateValue_(
    formAnswerValue_(named, title => title.includes('希望日'), '希望日', false)
  );
  if (!name || !email || !/^[A-E]$/.test(grade)) {
    throw new Error('回答から氏名・メールアドレス・級を取得できません。');
  }

  const sheet = e.range.getSheet();
  const structure = tournamentSheetStructure_(sheet, false);
  let result = null;
  let apiError = null;
  try {
    result = taikaiRegisterEntry_(match[1], grade, heldOn, name, email);
    if (result && result.entry && result.entry.id) {
      result.payment_summary = taikaiApiRequest_(
        'GET',
        '/entries/' + encodeURIComponent(String(result.entry.id)) + '/payment-summary'
      );
    }
  } catch (error) {
    apiError = error;
  }
  recordFormResponseInTournamentSheetV2_(
    sheet,
    structure,
    e.range.getRow(),
    { name: name, email: email, grade: grade },
    result,
    apiError
  );
  if (apiError && structure.version !== 2) throw apiError;
  if (!apiError) {
    refreshSiblingTournamentSheetsV2AfterResponse_(sheet);
  }
  return apiError
    ? { pending: true, error: apiError.message }
    : result;
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
