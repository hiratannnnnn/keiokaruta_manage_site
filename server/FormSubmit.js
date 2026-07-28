// ============================================================
// Googleフォーム回答の新DB登録
// ============================================================

function formAnswerValue_(namedValues, values, predicate, fallbackIndex) {
  const keys = Object.keys(namedValues || {});
  for (let i = 0; i < keys.length; i++) {
    if (predicate(keys[i])) return String((namedValues[keys[i]] || [''])[0] || '').trim();
  }
  return String((values || [])[fallbackIndex] || '').trim();
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
  if (structure.management.entries_by_source_row[sourceRow]) {
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
  const fee = tournamentSheetGradeFee_(structure, response.grade);
  const player = apiResult && apiResult.player ? apiResult.player : {};
  const entry = apiResult && apiResult.entry ? apiResult.entry : {};
  sheet.insertRowBefore(insertRow);
  const row = new Array(TOURNAMENT_SHEET_V2_WIDTH_).fill('');
  [
    '申込', sourceRow, response.email, response.name, response.grade, '',
    player.id || '', entry.id || '', entry.schedule_id || '', entry.canceled_at || '',
    fee === null ? '' : fee, 0, fee === null ? '' : fee,
    fee === null ? 'unpriced' : 'unpaid',
    apiError ? 'pending_api' : 'synced',
    apiError ? '' : new Date(),
    apiError ? String(apiError.message || apiError).slice(0, 5000) : '',
  ].forEach((value, index) => { row[index] = value; });
  sheet.getRange(insertRow, 1, 1, TOURNAMENT_SHEET_V2_WIDTH_).setValues([row]);

  const metadataRows = structure.management.metadata_rows;
  if (apiError) {
    sheet.getRange(metadataRows['同期状態'], 3).setValue('pending_api');
    sheet.getRange(metadataRows['同期エラー'], 3)
      .setValue(String(apiError.message || apiError).slice(0, 5000));
  } else {
    sheet.getRange(metadataRows['最終同期日時'], 3).setValue(new Date());
  }
}

function registerFormResponseToDatabase(e) {
  if (!e || !e.range) throw new Error('フォーム回答イベントがありません。');
  const sheetName = e.range.getSheet().getName();
  const match = sheetName.match(/^(.*?)([A-E]+)級$/);
  if (!match) throw new Error('回答シート名から大会名を取得できません: ' + sheetName);

  const named = e.namedValues || {};
  const values = e.values || [];
  const name = formAnswerValue_(named, values, title => title.includes('氏名'), 2);
  const email = formAnswerValue_(named, values, title => /メールアドレス/i.test(title), 1);
  const grade = formAnswerValue_(named, values, title => title === '級' || title.endsWith('級'), 4)
    .replace(/級/g, '').trim();
  const heldOn = formDateValue_(formAnswerValue_(named, values, title => title.includes('希望日'), -1));
  if (!name || !email || !/^[A-E]$/.test(grade)) {
    throw new Error('回答から氏名・メールアドレス・級を取得できません。');
  }

  const sheet = e.range.getSheet();
  const structure = tournamentSheetStructure_(sheet, false);
  let result = null;
  let apiError = null;
  try {
    result = taikaiRegisterEntry_(match[1], grade, heldOn, name, email);
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
  return apiError
    ? { pending: true, error: apiError.message }
    : result;
}
