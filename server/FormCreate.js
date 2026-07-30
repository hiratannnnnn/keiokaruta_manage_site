// ============================================================
// フォーム作成（createFormAndSheet の Web App 版）
// ============================================================

// フォームに質問を追加し、フォームカラム数（N）を返す
// questionsData : [[name, inc, req], ...]
// moshikomiStart: Date オブジェクト（公認大会出場回数タイトル計算用）
// grades        : 例 "ABCDE級"
function addQuestionsToForm(formId, questionsData, moshikomiStart, grades) {
  const TOURNAMENTS_QUESTION = '出場大会を全てお書きください。（略称等で構いません）';
  const form = FormApp.openById(formId);

  // テンプレートの既存項目をすべて削除
  form.getItems().forEach(item => form.deleteItem(item));

  const DANKAI   = ['無段', '初段', '二段', '三段', '四段', '五段', '六段', '七段', '八段'];
  const gradeSet = grades ? grades.replace('級', '').split('') : [];

  // 公認大会出場回数の動的タイトル計算
  const d  = moshikomiStart;
  const m  = d.getMonth();
  const y  = d.getFullYear();
  const fy = (m + 1) < 4 ? y - 1 : y;
  const kouninTitle =
    fy + '年度公認大会出場回数（' + fy + '年4月1日～' +
    y + '年' + (m + 1) + '月' + d.getDate() + '日）';

  let questionCount = 0;
  questionsData.forEach(([name, inc, req]) => {
    if (!inc) return;
    questionCount++;
    const isKounin    = name.includes('公認大会出場回数');
    const isTournamentList = name === TOURNAMENTS_QUESTION;
    const actualTitle = isKounin ? kouninTitle : name;

    if (name === '級') {
      const item = form.addMultipleChoiceItem();
      item.setTitle(actualTitle).setRequired(req === 1);
      if (gradeSet.length) item.setChoiceValues(gradeSet);
    } else if (name === '段位') {
      const item = form.addMultipleChoiceItem();
      item.setTitle(actualTitle).setRequired(req === 1).setChoiceValues(DANKAI);
    } else if (isTournamentList) {
      const item = form.addParagraphTextItem();
      item.setTitle(actualTitle).setRequired(true);
    } else {
      const item = form.addTextItem();
      item.setTitle(actualTitle).setRequired(req === 1);
      if (name.includes('氏名')) {
        const validation = FormApp.createTextValidation()
          .setHelpText('半角または全角スペースを含めてください。')
          .requireTextMatchesPattern('.*[ 　]+.*')
          .build();
        item.setHelpText('半角または全角スペースを含めてください。');
        item.setValidation(validation);
      } else if (isKounin) {
        item.setHelpText(
          '今年度の「申込開始日時点で、出場した、または出場することが分かっている公認大会」' +
          'の回数をお書きください。\n現在慶應かるた会botの出場回数確認は機能していません。自己管理をお願いします。'
        );
      }
    }
  });

  return questionCount; // タイムスタンプ・メールアドレス列は含まない
}

// カレンダーシートから大会行を検索し、行番号（1-indexed）を返す
// 見つからない場合は末尾に新規行を追加して返す
function findFromCalendar(calendarSheet, name) {
  const data = calendarSheet.getRange(1, 1, calendarSheet.getLastRow(), 1).getValues();
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][0]) === name) return i + 1;
  }
  const nextRow = calendarSheet.getLastRow() + 1;
  calendarSheet.getRange(nextRow, 1).setValue(name);
  return nextRow;
}

function formCreateRollbackAction_(result, target, action) {
  try {
    action();
    result.actions.push({ target: target, status: 'rolled_back' });
  } catch (error) {
    result.complete = false;
    result.actions.push({
      target: target,
      status: 'rollback_failed',
      error: String(error && error.message || error).slice(0, 500),
    });
  }
}

function formCreateCreatedResources_(state) {
  const resources = [];
  if (state.dbTournamentCreated) {
    resources.push('DB大会');
  } else if (state.dbTournamentCreationUncertain) {
    resources.push('DB大会（作成された可能性あり）');
  }
  if (state.formFileId) {
    resources.push('Googleフォーム');
  } else if (state.formCopyRequested) {
    resources.push('Googleフォーム（作成された可能性あり）');
  }
  if (state.responseSheetId) {
    resources.push('フォーム回答シート');
  } else if (state.destinationRequested) {
    resources.push('フォーム回答シート（作成された可能性あり）');
  }
  if (state.mailWritten) resources.push('メール管理行');
  if (state.calendarWritten) resources.push('カレンダー行');
  if (state.announceWritten) resources.push('案内メール作成設定');
  return resources;
}

function formCreateRollback_(state) {
  const result = {
    attempted: false,
    complete: true,
    actions: [],
    manual_action_required: false,
    unreverted_resources: [],
  };

  if (state.announceWritten && state.announceSheet) {
    result.attempted = true;
    formCreateRollbackAction_(result, '案内メール作成設定', () => {
      state.announceSnapshots.forEach(snapshot => {
        state.announceSheet.getRange(snapshot.a1).setValues(snapshot.values);
      });
    });
  }

  if (state.calendarWritten && state.calendarSheet && state.calendarMutation) {
    result.attempted = true;
    formCreateRollbackAction_(result, 'カレンダー行', () => {
      state.calendarMutation.snapshots.forEach(snapshot => {
        state.calendarSheet.getRange(
          state.calendarMutation.row, snapshot.column
        ).setValue(snapshot.value);
      });
    });
  }

  if (state.mailWritten && state.mailSheet && state.mailMutation) {
    result.attempted = true;
    formCreateRollbackAction_(result, 'メール管理行', () => {
      state.mailSheet.getRange(
        state.mailMutation.row, 1, 1, 6
      ).setValues(state.mailMutation.previousValues);
      state.mailSheet.getRange(2, 3).setValue(state.mailPointerBefore);
    });
  }

  if (state.destinationRequested && state.formId) {
    result.attempted = true;
    formCreateRollbackAction_(result, 'フォーム回答先の解除', () => {
      FormApp.openById(state.formId).removeDestination();
    });
  }

  if (state.destinationRequested && state.spreadsheet) {
    result.attempted = true;
    formCreateRollbackAction_(result, 'フォーム回答シート', () => {
      let candidates = [];
      if (state.responseSheetId) {
        const exact = state.spreadsheet.getSheetById(state.responseSheetId);
        if (exact) candidates = [exact];
      } else {
        candidates = state.spreadsheet.getSheets().filter(sheet =>
          !state.existingSheetIds[String(sheet.getSheetId())]
        );
      }
      if (candidates.length === 0) return;
      if (candidates.length !== 1) {
        throw new Error(
          '追加シートを一意に特定できません（候補'
          + candidates.length + '件）。'
        );
      }
      state.spreadsheet.deleteSheet(candidates[0]);
    });
  }

  if (state.formFileId) {
    result.attempted = true;
    formCreateRollbackAction_(result, 'Googleフォーム', () => {
      const file = DriveApp.getFileById(state.formFileId);
      if (!file.isTrashed()) file.setTrashed(true);
    });
  } else if (state.formCopyRequested) {
    result.attempted = true;
    result.complete = false;
    result.actions.push({
      target: 'Googleフォーム（作成確認）',
      status: 'rollback_failed',
      error: 'フォームIDを取得する前に失敗したため、自動確認・削除できません。',
    });
  }

  if (state.dbTournamentCreated && state.dbTournamentId) {
    result.attempted = true;
    formCreateRollbackAction_(result, 'DB大会', () => {
      taikaiApiRequest_(
        'DELETE',
        '/tournaments/' + encodeURIComponent(String(state.dbTournamentId))
      );
    });
  } else if (state.dbTournamentCreated || state.dbTournamentCreationUncertain) {
    result.attempted = true;
    result.complete = false;
    result.actions.push({
      target: 'DB大会（作成確認）',
      status: 'rollback_failed',
      error: '大会IDを取得できなかったため、自動確認・削除できません。',
    });
  }

  result.unreverted_resources = result.actions
    .filter(action => action.status === 'rollback_failed')
    .map(action => action.target);
  result.manual_action_required = result.unreverted_resources.length > 0;
  return result;
}

// Web App からフォームとシートを作成する
function createFormFromWeb(paramsJson) {
  const state = {
    phase: '入力検証',
    title: '',
    grades: '',
    sheetName: '',
    spreadsheet: null,
    calendarSheet: null,
    calendarMutation: null,
    calendarWritten: false,
    mailSheet: null,
    mailMutation: null,
    mailPointerBefore: '',
    mailWritten: false,
    announceSheet: null,
    announceSnapshots: [],
    announceWritten: false,
    dbTournamentCreated: false,
    dbTournamentCreationUncertain: false,
    dbTournamentId: '',
    formCopyRequested: false,
    formFileId: '',
    formId: '',
    formUrl: '',
    existingSheetIds: {},
    responseSheetId: '',
    destinationRequested: false,
  };
  const operationLock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = operationLock.tryLock(1000);
    if (!lockAcquired) {
      throw new Error(
        '別のフォーム作成処理が実行中です。完了後に再実行してください。'
      );
    }
    const p = JSON.parse(paramsJson);
    const { title, grades, questionsData,
            moshikomiStartStr, moshiDeadStr, raffleStr, huriDeadStr, isKoen } = p;
    state.title = String(title || '');
    state.grades = String(grades || '');
    const hasEditionNumber = /第\s*[0-9０-９一二三四五六七八九十百千〇零]+\s*回/.test(String(title || ''));
    if (!hasEditionNumber && p.allowNonUniqueTitle !== true) {
      throw new Error(
        '大会名に「第〇回」がありません。例外として作成する場合は、警告欄で明示的に確認してください。'
      );
    }

    const moshikomiStart = new Date(moshikomiStartStr);
    const moshiDead      = new Date(moshiDeadStr);
    const raffle         = raffleStr   ? new Date(raffleStr)   : null;
    const huriDead       = huriDeadStr ? new Date(huriDeadStr) : null;
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    state.spreadsheet = ss;
    const calendarSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    state.calendarSheet = calendarSheet;
    if (!calendarSheet) {
      throw new Error(
        'カレンダーシートが見つかりません: '
        + CONFIG.SHEET_NAMES.CALENDAR
      );
    }
    const sheetName = title + grades;
    state.sheetName = sheetName;
    if (ss.getSheetByName(sheetName)) {
      throw new Error('同名の大会フォームが既に存在します: ' + sheetName);
    }

    // 大会本体を先にDBへ登録する。ここで失敗した場合はGoogle側を作成しない。
    // 級別日程は開催日が確定するまで作成できないため、後のsaveTournamentDatesで登録する。
    state.phase = 'DBへの大会登録';
    const ensured = taikaiEnsureTournamentWithState_(title, {
      operation: 'フォーム作成',
      outcome: '失敗時はフォーム作成を中断',
    });
    state.dbTournamentCreated = ensured.created === true;
    state.dbTournamentId = String(ensured.tournament.id || '');

    const formTitle     = title + grades + '\u3000参加表明フォーム';

    // フォームをテンプレートからコピー
    state.phase = 'Googleフォームの作成';
    const originalFile = DriveApp.getFileById(CONFIG.FORM_TEMPLATE_ID);
    const folder       = DriveApp.getFolderById(CONFIG.FORM_FOLDER_TO);
    state.formCopyRequested = true;
    const newFormFile  = originalFile.makeCopy(formTitle, folder);
    state.formFileId = newFormFile.getId();
    const form         = FormApp.openById(newFormFile.getId());

    form.setDescription(
      'こちらは' + title + grades + 'の参加表明フォームです。\n' +
      '該当項目に回答の上、送信してください。\n' +
      '回答が正しく送信されている場合、入力いただいたメールアドレスに回答のコピーが届きますのでご確認ください。'
    );
    try { form.setPublished(true); } catch(e) {}  // 旧バージョン互換
    form.setAcceptingResponses(true);
    form.setCollectEmail(true);
    form.setLimitOneResponsePerUser(false);

    const formId  = form.getId();
    state.formId = formId;
    addQuestionsToForm(formId, questionsData, moshikomiStart, grades);
    const formUrl = form.getPublishedUrl();
    state.formUrl = formUrl;
    const editUrl = form.getEditUrl();

    DriveApp.getFileById(formId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const existingSheetIds = state.existingSheetIds;
    ss.getSheets().forEach(sheet => {
      existingSheetIds[String(sheet.getSheetId())] = true;
    });
    state.phase = 'フォーム回答シートの作成';
    state.destinationRequested = true;
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    // setDestination で追加されたシートだけを特定する。シート順には依存しない。
    let responseSheet = null;
    for (let attempt = 0; attempt < 8 && !responseSheet; attempt++) {
      Utilities.sleep(500);
      SpreadsheetApp.flush();
      const addedSheets = ss.getSheets().filter(sheet =>
        !existingSheetIds[String(sheet.getSheetId())]
      );
      if (addedSheets.length > 1) {
        throw new Error('フォーム回答シートが複数作成されたため特定できません。');
      }
      responseSheet = addedSheets[0] || null;
    }
    if (!responseSheet) {
      throw new Error('新しく作成されたフォーム回答シートを特定できません。');
    }
    state.responseSheetId = responseSheet.getSheetId();
    responseSheet.setName(sheetName);
    const responseColumnCount = responseSheet.getLastColumn();

    // メール管理シートへ書き込み
    const mailSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (mailSheet) {
      state.phase = 'メール管理への登録';
      state.mailSheet = mailSheet;
      state.mailPointerBefore = mailSheet.getRange(2, 3).getValue();
      const nextRow = mailSheet.getLastRow() + 1;
      state.mailMutation = {
        row: nextRow,
        previousValues: mailSheet.getRange(nextRow, 1, 1, 6).getValues(),
      };
      state.mailWritten = true;
      mailSheet.getRange(nextRow, 1, 1, 6).setValues(
        [[title, grades, '', 'リマインダー', sheetName + '\u3000案内', formUrl]]
      );
      mailSheet.getRange(2, 3).setValue(nextRow);
    }

    state.phase = '大会シート管理領域の作成';
    if (responseColumnCount > 4) {
      responseSheet.hideColumns(4, responseColumnCount - 4);
    }
    const grades2 = tournamentSheetDeclaredGrades_(sheetName);
    const cd = title.includes('鳳玉') ? 3000 : 2000;
    const eFee = title.includes('鳳玉') ? 2500 : 1500;
    const initialFees = { A: 2500, B: 2500, C: cd, D: cd, E: eFee };
    const initialSnapshot = {
      form_id: formId,
      form_edit_url: editUrl,
      is_sanctioned: !isKoen,
      settings: {
        '申込開始日': moshikomiStart,
        'リマインダー': new Date(
          moshiDead.getTime() - 6 * 24 * 60 * 60 * 1000
        ),
        '本申込期限': moshiDead,
        '抽選日': raffle || '',
        '本振込期限': huriDead || '',
        '大会の日時': '',
        'メモ': '',
        '後納制': huriDead ? 'before_tournament' : '',
        '振込先': '',
      },
      schedules: grades2.map(grade => ({
        grade: grade,
        participation_fee_yen: initialFees[grade],
        held_on: '',
      })),
    };
    tournamentSheetV2ValidateSnapshot_(initialSnapshot, false);
    const managementRows = tournamentSheetV2Rows_(initialSnapshot);
    const paymentStatusColumn = responseColumnCount + 1;
    const requiredColumns = paymentStatusColumn + 3;
    if (responseSheet.getMaxColumns() < requiredColumns) {
      responseSheet.insertColumnsAfter(
        responseSheet.getMaxColumns(),
        requiredColumns - responseSheet.getMaxColumns()
      );
    }
    responseSheet.getRange(1, paymentStatusColumn).setValue('振込み済みか');
    const managementStartRow = 3;
    const requiredRows = managementStartRow + managementRows.length - 1;
    if (responseSheet.getMaxRows() < requiredRows) {
      responseSheet.insertRowsAfter(
        responseSheet.getMaxRows(), requiredRows - responseSheet.getMaxRows()
      );
    }
    responseSheet.getRange(
      managementStartRow, 1, managementRows.length, TOURNAMENT_SHEET_V2_WIDTH_
    ).setValues(managementRows);
    responseSheet.getRange(managementStartRow, 1, 1, TOURNAMENT_SHEET_V2_WIDTH_)
      .setBackground('#d9ead3').setFontWeight('bold');
    // 作成した大会シートが共通構造判定で一意に読めることを検証する。
    tournamentSheetStructure_(responseSheet, true);

    // カレンダーシートへの書き込み
    state.phase = 'カレンダーへの登録';
    const calendarLastRowBefore = calendarSheet.getLastRow();
    const calendarNamesBefore = calendarLastRowBefore > 0
      ? calendarSheet.getRange(
        1, 1, calendarLastRowBefore, 1
      ).getValues()
      : [];
    const existingCalendarRows = [];
    for (let index = 2; index < calendarNamesBefore.length; index++) {
      if (String(calendarNamesBefore[index][0]) === sheetName) {
        existingCalendarRows.push(index + 1);
      }
    }
    if (existingCalendarRows.length > 1) {
      throw new Error(
        'カレンダーに同名の大会行が複数あります: ' + sheetName
      );
    }
    const existingCalendarRow = existingCalendarRows[0] || 0;
    const rowNum = existingCalendarRow || calendarLastRowBefore + 1;
    const writtenColumns = existingCalendarRow
      ? [3, 6, 8, 11]
      : [1, 3, 6, 8, 11];
    state.calendarMutation = {
      row: rowNum,
      snapshots: writtenColumns.map(column => ({
        column: column,
        value: calendarSheet.getRange(rowNum, column).getValue(),
      })),
    };
    state.calendarWritten = true;
    if (!existingCalendarRow) {
      calendarSheet.getRange(rowNum, 1).setValue(sheetName);
    }
    calendarSheet.getRange(rowNum, 3).setValue(Utilities.formatDate(moshikomiStart, 'JST', 'y/M/d'));
    calendarSheet.getRange(rowNum, 6).setValue(moshiDead);
    calendarSheet.getRange(rowNum, 8).setValue(raffle  || '未定');
    calendarSheet.getRange(rowNum, 11).setValue(huriDead || '未定');

    // 案内メール作成シートへの書き込み
    const announceSheet = ss.getSheetByName('案内メール作成');
    if (announceSheet) {
      state.phase = '案内メール作成設定の更新';
      state.announceSheet = announceSheet;
      ['B3:B4', 'B17', 'B28:B30'].forEach(a1 => {
        state.announceSnapshots.push({
          a1: a1,
          values: announceSheet.getRange(a1).getValues(),
        });
      });
      state.announceWritten = true;
      announceSheet.getRange(3, 2, 2, 1).setValues([[title], [grades]]);
      const tournamentQuestion = questionsData.find(question =>
        question[0] === '出場大会を全てお書きください。（略称等で構いません）'
      );
      const showMoshikomi = (tournamentQuestion && tournamentQuestion[1] === 1)
        ? Utilities.formatDate(moshikomiStart, 'JST', 'y/M/d')
        : '';
      announceSheet.getRange(28, 2).setValue(showMoshikomi);
      announceSheet.getRange(30, 2).setValue(formUrl);
      const reminderTime = new Date(moshiDead.getTime() - 6 * 24 * 60 * 60 * 1000);
      announceSheet.getRange(29, 2).setValue(reminderTime);
      announceSheet.getRange(17, 2).setValue(raffle || '');
    }

    return JSON.stringify({
      ok: true,
      formUrl: formUrl,
      db_sync_pending: false,
      warning: '',
    });
  } catch (err) {
    if (err && err.tournament_creation_uncertain === true) {
      state.dbTournamentCreationUncertain = true;
    }
    const createdResources = formCreateCreatedResources_(state);
    const rollback = formCreateRollback_(state);
    return JSON.stringify({
      error: String(err && err.message || err),
      phase: state.phase,
      created_resources: createdResources,
      rollback: rollback,
      google_form_creation_requested: Boolean(state.formCopyRequested),
      google_form_created: Boolean(state.formFileId),
      response_sheet_created: Boolean(state.responseSheetId),
      response_sheet_creation_requested: Boolean(state.destinationRequested),
    });
  } finally {
    if (lockAcquired && operationLock.hasLock()) operationLock.releaseLock();
  }
}
