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

// Web App からフォームとシートを作成する
function createFormFromWeb(paramsJson) {
  try {
    const p = JSON.parse(paramsJson);
    const { title, grades, questionsData,
            moshikomiStartStr, moshiDeadStr, raffleStr, huriDeadStr, isKoen } = p;
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

    // フォーム作成だけは回答シートを作る必要があるが、大会本体は先に新DBへ登録する。
    // 級別日程は開催日が確定するまで作成できないため、後のsaveTournamentDatesで登録する。
    let dbSyncPending = false;
    let dbSyncWarning = '';
    let dbTournament = null;
    try {
      dbTournament = taikaiEnsureTournament_(title, {
        operation: 'フォーム作成',
        outcome: '一時障害の場合はフォーム作成を継続し、DB未同期として記録',
      });
    } catch (apiError) {
      if (!taikaiIsTransientApiError_(apiError)) throw apiError;
      dbSyncPending = true;
      dbSyncWarning =
        'taikai_manage APIの一時障害によりDB登録を保留しました。'
        + ' API復旧後に「今年度のシート→DB完全同期」を実行してください。';
    }

    const ss            = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const calendarSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
    const formTitle     = title + grades + '\u3000参加表明フォーム';

    // フォームをテンプレートからコピー
    const originalFile = DriveApp.getFileById(CONFIG.FORM_TEMPLATE_ID);
    const folder       = DriveApp.getFolderById(CONFIG.FORM_FOLDER_TO);
    const newFormFile  = originalFile.makeCopy(formTitle, folder);
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
    addQuestionsToForm(formId, questionsData, moshikomiStart, grades);
    const formUrl = form.getPublishedUrl();
    const editUrl = form.getEditUrl();

    DriveApp.getFileById(formId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const existingSheetIds = {};
    ss.getSheets().forEach(sheet => {
      existingSheetIds[String(sheet.getSheetId())] = true;
    });
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
    responseSheet.setName(title + grades);
    const responseColumnCount = responseSheet.getLastColumn();

    // メール管理シートへ書き込み
    const mailSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (mailSheet) {
      const nextRow = mailSheet.getLastRow() + 1;
      mailSheet.getRange(nextRow, 1, 1, 6).setValues(
        [[title, grades, '', 'リマインダー', title + grades + '\u3000案内', formUrl]]
      );
      mailSheet.getRange(2, 3).setValue(nextRow);
    }

    if (responseColumnCount > 4) {
      responseSheet.hideColumns(4, responseColumnCount - 4);
    }
    const grades2 = ['A', 'B', 'C', 'D', 'E'];
    const cd = title.includes('鳳玉') ? 3000 : 2000;
    const eFee = title.includes('鳳玉') ? 2500 : 1500;
    const initialFees = { A: 2500, B: 2500, C: cd, D: cd, E: eFee };
    const initialSnapshot = {
      tournament_name: title,
      tournament_id: dbTournament ? dbTournament.id : '',
      form_id: formId,
      form_public_url: formUrl,
      form_edit_url: editUrl,
      registration_completed: false,
      payment_completed: false,
      is_sanctioned: !isKoen,
      sync_status: dbSyncPending ? 'pending_api' : 'pending_schedule',
      synced_at: dbSyncPending ? '' : new Date(),
      sync_error: dbSyncWarning,
      schedules: grades2.map(grade => ({
        grade: grade,
        participation_fee_yen: initialFees[grade],
        held_on: '',
        id: '',
        application_deadline: moshiDead,
        internal_payment_deadline: '',
        payment_deadline: huriDead || '',
        payment_timing: huriDead ? 'before_tournament' : '',
        lottery_result_date: raffle || '',
        venue: '',
        reception_ends_at: '',
        is_sanctioned: !isKoen,
        payment_instructions: '',
        sync_status: 'pending_schedule',
        synced_at: '',
      })),
      entries: [],
      announcements: [],
      email_jobs: [],
      legacy_records: [],
    };
    tournamentSheetV2ValidateSnapshot_(initialSnapshot, false);
    const managementRows = tournamentSheetV2Rows_(initialSnapshot);
    if (responseSheet.getMaxColumns() < TOURNAMENT_SHEET_V2_WIDTH_) {
      responseSheet.insertColumnsAfter(
        responseSheet.getMaxColumns(),
        TOURNAMENT_SHEET_V2_WIDTH_ - responseSheet.getMaxColumns()
      );
    }
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
    const rowNum = findFromCalendar(calendarSheet, title + grades);
    calendarSheet.getRange(rowNum, 3).setValue(Utilities.formatDate(moshikomiStart, 'JST', 'y/M/d'));
    calendarSheet.getRange(rowNum, 6).setValue(moshiDead);
    calendarSheet.getRange(rowNum, 8).setValue(raffle  || '未定');
    calendarSheet.getRange(rowNum, 11).setValue(huriDead || '未定');

    // 案内メール作成シートへの書き込み
    const announceSheet = ss.getSheetByName('案内メール作成');
    if (announceSheet) {
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

    if (dbSyncPending) {
      if (!taikaiMarkTournamentPending_(title, dbSyncWarning)) {
        dbSyncWarning +=
          ' なお、DB未同期状態の保存に失敗したため、この警告を控えておいてください。';
      }
    } else {
      taikaiClearPendingTournaments_([title]);
    }
    return JSON.stringify({
      ok: true,
      formUrl: formUrl,
      db_sync_pending: dbSyncPending,
      warning: dbSyncWarning,
    });
  } catch (err) {
    return JSON.stringify({ error: err.message });
  }
}
