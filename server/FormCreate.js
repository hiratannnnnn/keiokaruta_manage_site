// ============================================================
// フォーム作成（createFormAndSheet の Web App 版）
// ============================================================

// フォームに質問を追加し、フォームカラム数（N）を返す
// questionsData : [[name, inc, req], ...]
// moshikomiStart: Date オブジェクト（公認大会出場回数タイトル計算用）
// grades        : 例 "ABCDE級"
function addQuestionsToForm(formId, questionsData, moshikomiStart, grades) {
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
    const actualTitle = isKounin ? kouninTitle : name;

    if (name === '級') {
      const item = form.addMultipleChoiceItem();
      item.setTitle(actualTitle).setRequired(req === 1);
      if (gradeSet.length) item.setChoiceValues(gradeSet);
    } else if (name === '段位') {
      const item = form.addMultipleChoiceItem();
      item.setTitle(actualTitle).setRequired(req === 1).setChoiceValues(DANKAI);
    } else {
      const item = form.addTextItem();
      item.setTitle(actualTitle).setRequired(req === 1);
      if (name.includes('氏名')) {
        item.setHelpText('半角または全角スペースを含めてください。');
      } else if (isKounin) {
        item.setHelpText(
          '今年度の「申込開始日時点で、出場した、または出場することが分かっている公認大会」' +
          'の回数をお書きください。\n現在慶應かるた会botの出場回数確認は機能していません。自己管理をお願いします。'
        );
      }
    }
  });

  return 1 + questionCount; // 1 = タイムスタンプ列
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

    const moshikomiStart = new Date(moshikomiStartStr);
    const moshiDead      = new Date(moshiDeadStr);
    const raffle         = raffleStr   ? new Date(raffleStr)   : null;
    const huriDead       = huriDeadStr ? new Date(huriDeadStr) : null;

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
    form.setCollectEmail(false);
    form.setLimitOneResponsePerUser(false);

    const formId  = form.getId();
    const count   = addQuestionsToForm(formId, questionsData, moshikomiStart, grades);
    const formUrl = form.getPublishedUrl();
    const editUrl = form.getEditUrl();

    DriveApp.getFileById(formId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    Utilities.sleep(2000);
    SpreadsheetApp.flush();

    // 新規作成されたフォーム回答シートを設定
    const firstSheet = ss.getSheets()[0];
    firstSheet.setName(title + grades);

    // メール管理シートへ書き込み
    const mailSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (mailSheet) {
      const nextRow = mailSheet.getLastRow() + 1;
      mailSheet.getRange(nextRow, 1, 1, 6).setValues(
        [[title, grades, '', 'リマインダー', title + grades + '\u3000案内', formUrl]]
      );
      mailSheet.getRange(2, 3).setValue(nextRow);
    }

    firstSheet.getRange(1, count + 6).setValue(count).setBackground('#D3D3D3');
    if (count > 2) firstSheet.hideColumns(4, count - 2);
    firstSheet.getRange(1, count + 3, 1, 3).setValues([['振込み済みか', formId, editUrl]]);
    firstSheet.getRange(4, count + 4, 3, 2).setValues([
      ['催促メール設定（「☆大会フォーム」から）', '↓送信予定日時（yyyy-MM-dd HH:mm:ss）'],
      ['未設定', ''],
      ['後納制の場合は→に何か文字を', ''],
    ]).setBackground('#D3D3D3');
    firstSheet.getRange(2, count + 6, 5, 1).setValues(
      [['setPaymentReminders'], [''], ['moveToDone'], [''], ['deleteSheet']]
    );
    firstSheet.hideColumns(count + 6);
    firstSheet.getRange(5, count + 5).setBackground('#FFF2CC');
    firstSheet.getRange(6, count + 5).setBackground('#FFF2CC');
    firstSheet.getRange(1, count + 3).setBackground('#FFF2CC');
    firstSheet.getRange(12, 1, 4, 3).setValues([
      ['registerDatabase', '非公認の場合は下に文字', '↓振込先'],
      ['', '', ''],
      ['countMatches', '申し込みのときに回数数える', ''],
      ['', '', ''],
    ]).setBackground('#d3d3d3');

    if (isKoen) firstSheet.getRange(13, 2).setValue('非公認');

    firstSheet.getRange(13, 1, 1, 3).setBackground('#FFF2CC');
    firstSheet.getRange(15, 1).setBackground('#FFF2CC');
    firstSheet.getRange(14, 2, 2, 1).setBackground('#FFFFFF');
    firstSheet.getRange(14, 3, 4, 1).setBackground('#FFF2CC');

    const grades2 = ['A', 'B', 'C', 'D', 'E'];
    for (let i = 0; i < grades2.length; i++) {
      firstSheet.getRange(4 + i, 1).setValue(grades2[i]);
      const col = count + 3;
      const formula =
        '=COUNTIFS(E$1:E, "' + grades2[i] + '", INDIRECT("R1C" & (' + col + ') & ":R" & ROWS(E:E) & "C" & (' + col + '), FALSE), "済")' +
        ' + COUNTIFS(E$1:E, "' + grades2[i] + '", INDIRECT("R1C" & (' + col + ') & ":R" & ROWS(E:E) & "C" & (' + col + '), FALSE), "*繰*越*")';
      firstSheet.getRange(4 + i, 3).setFormula(formula);
      firstSheet.getRange(4 + i, count + 2).setFormula('=MULTIPLY(B' + (4 + i) + ',C' + (4 + i) + ')');
    }
    firstSheet.getRange(4, 1, 5, 1).setHorizontalAlignment('right');

    const cd = title.includes('鳳玉') ? 3000 : 2000;
    const e  = title.includes('鳳玉') ? 2500 : 1500;
    firstSheet.getRange('B4:B5').setValue(2500);
    firstSheet.getRange('B6:B7').setValue(cd);
    firstSheet.getRange('B8').setValue(e);
    firstSheet.getRange(9, 3).setFormula('=SUM($C4:$C8)');
    firstSheet.getRange(9, count + 2).setFormula('=SUMPRODUCT(B4:B8, C4:C8)');
    firstSheet.getRange(10, 1, 2, 3).setValues([
      ['原則として、振込が確認出来たら「済」と入れる。', '', ''],
      ['何らかの理由ですでに振込がある分から回す場合は「繰り越し」と入れる。', '', ''],
    ]).setBackground('#d3d3d3');

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
      const showMoshikomi = (questionsData[6] && questionsData[6][1] === 1)
        ? Utilities.formatDate(moshikomiStart, 'JST', 'y/M/d')
        : '';
      announceSheet.getRange(28, 2).setValue(showMoshikomi);
      announceSheet.getRange(30, 2).setValue(formUrl);
      const reminderTime = new Date(moshiDead.getTime() - 6 * 24 * 60 * 60 * 1000);
      announceSheet.getRange(29, 2).setValue(reminderTime);
      announceSheet.getRange(17, 2).setValue(raffle || '');
    }

    return JSON.stringify({ ok: true, formUrl });
  } catch (err) {
    return JSON.stringify({ error: err.message });
  }
}
