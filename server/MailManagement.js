// ============================================================
// メール管理シート操作
// ============================================================

// メール管理シートのデータを取得（行6以降）
function getMailManagement() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (!sheet) throw new Error('メール管理シートが見つかりません');

    const lastRow = sheet.getLastRow();
    if (lastRow < 6) return JSON.stringify([]);

    const rows = sheet.getRange(6, 1, lastRow - 5, 9).getValues()
      .map((row, i) => ({
        rowNum:         i + 6,
        tournamentName: String(row[0] || ''),
        grades:         String(row[1] || ''),
        sendDateTime:   row[2] instanceof Date
          ? Utilities.formatDate(row[2], 'JST', 'yyyy-MM-dd HH:mm:ss')
          : String(row[2] || ''),
        mailType:       String(row[3] || ''),
        threadTitle:    String(row[4] || ''),
        formLink:       String(row[5] || ''),
        reminderSet:    String(row[6] || ''),
        sent:           String(row[7] || ''),
        includeNotPaid: !!row[8],
      }))
      .filter(r => r.tournamentName);

    return JSON.stringify(rows);
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// メール管理シートの行を削除
function deleteMailManagementRow(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (!sheet) throw new Error('メール管理シートが見つかりません');
    sheet.deleteRow(rowNum);
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// メール管理シートに行を追加
function addMailManagementRow(json) {
  try {
    const d     = JSON.parse(json);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    const lastRow = Math.max(sheet.getLastRow(), 5);
    sheet.getRange(lastRow + 1, 1, 1, 9).setValues([[
      d.tournamentName || '',
      d.grades         || '',
      d.sendDateTime   || '',
      d.mailType       || 'リマインダー',
      d.threadTitle    || '',
      d.formLink       || '',
      d.mailType === '振込確認' ? '済' : '',
      '',
      false,
    ]]);

    // 振込確認の場合、カレンダーシートにも振込確認送信日を記録（col 9）
    if (d.mailType === '振込確認' && d.tournamentName && d.sendDateTime) {
      try {
        const calSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
        if (calSheet) {
          const fullName = (d.tournamentName || '') + (d.grades || '');
          const calData  = calSheet.getRange(1, 1, calSheet.getLastRow(), 1).getValues();
          for (let i = 2; i < calData.length; i++) {
            if (String(calData[i][0]) === fullName) {
              const dateStr = d.sendDateTime.slice(0, 10).replace(/-/g, '/');
              calSheet.getRange(i + 1, 9).setValue(dateStr);
              break;
            }
          }
        }
      } catch(calErr) {
        // カレンダー更新のエラーは無視して続行
      }
    }

    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// ------------------------------------------------
// 内部: 参加者リスト構築（makeParticipantList 相当）
// ------------------------------------------------
function makeParticipantList_(tournamentName, grades, includeNotPaid) {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tournamentName + grades);
  if (!sheet) return '（シートが見つかりません）\n';

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const count = headerRow.find(c => typeof c === 'number' && Number.isFinite(c));
  if (count == null) return '（N が取得できません）\n';

  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, count + 1).getValues()
    .filter(row => row[2] !== '');

  const gradesArray = grades.replace('級', '').split('');
  const byGrade = {};
  gradesArray.forEach(g => { byGrade[g + '級'] = []; });
  const paidList = [];

  data.forEach(row => {
    const name      = String(row[0]).replace('　', ' ');
    const gradeStr  = String(row[2]);
    const isPaid    = String(row[row.length - 1]).trim() === '済';

    if (isPaid && !paidList.includes(name)) paidList.push(name);

    gradeStr.replace('級', '').split('').forEach(g => {
      const key = g + '級';
      if (byGrade[key] && !byGrade[key].includes(name)) byGrade[key].push(name);
    });
  });

  const total = Object.values(byGrade).reduce((s, a) => s + a.length, 0);
  let result = '';
  if (total === 0) {
    result = '現在、どなたからも申し込みは頂いておりません。\n';
  } else {
    result += '現在、以下の方々からご連絡を頂いております。(敬称略)\n';
    for (const grade in byGrade) {
      result += grade + '：' + byGrade[grade].join('、') + '\n';
    }
    if (includeNotPaid && paidList.length > 0) {
      result +=
        '\nそのうち、振込の完了を確認できている方：\n' + paidList.join('、') + '\n\n' +
        '確認漏れがございましたら申し訳ありません。\n' +
        '明日までに振込を確認できない場合はキャンセルとせざるを得ない場合がありますので、ご了承ください。\n' +
        'ご相談等がある場合は、このメールに返信するか、会長 / 副会長までご連絡ください。\n';
    }
  }
  return result;
}

// 内部: Gmailから件名で最新スレッドを引用（quoteEmail 相当）
function quoteEmail_(subject) {
  if (!subject) return '';
  try {
    const threads = GmailApp.search('subject:"' + subject + '"', 0, 1);
    if (threads.length === 0) return '';
    const messages = threads[0].getMessages();
    const body = messages[messages.length - 1].getPlainBody();
    return body.split('\n').map(l => '> ' + l).join('\n');
  } catch(e) {
    return '';
  }
}

// 内部: メール本文生成（createEmailBody 相当）
function createEmailBody_(tournamentName, grades, participantList, tomorrowDateStr, formLink, mailType, quotedContent) {
  let body = 'おはようございます。\n\nこのメールは、慶應かるた会のメールシステムにより自動で送信されています。\n\n';

  if (mailType === 'リマインダー') {
    body +=
      'こちらは、' + tournamentName + grades + 'のリマインダーです。\n\n' +
      participantList + '\n' +
      'これから参加表明をされる方は、明日' + tomorrowDateStr + 'までに以下のリンクからフォームにご回答ください。\n' +
      formLink + '\n\n' +
      'よろしくお願いします。';
  } else {
    body +=
      'こちらは、' + tournamentName + grades + 'の振込確認です。\n\n' +
      participantList + '\n' +
      'まだお済みでない方は、お早めにお振込みください。\n\n' +
      'よろしくお願いします。';
  }

  if (quotedContent) {
    body += '\n\n以下、先日の本件に関するメールのコピーです。\n' + quotedContent;
  }
  return body;
}

// リマインダーメール本文プレビュー取得
function getReminderPreview(json) {
  try {
    const { rowNum, includeNotPaid, sendDateTime } = JSON.parse(json);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    const row   = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];

    const tournamentName = String(row[0]);
    const grades         = String(row[1]);
    const sheetDateStr   = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'JST', 'yyyy-MM-dd HH:mm:ss')
      : String(row[2] || '');
    const sendDateStr = sendDateTime || sheetDateStr;
    const mailType    = String(row[3]);
    const threadTitle = String(row[4]);
    const formLink    = String(row[5]);

    if (!sendDateStr) return JSON.stringify({ error: '送信予定日時が未設定です' });

    const targetDate = new Date(sendDateStr);
    if (targetDate - new Date() < 60 * 60 * 1000) return JSON.stringify({ error: '送信予定日時は現在時刻から1時間以上後に設定してください' });

    const participantList = makeParticipantList_(tournamentName, grades, includeNotPaid);

    const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];
    const tomorrow = new Date(targetDate.getTime());
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowDateStr = (tomorrow.getMonth() + 1) + '月' + tomorrow.getDate() + '日（' + WEEKDAYS[tomorrow.getDay()] + '）';

    const quotedContent = quoteEmail_(threadTitle);
    const body          = createEmailBody_(tournamentName, grades, participantList, tomorrowDateStr, formLink, mailType, quotedContent);
    const sendDateFormatted = Utilities.formatDate(targetDate, 'JST', 'yyyy年MM月dd日 HH時mm分');

    return JSON.stringify({ ok: true, body, sendDateFormatted, tournamentName, grades, mailType });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// リマインダートリガー設定（setReminders 相当）
function setReminderTrigger(json) {
  try {
    const { rowNum, includeNotPaid, sendDateTime } = JSON.parse(json);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    const row   = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];

    // モーダルで指定された日時があればシートに上書き
    if (sendDateTime) sheet.getRange(rowNum, 3).setValue(sendDateTime);

    const sendDateStr = sendDateTime || (row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'JST', 'yyyy-MM-dd HH:mm:ss')
      : String(row[2] || ''));
    const targetDate = new Date(sendDateStr);
    if (targetDate - new Date() < 60 * 60 * 1000) return JSON.stringify({ error: '送信予定日時は現在時刻から1時間以上後に設定してください' });

    sheet.getRange(rowNum, 7, 1, 3).setValues([['済', '', includeNotPaid]]);

    const tournamentName = String(row[0]);
    const grades         = String(row[1]);
    const mailType       = String(row[3]);

    // リマインダーの場合: フォーム説明更新 + カレンダー更新
    if (mailType === 'リマインダー') {
      try {
        const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];
        const tomorrow = new Date(targetDate.getTime());
        tomorrow.setDate(tomorrow.getDate() + 1);
        const tomorrowStr = (tomorrow.getMonth() + 1) + '月' + tomorrow.getDate() + '日（' + WEEKDAYS[tomorrow.getDay()] + '）';

        const tSheet = ss.getSheetByName(tournamentName + grades);
        if (tSheet) {
          const headerRow = tSheet.getRange(1, 1, 1, tSheet.getLastColumn()).getValues()[0];
          const count = headerRow.find(c => typeof c === 'number' && Number.isFinite(c));
          if (count != null) {
            const formId = String(tSheet.getRange(1, count + 4).getValue());
            FormApp.openById(formId).setDescription(
              'こちらは' + tournamentName + grades + 'の参加表明フォームです。\n' +
              '該当項目に回答の上、送信してください。\n' +
              'このフォームの回答期限は【' + tomorrowStr + '】の23:59までです。\n' +
              '回答が正しく送信されている場合、入力いただいたメールアドレスに回答のコピーが届きますのでご確認ください。'
            );
          }
        }

        const calSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
        if (calSheet) {
          const calData = calSheet.getRange(1, 1, calSheet.getLastRow(), 1).getValues();
          for (let i = 1; i < calData.length; i++) {
            if (String(calData[i][0]) === tournamentName + grades) {
              calSheet.getRange(i + 1, 4).setValue(Utilities.formatDate(targetDate, 'JST', 'yyyy/MM/dd'));
              const moshikomiStart = new Date(targetDate.getTime() + 48 * 60 * 60 * 1000);
              calSheet.getRange(i + 1, 5).setValue(Utilities.formatDate(moshikomiStart, 'JST', 'yyyy/MM/dd'));
              break;
            }
          }
        }
      } catch(innerErr) {
        // フォーム/カレンダー更新のエラーは無視して続行
      }
    }

    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}
