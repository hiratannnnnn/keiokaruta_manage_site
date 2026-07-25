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

  const schedule = taikaiResolveSchedule_(match[1], grade, heldOn);
  const player = taikaiSplitPlayerName_(name);
  return taikaiApiRequest_('POST', '/registrations', {
    schedule_id: String(schedule.id),
    player: {
      family_name: player.family_name,
      given_name: player.given_name,
      email: email,
    },
  });
}

// フォーム作成時に一度だけ、回答先スプレッドシートの送信トリガーを登録する。
function ensureDatabaseFormSubmitTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === 'registerFormResponseToDatabase'
  );
  if (!exists) {
    ScriptApp.newTrigger('registerFormResponseToDatabase')
      .forSpreadsheet(CONFIG.SPREADSHEET_ID)
      .onFormSubmit()
      .create();
  }
}

// 既存フォームへ移行する際に、GASエディタから一度実行するための公開入口。
function setupDatabaseFormSubmitTrigger() {
  try {
    ensureDatabaseFormSubmitTrigger_();
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
