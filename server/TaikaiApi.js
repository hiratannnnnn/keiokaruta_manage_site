// ============================================================
// taikai_manage 新DB API共通クライアント
// ============================================================

function taikaiApiConfig_() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = String(props.getProperty('TAIKAI_API_BASE_URL') || '').replace(/\/+$/, '');
  if (!baseUrl) throw new Error('TAIKAI_API_BASE_URL が設定されていません。');
  // 暫定運用: API側のBearer認証が安定するまで、認証ヘッダーを送らない。
  return { baseUrl };
}

function taikaiApiQuery_(query) {
  return Object.keys(query || {}).filter(key => {
    const value = query[key];
    return value !== undefined && value !== null && value !== '';
  }).map(key => encodeURIComponent(key) + '=' + encodeURIComponent(String(query[key]))).join('&');
}

function taikaiApiRequest_(method, path, body, query) {
  const config = taikaiApiConfig_();
  const queryText = taikaiApiQuery_(query);
  const options = {
    method: String(method).toLowerCase(),
    muteHttpExceptions: true,
    headers: {},
  };
  if (body !== undefined && body !== null) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }

  const url = config.baseUrl + path + (queryText ? '?' + queryText : '');
  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const text = response.getContentText();
  taikaiRecordApiDebug_(method, path, status, text);
  let parsed = null;
  try { parsed = text ? JSON.parse(text) : null; } catch (e) {
    throw new Error('taikai_manage APIから不正なJSON応答を受信しました。');
  }
  if (status < 200 || status >= 300) {
    const error = parsed && parsed.error;
    const message = error && error.message ? error.message : 'taikai_manage APIへの接続に失敗しました。';
    throw new Error(message + ' (HTTP ' + status + ')');
  }
  return parsed && Object.prototype.hasOwnProperty.call(parsed, 'data') ? parsed.data : parsed;
}

// 開発用: 直近のAPI生レスポンスをユーザー単位の一時キャッシュへ保存する。
// 認証情報は現在送信していないが、将来復旧してもヘッダーは保存しない。
function taikaiRecordApiDebug_(method, path, status, text) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (String(props.getProperty('TAIKAI_API_DEBUG_LOG') || 'true').toLowerCase() === 'false') return;
    CacheService.getUserCache().put('taikai_api_debug_last', JSON.stringify({
      at: Utilities.formatDate(new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"),
      method: String(method).toUpperCase(),
      path: path,
      status: status,
      body: String(text || '').slice(0, 90000),
    }), 600);
  } catch (e) {
    // デバッグログの失敗で本来のAPI処理を失敗させない。
  }
}

function getTaikaiApiDebugLog() {
  try {
    const raw = CacheService.getUserCache().get('taikai_api_debug_last');
    return raw || JSON.stringify({ message: 'APIレスポンスはまだありません。' });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function taikaiFindPlayer_(name, email) {
  const query = email ? { email: email } : { name: name };
  const players = taikaiApiRequest_('GET', '/players', null, query) || [];
  if (!players.length) return null;
  if (players.length !== 1) throw new Error('同姓同名の選手がいるため、選手を特定できません。');
  return players[0];
}

function taikaiFindTournament_(name) {
  const tournaments = taikaiApiRequest_('GET', '/tournaments', null, { name: name }) || [];
  const exact = tournaments.filter(item => String(item.name) === String(name));
  if (exact.length !== 1) throw new Error('大会を一意に特定できません: ' + name);
  return exact[0];
}

function taikaiEnsureTournament_(name) {
  const tournaments = taikaiApiRequest_('GET', '/tournaments', null, { name: name }) || [];
  const exact = tournaments.filter(item => String(item.name) === String(name));
  if (exact.length > 1) throw new Error('同じ大会名が複数登録されています: ' + name);
  if (exact.length === 1) return exact[0];
  return taikaiApiRequest_('POST', '/tournaments', {
    name: name,
    registration_completed: false,
    payment_completed: false,
  });
}

function taikaiDeleteTournament_(name) {
  const tournament = taikaiFindTournament_(name);
  return taikaiApiRequest_(
    'DELETE',
    '/tournaments/' + encodeURIComponent(String(tournament.id))
  );
}

function taikaiJstDateTime_(value) {
  const text = String(value || '').trim().replace('T', ' ');
  const match = text.match(/^(\d{4}-\d{2}-\d{2})[ ](\d{2}:\d{2})(?::(\d{2}))?$/);
  if (!match) throw new Error('送信予定日時の形式が正しくありません。');
  return match[1] + 'T' + match[2] + ':' + (match[3] || '00') + '+09:00';
}

function taikaiRegisterPaymentEmailJob_(tournamentName, grades, sendDateTime, paymentDeadline) {
  const deadline = taikaiFormatDate_(paymentDeadline);
  if (!deadline) throw new Error('振込期限を入力してください。');

  const tournament = taikaiFindTournament_(tournamentName);
  const gradeSet = String(grades || '').replace(/級/g, '').toUpperCase().split('').filter(Boolean);
  const schedules = taikaiApiRequest_(
    'GET',
    '/tournaments/' + encodeURIComponent(String(tournament.id)) + '/schedules'
  ) || [];
  const targets = schedules.filter(schedule => gradeSet.includes(String(schedule.grade).toUpperCase()));
  if (!targets.length) {
    throw new Error('対象級の大会日程がDBに登録されていません。');
  }

  targets.forEach(schedule => {
    taikaiApiRequest_('PATCH', '/schedules/' + encodeURIComponent(String(schedule.id)), {
      payment_deadline: deadline,
    });
  });

  return taikaiApiRequest_('POST', '/email-jobs', {
    scheduled_at: taikaiJstDateTime_(sendDateTime),
    mail_type: 'payment_confirmation',
    announcement_id: null,
    recipient_group_id: null,
    schedule_ids: targets.map(schedule => String(schedule.id)),
  });
}

function taikaiFindSchedule_(tournamentName, grade, heldOn) {
  const tournament = taikaiFindTournament_(tournamentName);
  const schedules = taikaiApiRequest_('GET', '/schedules', null, {
    tournament_id: tournament.id,
    grade: String(grade).replace('級', '').toUpperCase(),
    held_on: heldOn,
  }) || [];
  if (schedules.length !== 1) {
    throw new Error('大会日程を一意に特定できません: ' + tournamentName + ' ' + grade + '級 ' + heldOn);
  }
  return schedules[0];
}

function taikaiResolveSchedule_(tournamentName, grade, heldOn) {
  const tournament = taikaiFindTournament_(tournamentName);
  const query = {
    tournament_id: tournament.id,
    grade: String(grade).replace('級', '').toUpperCase(),
  };
  if (heldOn) query.held_on = heldOn;
  const schedules = taikaiApiRequest_('GET', '/schedules', null, query) || [];
  if (schedules.length !== 1) {
    throw new Error('大会日程を一意に特定できません: ' + tournamentName + ' ' + grade + '級');
  }
  return schedules[0];
}

function taikaiSplitPlayerName_(name) {
  const parts = String(name || '').trim().split(/[ 　]+/).filter(Boolean);
  if (parts.length < 2) throw new Error('氏名を名字と名前に分けられません: ' + name);
  return { family_name: parts.shift(), given_name: parts.join(' ') };
}

function taikaiFormatDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'JST', 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return '';
  const date = new Date(text);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, 'JST', 'yyyy-MM-dd');
}

function taikaiGetParticipations_(name, beforeDate, email) {
  const player = taikaiFindPlayer_(name, email);
  if (!player) return [];
  const query = {};
  if (beforeDate) query.held_on_to = taikaiFormatDate_(beforeDate);
  const path = '/players/' + encodeURIComponent(String(player.id)) + '/participations';
  const activeRows = taikaiApiRequest_(
    'GET', path, null, Object.assign({}, query, { canceled: false })
  ) || [];
  const canceledRows = taikaiApiRequest_(
    'GET', path, null, Object.assign({}, query, { canceled: true })
  ) || [];
  const latestByTournament = {};
  activeRows.concat(canceledRows).forEach(item => {
    const key = String(item.tournament_id);
    const current = latestByTournament[key];
    if (!current || taikaiCompareIds_(item.entry_id, current.entry_id) > 0) {
      latestByTournament[key] = item;
    }
  });
  return Object.keys(latestByTournament).map(key => latestByTournament[key])
    .filter(item => !item.canceled_at)
    .sort((a, b) =>
      String(a.held_on || '').localeCompare(String(b.held_on || '')) ||
      taikaiCompareIds_(a.entry_id, b.entry_id)
    )
    .map(item => ({
    date: item.held_on || '',
    location: item.tournament_name || '',
    raffleDate: item.lottery_result_date || '',
    isOfficial: item.is_sanctioned === true,
    grade: item.grade || '',
    canceledAt: item.canceled_at || null,
    raw: item,
    }));
}

function taikaiCompareIds_(left, right) {
  const leftText = String(left === undefined || left === null ? '' : left);
  const rightText = String(right === undefined || right === null ? '' : right);
  if (/^\d+$/.test(leftText) && /^\d+$/.test(rightText) &&
      leftText.length !== rightText.length) {
    return leftText.length - rightText.length;
  }
  return leftText.localeCompare(rightText);
}

function taikaiRegisterEntry_(tournamentName, grade, heldOn, playerName, email) {
  const schedule = taikaiResolveSchedule_(tournamentName, grade, heldOn);
  const player = taikaiSplitPlayerName_(playerName);
  return taikaiApiRequest_('POST', '/registrations', {
    schedule_id: String(schedule.id),
    player: {
      family_name: player.family_name,
      given_name: player.given_name,
      email: String(email || '').trim(),
    },
  });
}

function taikaiFindTournamentEntry_(tournamentName, playerName) {
  const tournament = taikaiFindTournament_(tournamentName);
  const player = taikaiFindPlayer_(playerName);
  if (!player) throw new Error('API上で選手が見つかりません: ' + playerName);
  const entries = taikaiApiRequest_('GET', '/entries', null, {
    tournament_id: tournament.id,
    player_id: player.id,
  }) || [];
  if (!entries.length) throw new Error('大会と選手に対応する出場登録が見つかりません。');
  return entries.reduce((latest, entry) =>
    !latest || taikaiCompareIds_(entry.id, latest.id) > 0 ? entry : latest
  , null);
}

function taikaiSetPaymentByPlayer_(tournamentName, playerName, isPaid) {
  const entry = taikaiFindTournamentEntry_(tournamentName, playerName);
  return taikaiApiRequest_('PUT', '/entries/' + encodeURIComponent(String(entry.id)) + '/payment', {
    is_paid: Boolean(isPaid),
  });
}

function taikaiSetTournamentSanctioned_(tournamentName, isSanctioned) {
  const tournament = taikaiFindTournament_(tournamentName);
  const schedules = taikaiApiRequest_('GET', '/tournaments/' + encodeURIComponent(String(tournament.id)) + '/schedules') || [];
  schedules.forEach(schedule => {
    taikaiApiRequest_('PATCH', '/schedules/' + encodeURIComponent(String(schedule.id)), {
      is_sanctioned: Boolean(isSanctioned),
    });
  });
  return schedules.length;
}

function taikaiSyncTournamentSchedulesFromSheet_(sheetName, gradeDates) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const calendar = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
  const tournamentSheet = ss.getSheetByName(sheetName);
  if (!calendar || !tournamentSheet) throw new Error('大会情報シートが見つかりません。');

  const baseName = String(sheetName).replace(/[A-E]+級$/, '');
  const tournament = taikaiEnsureTournament_(baseName);
  const calendarRows = calendar.getRange(1, 1, calendar.getLastRow(), 12).getValues();
  const calendarRow = calendarRows.slice(2).find(row => String(row[0]) === sheetName);
  if (!calendarRow) throw new Error('カレンダーに大会情報がありません: ' + sheetName);

  const header = tournamentSheet.getRange(1, 1, 1, tournamentSheet.getLastColumn()).getValues()[0];
  const registerIndex = tournamentSheet.getDataRange().getValues().findIndex(row => String(row[0]).trim() === 'registerDatabase');
  const isSanctioned = registerIndex < 0 || String(tournamentSheet.getRange(registerIndex + 2, 2).getValue()).trim() === '';
  const applicationDeadline = taikaiFormatDate_(calendarRow[5]);
  if (!applicationDeadline) throw new Error('申込期限が未設定のため、APIへ日程を登録できません。');
  const paymentDeadline = taikaiFormatDate_(calendarRow[10]) || null;
  const lotteryDate = taikaiFormatDate_(calendarRow[7]) || null;

  Object.keys(gradeDates || {}).forEach(grade => {
    const heldOn = taikaiFormatDate_(gradeDates[grade]);
    if (!heldOn) return;
    const schedules = taikaiApiRequest_('GET', '/schedules', null, {
      tournament_id: tournament.id,
      grade: grade,
      held_on: heldOn,
    }) || [];
    const body = {
      held_on: heldOn,
      grade: grade,
      application_deadline: applicationDeadline,
      payment_deadline: paymentDeadline,
      lottery_result_date: lotteryDate,
      payment_timing: null,
      venue: null,
      reception_ends_at: null,
      is_sanctioned: isSanctioned,
    };
    if (schedules.length === 0) {
      taikaiApiRequest_('POST', '/tournaments/' + encodeURIComponent(String(tournament.id)) + '/schedules', body);
    } else if (schedules.length === 1) {
      taikaiApiRequest_('PATCH', '/schedules/' + encodeURIComponent(String(schedules[0].id)), body);
    } else {
      throw new Error('同じ大会・級・開催日のAPI日程が複数あります: ' + sheetName + ' ' + grade);
    }
  });
  return tournament;
}
