// ============================================================
// taikai_manage 新DB API共通クライアント
// ============================================================

function taikaiApiConfig_() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = String(props.getProperty('TAIKAI_API_BASE_URL') || '').replace(/\/+$/, '');
  const token = String(props.getProperty('TAIKAI_API_TOKEN') || '').trim();
  if (!baseUrl) throw new Error('TAIKAI_API_BASE_URL が設定されていません。');
  if (!token) throw new Error('TAIKAI_API_TOKEN が設定されていません。');
  return { baseUrl, token };
}

function taikaiApiQuery_(query) {
  return Object.keys(query || {}).filter(key => {
    const value = query[key];
    return value !== undefined && value !== null && value !== '';
  }).map(key => encodeURIComponent(key) + '=' + encodeURIComponent(String(query[key]))).join('&');
}

function taikaiApiError_(message, details) {
  const error = new Error(message);
  const info = details || {};
  error.taikai_api_error = true;
  error.transient = info.transient === true;
  error.http_status = info.http_status || null;
  error.api_method = String(info.api_method || '');
  error.api_path = String(info.api_path || '');
  error.api_error_code = String(info.api_error_code || '');
  error.taikai_request_id = String(info.request_id || '');
  error.validation_fields = info.validation_fields || {};
  return error;
}

function taikaiApiSafeToken_(value, pattern, maxLength) {
  const text = String(value || '');
  return text.length <= maxLength && pattern.test(text) ? text : '';
}

function taikaiApiSafeValidationFields_(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  const result = {};
  Object.keys(value).slice(0, 20).forEach(key => {
    const safeKey = taikaiApiSafeToken_(
      key, /^[a-zA-Z0-9_.-]+$/, 64
    );
    const message = String(value[key] || '').replace(/\s+/g, ' ').trim();
    if (safeKey && message && message.length <= 300) {
      result[safeKey] = message;
    }
  });
  return result;
}

function taikaiApiFailureMessage_(status, method, path, apiError) {
  const error = apiError && typeof apiError === 'object' ? apiError : {};
  const code = taikaiApiSafeToken_(
    error.code, /^[a-z0-9_]+$/, 64
  );
  const requestId = taikaiApiSafeToken_(
    error.request_id, /^req_[a-f0-9]+$/i, 80
  );
  const fields = taikaiApiSafeValidationFields_(error.fields);
  let message = String(error.message || '').trim()
    || 'taikai_manage APIへの接続に失敗しました。';
  if (Number(status) === 401) {
    message = 'API認証に失敗しました。'
      + 'GASのTAIKAI_API_TOKENとPHPのAPI_TOKENが一致しているか確認してください。';
  }
  const diagnostics = [
    'HTTP ' + status,
    code ? 'code: ' + code : '',
    String(method || '').toUpperCase() + ' ' + String(path || ''),
    requestId ? 'request ID: ' + requestId : '',
  ].filter(Boolean);
  const fieldMessages = Object.keys(fields).map(key =>
    key + ': ' + fields[key]
  );
  return message + ' [' + diagnostics.join(' / ') + ']'
    + (fieldMessages.length
      ? ' 入力エラー: ' + fieldMessages.join('、')
      : '');
}

function taikaiIsTransientApiError_(error) {
  return Boolean(error && error.taikai_api_error && error.transient === true);
}

function taikaiApiAlertKey_(method, path, status) {
  const normalizedPath = String(path || '')
    .replace(/\/\d+(?=\/|$)/g, '/:id')
    .replace(
      /\/[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}(?=\/|$)/ig,
      '/:id'
    );
  const source = [method, normalizedPath, status || 'connection'].join('|');
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    source,
    Utilities.Charset.UTF_8
  );
  return 'taikai_api_alert_' + digest.slice(0, 12).map(value =>
    ('0' + ((value + 256) % 256).toString(16)).slice(-2)
  ).join('');
}

function taikaiNotifyApiFailure_(details) {
  try {
    const props = PropertiesService.getScriptProperties();
    const recipient = String(props.getProperty('TAIKAI_API_ALERT_EMAIL') || '').trim();
    if (!recipient) return;
    const method = String(details.method || '').toUpperCase();
    const path = String(details.path || '');
    const status = details.status || '';
    const cache = CacheService.getScriptCache();
    const cacheKey = taikaiApiAlertKey_(method, path, status);
    if (cache.get(cacheKey)) return;

    const lines = [
      'taikai_manage APIの障害を検出しました。',
      '',
      '発生日時: ' + Utilities.formatDate(
        new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"
      ),
      '処理: ' + String(details.operation || 'API呼び出し'),
      'API: ' + method + ' ' + path,
      '状態: ' + (status ? 'HTTP ' + status : '接続エラー'),
      '概要: ' + String(details.message || '応答を取得できませんでした。'),
      '対象大会: ' + String(details.tournament_name || '不明'),
      '処理結果: ' + String(details.outcome || '呼び出し元で中断または再処理'),
      '',
      '認証情報・APIレスポンス本文・個人情報はこの通知に含めていません。',
    ];
    MailApp.sendEmail({
      to: recipient,
      subject: '【要確認】taikai_manage API障害',
      body: lines.join('\n'),
    });
    cache.put(cacheKey, 'sent', 900);
  } catch (e) {
    // 障害通知の失敗で、本来の処理やエラー応答を壊さない。
  }
}

function taikaiApiRequest_(method, path, body, query, context) {
  const config = taikaiApiConfig_();
  const queryText = taikaiApiQuery_(query);
  const options = {
    method: String(method).toLowerCase(),
    muteHttpExceptions: true,
    headers: {},
  };
  options.headers.Authorization = 'Bearer ' + config.token;
  if (body !== undefined && body !== null) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }

  const url = config.baseUrl + path + (queryText ? '?' + queryText : '');
  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    const message = 'taikai_manage APIへ接続できませんでした。';
    taikaiNotifyApiFailure_(Object.assign({}, context || {}, {
      method: method,
      path: path,
      message: message,
    }));
    throw taikaiApiError_(message, {
      transient: true,
      api_method: method,
      api_path: path,
    });
  }
  const status = response.getResponseCode();
  const text = response.getContentText();
  let parsed = null;
  try { parsed = text ? JSON.parse(text) : null; } catch (e) {
    const message = 'taikai_manage APIから不正なJSON応答を受信しました。';
    const transient = status === 408 || status === 429 || status >= 500
      || (status >= 200 && status < 300);
    taikaiNotifyApiFailure_(Object.assign({}, context || {}, {
      method: method,
      path: path,
      status: status,
      message: message,
    }));
    throw taikaiApiError_(message, {
      transient: transient,
      http_status: status,
      api_method: method,
      api_path: path,
    });
  }
  if (status < 200 || status >= 300) {
    const error = parsed && parsed.error;
    const message = taikaiApiFailureMessage_(
      status, method, path, error
    );
    const errorCode = taikaiApiSafeToken_(
      error && error.code, /^[a-z0-9_]+$/, 64
    );
    const requestId = taikaiApiSafeToken_(
      error && error.request_id, /^req_[a-f0-9]+$/i, 80
    );
    const validationFields = taikaiApiSafeValidationFields_(
      error && error.fields
    );
    const transient = status === 408 || status === 429 || status >= 500;
    if (transient) {
      taikaiNotifyApiFailure_(Object.assign({}, context || {}, {
        method: method,
        path: path,
        status: status,
        message: 'APIがHTTPエラーを返しました。',
      }));
    }
    throw taikaiApiError_(message, {
      transient: transient,
      http_status: status,
      api_method: method,
      api_path: path,
      api_error_code: errorCode,
      request_id: requestId,
      validation_fields: validationFields,
    });
  }
  return parsed && Object.prototype.hasOwnProperty.call(parsed, 'data') ? parsed.data : parsed;
}

function taikaiFindPlayer_(name, email) {
  const query = email ? { email: pseudonymousEmailFor_(email) } : { name: name };
  const players = taikaiApiRequest_('GET', '/players', null, query) || [];
  if (!players.length) return null;
  if (players.length !== 1) throw new Error('同姓同名の選手がいるため、選手を特定できません。');
  return players[0];
}

function taikaiFindTournament_(name) {
  const tournaments = taikaiApiRequest_('GET', '/tournaments', null, { name: name }) || [];
  const exact = tournaments.filter(item => String(item.name) === String(name));
  if (exact.length === 0) {
    throw new Error(
      'DBに大会が登録されていません: ' + name
      + '。今年度開催の大会は、先に「今年度のシート→DB完全同期」を実行してください。'
    );
  }
  if (exact.length > 1) {
    throw new Error(
      'DBに同名の大会が' + exact.length + '件登録されています: ' + name
      + '。tournaments.nameの重複データを解消してください。'
    );
  }
  return exact[0];
}

function taikaiEnsureTournamentWithState_(name, context) {
  const requestContext = Object.assign({ tournament_name: name }, context || {});
  const tournaments = taikaiApiRequest_(
    'GET', '/tournaments', null, { name: name }, requestContext
  ) || [];
  if (!Array.isArray(tournaments)) {
    const message = 'taikai_manage APIの大会一覧応答形式が不正です。';
    taikaiNotifyApiFailure_(Object.assign({}, requestContext, {
      method: 'GET',
      path: '/tournaments',
      status: 200,
      message: message,
    }));
    throw taikaiApiError_(message, {
      transient: true,
      http_status: 200,
      api_method: 'GET',
      api_path: '/tournaments',
    });
  }
  const exact = tournaments.filter(item => String(item.name) === String(name));
  if (exact.length > 1) throw new Error('同じ大会名が複数登録されています: ' + name);
  if (exact.length === 1) {
    return { tournament: exact[0], created: false };
  }
  let tournament;
  try {
    tournament = taikaiApiRequest_('POST', '/tournaments', {
      name: name,
      registration_completed: false,
      payment_completed: false,
    }, null, requestContext);
  } catch (error) {
    if (taikaiIsTransientApiError_(error)) {
      error.tournament_creation_uncertain = true;
    }
    throw error;
  }
  return { tournament: tournament, created: true };
}

function taikaiEnsureTournament_(name, context) {
  return taikaiEnsureTournamentWithState_(name, context).tournament;
}

function taikaiPendingTournaments_() {
  const props = PropertiesService.getScriptProperties();
  try {
    const parsed = JSON.parse(props.getProperty('TAIKAI_DB_PENDING_TOURNAMENTS') || '{}');
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (e) {
    return {};
  }
}

function taikaiMarkTournamentPending_(name, message) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const pending = taikaiPendingTournaments_();
    pending[String(name)] = {
      recorded_at: Utilities.formatDate(new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"),
      reason: String(message || 'API一時障害'),
    };
    PropertiesService.getScriptProperties().setProperty(
      'TAIKAI_DB_PENDING_TOURNAMENTS', JSON.stringify(pending)
    );
    return true;
  } catch (e) {
    return false;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function taikaiClearPendingTournaments_(names) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const pending = taikaiPendingTournaments_();
    (names || []).forEach(name => { delete pending[String(name)]; });
    const props = PropertiesService.getScriptProperties();
    if (Object.keys(pending).length) {
      props.setProperty('TAIKAI_DB_PENDING_TOURNAMENTS', JSON.stringify(pending));
    } else {
      props.deleteProperty('TAIKAI_DB_PENDING_TOURNAMENTS');
    }
    return true;
  } catch (e) {
    return false;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function taikaiDeleteTournament_(name) {
  const tournament = taikaiFindTournament_(name);
  return taikaiApiRequest_(
    'DELETE',
    '/tournaments/' + encodeURIComponent(String(tournament.id))
  );
}

function taikaiDeleteTournamentSchedules_(tournamentId, scheduleIds) {
  const ids = (scheduleIds || []).map(id => String(id || '').trim())
    .filter(Boolean);
  if (!String(tournamentId || '').trim() || !ids.length) {
    throw new Error('大会IDまたは削除対象schedule IDが不足しています。');
  }
  return taikaiApiRequest_(
    'DELETE',
    '/tournaments/' + encodeURIComponent(String(tournamentId)) + '/schedules',
    { schedule_ids: ids }
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

function taikaiPaymentSchedule_(value, heldOn) {
  if (String(value || '').trim() === '当日支払い') {
    return {
      payment_deadline: heldOn,
      payment_timing: 'on_tournament_day',
    };
  }
  return {
    payment_deadline: taikaiFormatDate_(value) || null,
    payment_timing: null,
  };
}

function taikaiGetParticipations_(name, beforeDate, email) {
  const player = taikaiFindPlayer_(name, email);
  if (!player) return [];
  return taikaiGetParticipationsByPlayerId_(player.id, beforeDate);
}

function taikaiGetParticipationsByPlayerId_(playerId, beforeDate) {
  const query = {};
  if (beforeDate) query.held_on_to = taikaiFormatDate_(beforeDate);
  const path = '/players/' + encodeURIComponent(String(playerId)) + '/participations';
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
  const dbEmail = rememberPseudonymousEmail_(email, playerName);
  return taikaiApiRequest_('POST', '/registrations', {
    schedule_id: String(schedule.id),
    player: {
      family_name: player.family_name,
      given_name: player.given_name,
      email: dbEmail,
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

function taikaiRecordFullPaymentByEntry_(entryId, useDeposit) {
  const summary = taikaiApiRequest_(
    'GET', '/entries/' + encodeURIComponent(String(entryId)) + '/payment-summary'
  );
  const balance = Number(summary.balance_yen);
  if (!Number.isFinite(balance)) {
    throw new Error('参加費が未設定のため、支払いを登録できません。');
  }
  if (balance <= 0) return summary;
  const body = {
    idempotency_key: 'detail-' + Utilities.getUuid(),
    amount_yen: balance,
    paid_at: Utilities.formatDate(
      new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"
    ),
    note: '大会詳細から登録',
  };
  if (useDeposit) {
    const deposit = Number(summary.deposit_balance_yen || 0);
    if (deposit !== balance) {
      throw new Error(
        'デポジット残高と参加費残額が一致しません。出納管理から金額を確認してください。'
      );
    }
    return taikaiApiRequest_(
      'POST',
      '/entries/' + encodeURIComponent(String(entryId)) + '/deposit-application',
      body
    );
  }
  body.method = 'bank_transfer';
  return taikaiApiRequest_(
    'POST',
    '/entries/' + encodeURIComponent(String(entryId)) + '/payments',
    body
  );
}

function taikaiSetEntryCancellationStatus_(
  entryId, canceledAt, movePaymentsToDeposit
) {
  return taikaiApiRequest_(
    'POST',
    '/entries/' + encodeURIComponent(String(entryId))
      + '/cancellation-status',
    {
      idempotency_key: 'detail-cancel-' + Utilities.getUuid(),
      canceled_at: canceledAt,
      move_payments_to_deposit: movePaymentsToDeposit === true,
      occurred_at: Utilities.formatDate(
        new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"
      ),
      note: 'キャンセルにより参加費からデポジットへ振替',
    }
  );
}

function taikaiSetTournamentSanctioned_(tournamentName, isSanctioned, grades) {
  const tournament = taikaiFindTournament_(tournamentName);
  const schedules = taikaiApiRequest_('GET', '/tournaments/' + encodeURIComponent(String(tournament.id)) + '/schedules') || [];
  const gradeSet = (grades || []).map(grade =>
    String(grade || '').replace('級', '').toUpperCase()
  );
  const targets = gradeSet.length
    ? schedules.filter(schedule =>
      gradeSet.includes(String(schedule.grade || '').toUpperCase())
    )
    : schedules;
  if (!targets.length) {
    throw new Error(
      '対象級の大会日程がDBに登録されていません: '
      + tournamentName + ' ' + gradeSet.join('') + '級'
    );
  }
  targets.forEach(schedule => {
    taikaiApiRequest_('PATCH', '/schedules/' + encodeURIComponent(String(schedule.id)), {
      is_sanctioned: Boolean(isSanctioned),
    });
  });
  return targets.length;
}

function taikaiGradeFeesFromSheetData_(data, sheetName, grades) {
  const fees = {};
  const errors = [];
  let gradeRows = {};
  try {
    const responseEndIndex = tournamentSheetResponseEndIndex_(data);
    gradeRows = tournamentSheetGradeRows_(data, responseEndIndex, false);
  } catch (e) {
    errors.push(sheetName + ': ' + e.message);
  }
  (grades || []).forEach(grade => {
    const rowNumber = gradeRows[grade];
    if (!rowNumber) {
      errors.push(sheetName + ': ' + grade + '級の参加費設定行が見つかりません。');
      return;
    }
    const fee = (data[rowNumber - 1] || [])[1];
    if (typeof fee !== 'number' || !Number.isFinite(fee) ||
        !Number.isInteger(fee) || fee < 0) {
      errors.push(
        sheetName + ': ' + grade + '級の参加費が不正です（'
        + rowNumber + '行目B列）。0以上の整数を入力してください。'
      );
      return;
    }
    fees[grade] = fee;
  });
  return { fees: fees, errors: errors };
}

function taikaiGradeFeesFromStructure_(structure, sheetName, grades) {
  if (structure.version !== 2) {
    return taikaiGradeFeesFromSheetData_(
      structure.data, sheetName, grades
    );
  }
  const fees = {};
  const errors = [];
  (grades || []).forEach(grade => {
    const fee = tournamentSheetGradeFee_(structure, grade);
    if (typeof fee !== 'number' || !Number.isFinite(fee)
        || !Number.isInteger(fee) || fee < 0) {
      errors.push(
        sheetName + ': ' + grade
        + '級の参加費が不正です。0以上の整数を入力してください。'
      );
      return;
    }
    fees[grade] = fee;
  });
  return { fees: fees, errors: errors };
}

function taikaiSyncTournamentSchedulesFromSheet_(sheetName, gradeDates) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const calendar = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
  const tournamentSheet = ss.getSheetByName(sheetName);
  if (!calendar || !tournamentSheet) throw new Error('大会情報シートが見つかりません。');

  const baseName = tournamentSheetBaseName_(sheetName);
  const tournament = taikaiEnsureTournament_(baseName);
  const calendarRows = calendar.getRange(1, 1, calendar.getLastRow(), 16).getValues();
  const calendarRow = calendarRows.slice(2).find(row => String(row[0]) === sheetName);
  if (!calendarRow) throw new Error('カレンダーに大会情報がありません: ' + sheetName);

  const structure = tournamentSheetStructure_(tournamentSheet, true);
  const tournamentData = structure.data;
  const isSanctioned = tournamentSheetIsSanctioned_(structure);
  const applicationDeadline = structure.version === 2
    ? taikaiFormatDate_(
      tournamentSheetManagementValue_(structure, '本申込期限')
    )
    : taikaiFormatDate_(calendarRow[5]);
  if (!applicationDeadline) throw new Error('申込期限が未設定のため、APIへ日程を登録できません。');
  const lotteryDate = structure.version === 2
    ? (taikaiFormatDate_(
      tournamentSheetManagementValue_(structure, '抽選日')
    ) || null)
    : (taikaiFormatDate_(calendarRow[7]) || null);
  const targetGrades = Object.keys(gradeDates || {}).filter(grade =>
    Boolean(taikaiFormatDate_(gradeDates[grade]))
  );
  const declaredGrades = tournamentSheetDeclaredGrades_(sheetName);
  const outOfScopeGrades = targetGrades.filter(
    grade => !declaredGrades.includes(grade)
  );
  if (outOfScopeGrades.length) {
    throw new Error(
      'このフォームの担当外の級は同期できません: ' + outOfScopeGrades.join(',')
    );
  }
  const feeResult = taikaiGradeFeesFromStructure_(
    structure, sheetName, targetGrades
  );
  if (feeResult.errors.length) throw new Error(feeResult.errors.join('\n'));
  const paymentInstructions = tournamentSheetPaymentInstructionsFromStructure_(structure);
  let internalPaymentDeadlineIndex = null;
  try {
    internalPaymentDeadlineIndex = fiscalSyncCalendarColumn_(
      calendarRows.slice(0, 2), '振込開始'
    );
  } catch (e) {
    // API側の段階的な配置中も既存の日程同期を止めない。
    // 完全同期のプレビューでは検証エラーとして明示される。
  }
  const internalPaymentDeadline = internalPaymentDeadlineIndex === null
    ? null
    : fiscalSyncOptionalDate_(calendarRow[internalPaymentDeadlineIndex]);
  if (internalPaymentDeadline && !internalPaymentDeadline.valid) {
    throw new Error(
      '会内振込期限を日付として認識できません: '
      + String(calendarRow[internalPaymentDeadlineIndex])
    );
  }

  targetGrades.forEach(grade => {
    const heldOn = taikaiFormatDate_(gradeDates[grade]);
    const paymentSetting = structure.version === 2
      ? tournamentSheetManagementValue_(structure, '本振込期限')
      : calendarRow[10];
    const payment = taikaiPaymentSchedule_(paymentSetting, heldOn);
    if (structure.version === 2) {
      payment.payment_timing =
        tournamentSheetPaymentTimingFromStructure_(structure);
    }
    const schedules = taikaiApiRequest_('GET', '/schedules', null, {
      tournament_id: tournament.id,
      grade: grade,
      held_on: heldOn,
    }) || [];
    const body = {
      held_on: heldOn,
      grade: grade,
      application_deadline: applicationDeadline,
      payment_deadline: payment.payment_deadline,
      lottery_result_date: lotteryDate,
      payment_timing: payment.payment_timing,
      participation_fee_yen: feeResult.fees[grade],
      is_sanctioned: isSanctioned,
      payment_instructions: paymentInstructions,
    };
    if (structure.version !== 2) {
      body.venue = null;
      body.reception_ends_at = null;
    }
    if (structure.version !== 2 && internalPaymentDeadline) {
      body.internal_payment_deadline = internalPaymentDeadline.value;
    }
    let savedSchedule;
    if (schedules.length === 0) {
      savedSchedule = taikaiApiRequest_(
        'POST',
        '/tournaments/' + encodeURIComponent(String(tournament.id)) + '/schedules',
        body
      );
    } else if (schedules.length === 1) {
      savedSchedule = taikaiApiRequest_(
        'PATCH', '/schedules/' + encodeURIComponent(String(schedules[0].id)), body
      );
    } else {
      throw new Error('同じ大会・級・開催日のAPI日程が複数あります: ' + sheetName + ' ' + grade);
    }
    if (structure.version === 2) {
      tournamentSheetGradeFeeRange_(tournamentSheet, structure, grade)
        .setValue(savedSchedule.participation_fee_yen);
      tournamentSheetGradeDateRange_(tournamentSheet, structure, grade)
        .setValue(savedSchedule.held_on);
    }
  });
  return tournament;
}
