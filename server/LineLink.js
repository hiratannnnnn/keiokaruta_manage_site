// ============================================================
// LINEアカウント連携
// 実メール・LINEユーザーIDは非公開スプレッドシート内だけに保持する。
// ============================================================

const LINE_LINK_HEADERS_ = {
  PLAYER_ID: 'player_id',
  LINE_USER_ID: 'LINEユーザーID',
  LINKED_AT: 'LINE連携日時',
};

function lineLinkProperty_(name) {
  const value = String(
    PropertiesService.getScriptProperties().getProperty(name) || ''
  ).trim();
  if (!value) throw new Error(name + ' が設定されていません。');
  return value;
}

function lineLinkHex_(bytes) {
  return bytes.map(value => {
    const unsigned = value < 0 ? value + 256 : value;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function lineLinkDigest_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
}

function lineLinkSafeEquals_(left, right) {
  const leftBytes = lineLinkDigest_(left);
  const rightBytes = lineLinkDigest_(right);
  let difference = leftBytes.length ^ rightBytes.length;
  for (let index = 0; index < leftBytes.length; index++) {
    difference |= leftBytes[index] ^ (rightBytes[index] || 0);
  }
  return difference === 0;
}

function lineLinkBindingHash_(lineUserId) {
  const normalized = String(lineUserId || '').trim();
  if (!normalized || normalized.length > 255) {
    throw new Error('LINE連携情報が不正です。');
  }
  return lineLinkHex_(Utilities.computeHmacSha256Signature(
    normalized,
    lineLinkProperty_('LINE_LINK_BINDING_SECRET'),
    Utilities.Charset.UTF_8
  ));
}

function lineLinkEmailMap_(createColumns) {
  const sheet = emailMapSheet_(createColumns === true);
  if (!sheet) throw new Error('非公開対応表が見つかりません。');
  const headers = sheet.getRange(
    1, 1, 1, Math.max(sheet.getLastColumn(), 1)
  ).getValues()[0].map(value => String(value || '').trim());
  const required = [
    '実メールアドレス',
    'DB用疑似メールアドレス',
    LINE_LINK_HEADERS_.PLAYER_ID,
    LINE_LINK_HEADERS_.LINE_USER_ID,
    LINE_LINK_HEADERS_.LINKED_AT,
  ];
  required.forEach(header => {
    if (headers.includes(header)) return;
    if (!createColumns) throw new Error('非公開対応表の列が不足しています。');
    headers.push(header);
    sheet.getRange(1, headers.length).setValue(header);
  });
  const indexes = {};
  required.forEach(header => {
    const matches = [];
    headers.forEach((value, index) => {
      if (value === header) matches.push(index);
    });
    if (matches.length !== 1) {
      throw new Error('非公開対応表の列を一意に特定できません。');
    }
    indexes[header] = matches[0];
  });
  return { sheet: sheet, headers: headers, indexes: indexes };
}

function lineLinkRows_(map) {
  if (map.sheet.getLastRow() < 2) return [];
  return map.sheet.getRange(
    2, 1, map.sheet.getLastRow() - 1, map.headers.length
  ).getValues();
}

function lineLinkFindEmailRow_(map, realEmail, pseudoEmail) {
  const realIndex = map.indexes['実メールアドレス'];
  const pseudoIndex = map.indexes['DB用疑似メールアドレス'];
  const matches = [];
  lineLinkRows_(map).forEach((row, index) => {
    const rowReal = String(row[realIndex] || '').trim().toLowerCase();
    const rowPseudo = String(row[pseudoIndex] || '').trim().toLowerCase();
    if (rowReal === realEmail && rowPseudo === pseudoEmail) {
      matches.push({ row_number: index + 2, values: row });
    }
  });
  if (matches.length !== 1) {
    throw new Error('登録情報を一意に特定できません。');
  }
  return matches[0];
}

function lineLinkAssertAvailable_(map, rowNumber, playerId, lineUserId) {
  const playerIndex = map.indexes[LINE_LINK_HEADERS_.PLAYER_ID];
  const lineIndex = map.indexes[LINE_LINK_HEADERS_.LINE_USER_ID];
  lineLinkRows_(map).forEach((row, index) => {
    const currentRow = index + 2;
    const rowPlayerId = String(row[playerIndex] || '').trim();
    const rowLineId = String(row[lineIndex] || '').trim();
    if (rowLineId === lineUserId && currentRow !== rowNumber) {
      throw new Error('このLINEアカウントは既に連携されています。');
    }
    if (currentRow === rowNumber) {
      if (rowPlayerId && rowPlayerId !== String(playerId)) {
        throw new Error('登録済みの選手情報と一致しません。');
      }
      if (rowLineId && rowLineId !== lineUserId) {
        throw new Error('別のLINEアカウントが既に連携されています。');
      }
    }
  });
}

function lineLinkPlayerByEmail_(realEmail) {
  const pseudo = pseudonymousEmailFor_(realEmail);
  const players = taikaiApiRequest_(
    'GET', '/players', null, { email: pseudo }, {
      operation: 'LINE連携の選手確認',
      outcome: 'LINE連携を中断',
    }
  ) || [];
  if (players.length !== 1) {
    throw new Error('登録情報を一意に特定できません。');
  }
  return { player: players[0], pseudo_email: pseudo };
}

function lineLinkStart_(request) {
  const realEmail = normalizePrivateEmail_(request.email);
  const lineUserId = String(request.line_user_id || '').trim();
  const bindingHash = lineLinkBindingHash_(lineUserId);
  const found = lineLinkPlayerByEmail_(realEmail);
  const map = lineLinkEmailMap_(true);
  const emailRow = lineLinkFindEmailRow_(map, realEmail, found.pseudo_email);
  lineLinkAssertAvailable_(
    map, emailRow.row_number, String(found.player.id), lineUserId
  );
  const verification = taikaiApiRequest_(
    'POST', '/line-link-verifications', {
      player_id: String(found.player.id),
      binding_hash: bindingHash,
    }, null, {
      operation: 'LINE連携OTP発行',
      outcome: '認証コードを発行できませんでした',
    }
  );
  const code = String(verification.code || '');
  if (!/^\d{6}$/.test(code)) {
    throw new Error('認証コードを発行できませんでした。');
  }
  MailApp.sendEmail({
    to: realEmail,
    subject: '【慶應かるた会】LINE連携認証コード',
    body: [
      'LINEアカウント連携の認証コードです。',
      '',
      code,
      '',
      '有効期限は10分です。',
      'この操作に心当たりがない場合は、コードを入力せず破棄してください。',
    ].join('\n'),
  });
  return {
    ok: true,
    status: 'otp_sent',
    verification_id: String(verification.verification_id),
    expires_at: String(verification.expires_at || ''),
  };
}

function lineLinkSave_(playerId, lineUserId) {
  const player = taikaiApiRequest_(
    'GET', '/players/' + encodeURIComponent(String(playerId))
  );
  const pseudo = String(player.email || '').trim().toLowerCase();
  if (!isPseudonymousEmail_(pseudo)) {
    throw new Error('選手情報の疑似化が完了していません。');
  }
  const realEmail = realEmailForPseudonym_(pseudo);
  const map = lineLinkEmailMap_(true);
  const emailRow = lineLinkFindEmailRow_(map, realEmail, pseudo);
  lineLinkAssertAvailable_(map, emailRow.row_number, playerId, lineUserId);
  map.sheet.getRange(
    emailRow.row_number, map.indexes[LINE_LINK_HEADERS_.PLAYER_ID] + 1
  ).setValue(String(playerId));
  map.sheet.getRange(
    emailRow.row_number, map.indexes[LINE_LINK_HEADERS_.LINE_USER_ID] + 1
  ).setValue(lineUserId);
  map.sheet.getRange(
    emailRow.row_number, map.indexes[LINE_LINK_HEADERS_.LINKED_AT] + 1
  ).setValue(new Date());
}

function lineLinkVerify_(request) {
  const verificationId = String(request.verification_id || '').trim();
  const code = String(request.code || '').trim();
  const lineUserId = String(request.line_user_id || '').trim();
  if (!/^\d+$/.test(verificationId) || !/^\d{6}$/.test(code)) {
    throw new Error('認証情報が正しくありません。');
  }
  const verified = taikaiApiRequest_(
    'POST',
    '/line-link-verifications/' + encodeURIComponent(verificationId) + '/verify',
    {
      binding_hash: lineLinkBindingHash_(lineUserId),
      code: code,
    },
    null,
    {
      operation: 'LINE連携OTP検証',
      outcome: 'LINE連携を完了できませんでした',
    }
  );
  if (!verified || verified.verified !== true || !verified.player_id) {
    throw new Error('認証情報が正しくありません。');
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    lineLinkSave_(String(verified.player_id), lineUserId);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return { ok: true, status: 'linked' };
}

function lineLinkUnlink_(request) {
  const lineUserId = String(request.line_user_id || '').trim();
  lineLinkBindingHash_(lineUserId);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const map = lineLinkEmailMap_(true);
    const lineIndex = map.indexes[LINE_LINK_HEADERS_.LINE_USER_ID];
    const matches = [];
    lineLinkRows_(map).forEach((row, index) => {
      if (String(row[lineIndex] || '').trim() === lineUserId) {
        matches.push(index + 2);
      }
    });
    if (matches.length > 1) {
      throw new Error('LINE連携情報が重複しています。');
    }
    if (matches.length === 1) {
      map.sheet.getRange(
        matches[0], map.indexes[LINE_LINK_HEADERS_.LINE_USER_ID] + 1
      ).clearContent();
      map.sheet.getRange(
        matches[0], map.indexes[LINE_LINK_HEADERS_.LINKED_AT] + 1
      ).clearContent();
    }
    return { ok: true, status: 'unlinked' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function lineUserIdForPlayer_(playerId) {
  const map = lineLinkEmailMap_(false);
  const playerIndex = map.indexes[LINE_LINK_HEADERS_.PLAYER_ID];
  const lineIndex = map.indexes[LINE_LINK_HEADERS_.LINE_USER_ID];
  const matches = lineLinkRows_(map).filter(row =>
    String(row[playerIndex] || '').trim() === String(playerId)
    && String(row[lineIndex] || '').trim() !== ''
  );
  if (matches.length > 1) {
    throw new Error('LINE連携情報が重複しています。');
  }
  return matches.length ? String(matches[0][lineIndex]).trim() : null;
}

function lineLinkWebhookPayload_(e) {
  const text = e && e.postData ? String(e.postData.contents || '') : '';
  if (!text || text.length > 10000) throw new Error('リクエストが不正です。');
  const parsed = JSON.parse(text);
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error('リクエストが不正です。');
  }
  return parsed;
}

function lineLinkPublicError_(error) {
  const message = String(error && error.message || '');
  if (error && error.http_status === 429) {
    return {
      status: 'rate_limited',
      message: 'しばらく待ってから再試行してください。',
    };
  }
  if (error && error.http_status === 422) {
    return {
      status: 'verification_failed',
      message: '認証コードが無効、期限切れ、または試行上限です。',
    };
  }
  const known = [
    ['認証に失敗しました。', 'unauthorized'],
    ['リクエストが不正です。', 'invalid_request'],
    ['操作が不正です。', 'invalid_action'],
    ['メールアドレスの形式が正しくありません。', 'invalid_email'],
    ['登録情報を一意に特定できません。', 'player_not_found'],
    ['このLINEアカウントは既に連携されています。', 'line_already_linked'],
    ['別のLINEアカウントが既に連携されています。', 'player_already_linked'],
    ['登録済みの選手情報と一致しません。', 'player_conflict'],
    ['認証情報が正しくありません。', 'verification_failed'],
  ];
  const match = known.find(item => message === item[0]);
  return match
    ? { status: match[1], message: match[0] }
    : { status: 'failed', message: '処理に失敗しました。' };
}

function lineLinkWebhookResponse_(e) {
  let result;
  try {
    const request = lineLinkWebhookPayload_(e);
    if (!lineLinkSafeEquals_(
      request.webhook_secret,
      lineLinkProperty_('LINE_LINK_WEBHOOK_SECRET')
    )) {
      throw new Error('認証に失敗しました。');
    }
    if (request.action === 'line_link_start') {
      result = lineLinkStart_(request);
    } else if (request.action === 'line_link_verify') {
      result = lineLinkVerify_(request);
    } else if (request.action === 'line_link_unlink') {
      result = lineLinkUnlink_(request);
    } else {
      throw new Error('操作が不正です。');
    }
  } catch (e) {
    const safe = lineLinkPublicError_(e);
    result = {
      ok: false,
      status: safe.status,
      message: safe.message,
    };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
