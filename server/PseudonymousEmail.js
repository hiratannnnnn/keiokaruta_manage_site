// ============================================================
// DB用疑似メールアドレス
// 実メールアドレスはスプレッドシート内だけに保持し、APIへ送信しない。
// ============================================================

function privateEmailSecret_() {
  const secret = String(
    PropertiesService.getScriptProperties().getProperty('PSEUDO_EMAIL_SECRET') || ''
  );
  if (!secret) {
    throw new Error(
      'PSEUDO_EMAIL_SECRET が設定されていません。実メールアドレスをDBへ送信しないため処理を中止しました。'
    );
  }
  return secret;
}

function normalizePrivateEmail_(email) {
  const normalized = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(normalized)) {
    throw new Error('メールアドレスの形式が正しくありません。');
  }
  return normalized;
}

function isPseudonymousEmail_(email) {
  return /^v1-[0-9a-f]{64}@example\.invalid$/i.test(String(email || '').trim());
}

function pseudonymousEmailFor_(realEmail) {
  const normalized = normalizePrivateEmail_(realEmail);
  if (isPseudonymousEmail_(normalized)) return normalized;
  const signature = Utilities.computeHmacSha256Signature(
    normalized,
    privateEmailSecret_(),
    Utilities.Charset.UTF_8
  );
  const hex = signature.map(value => {
    const unsigned = value < 0 ? value + 256 : value;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
  return 'v1-' + hex + '@example.invalid';
}

function emailMapSheet_(createIfMissing) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EMAIL_MAP);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.EMAIL_MAP);
    sheet.getRange(1, 1, 1, 5).setValues([[
      '実メールアドレス', 'DB用疑似メールアドレス', '氏名', '登録日時', '更新日時',
    ]]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function emailMapRows_() {
  const sheet = emailMapSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
}

function rememberPseudonymousEmail_(realEmail, playerName) {
  const real = normalizePrivateEmail_(realEmail);
  if (isPseudonymousEmail_(real)) {
    throw new Error('対応表には実メールアドレスを登録してください。');
  }
  const pseudo = pseudonymousEmailFor_(real);
  const sheet = emailMapSheet_(true);
  const rows = emailMapRows_();
  let targetRow = 0;
  rows.forEach((row, index) => {
    const rowReal = String(row[0] || '').trim().toLowerCase();
    const rowPseudo = String(row[1] || '').trim().toLowerCase();
    if (rowPseudo === pseudo && rowReal && rowReal !== real) {
      throw new Error('疑似メールアドレスの衝突を検出しました。処理を中止します。');
    }
    if (rowReal === real) targetRow = index + 2;
  });
  const now = new Date();
  if (targetRow) {
    sheet.getRange(targetRow, 2, 1, 4).setValues([[
      pseudo, String(playerName || ''), rows[targetRow - 2][3] || now, now,
    ]]);
  } else {
    sheet.appendRow([real, pseudo, String(playerName || ''), now, now]);
  }
  return pseudo;
}

function realEmailForPseudonym_(pseudoEmail) {
  const pseudo = String(pseudoEmail || '').trim().toLowerCase();
  if (!isPseudonymousEmail_(pseudo)) {
    throw new Error('DBから実メールアドレスが返されました。疑似化移行が完了していません。');
  }
  const matches = emailMapRows_().filter(row =>
    String(row[1] || '').trim().toLowerCase() === pseudo
  );
  if (matches.length !== 1) {
    throw new Error('DBメール対応表から実メールアドレスを一意に特定できません。');
  }
  return normalizePrivateEmail_(matches[0][0]);
}

function rememberSnapshotEmailMappings_(snapshot) {
  const requested = {};
  (snapshot.tournaments || []).forEach(tournament => {
    (tournament.entries || []).forEach(entry => {
      const real = normalizePrivateEmail_(entry.real_email);
      const pseudo = pseudonymousEmailFor_(real);
      requested[real] = {
        pseudo: pseudo,
        name: entry.family_name + ' ' + entry.given_name,
      };
    });
  });
  const realEmails = Object.keys(requested);
  if (!realEmails.length) return;

  // 年度同期では一人ずつシートを読み書きせず、対応表を一括更新する。
  const sheet = emailMapSheet_(true);
  const rows = emailMapRows_();
  const rowByReal = {};
  const realByPseudo = {};
  rows.forEach((row, index) => {
    const real = String(row[0] || '').trim().toLowerCase();
    const pseudo = String(row[1] || '').trim().toLowerCase();
    if (real) rowByReal[real] = index;
    if (pseudo) realByPseudo[pseudo] = real;
  });
  const now = new Date();
  realEmails.forEach(real => {
    const item = requested[real];
    if (realByPseudo[item.pseudo] && realByPseudo[item.pseudo] !== real) {
      throw new Error('疑似メールアドレスの衝突を検出しました。処理を中止します。');
    }
    if (Object.prototype.hasOwnProperty.call(rowByReal, real)) {
      const index = rowByReal[real];
      rows[index] = [real, item.pseudo, item.name, rows[index][3] || now, now];
    } else {
      rows.push([real, item.pseudo, item.name, now, now]);
    }
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}
