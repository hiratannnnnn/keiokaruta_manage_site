// ============================================================
// Gmail 下書き作成
// ============================================================

// 設定シート D13（to）・D14（bcc）から既定の宛先を取得
// 注意：振込確認メールは個別送信のため、この関数の戻り値を使用しないこと
function getDefaultRecipients_() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
    if (!sheet) return { to: '', bcc: '' };
    return {
      to:  String(sheet.getRange(13, 4).getValue() || ''),
      bcc: String(sheet.getRange(14, 4).getValue() || ''),
    };
  } catch(e) {
    return { to: '', bcc: '' };
  }
}

// タブ0：案内メール
function createDraft1(json) {
  try {
    const d = JSON.parse(json);
    const { to, bcc } = getDefaultRecipients_();
    const subject = (d.title || '') + (d.grades || '') + '\u3000案内';
    GmailApp.createDraft(to, subject, d.body || '', { bcc, name: '慶應かるた会' });
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// タブ1：振込案内
function createDraft2(json) {
  try {
    const d = JSON.parse(json);
    const { to, bcc } = getDefaultRecipients_();
    const subject = (d.title || '') + (d.grades || '') + '\u3000出場者確定のお知らせ';
    GmailApp.createDraft(to, subject, d.body || '', { bcc, name: '慶應かるた会' });
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// タブ1補助：大会シートから抽選結果（氏名リスト）を取得
// 当選判定は選考状態を参照する。入金状態は当落と混同しない。
// 表示名は dupName ロジック（名簿の重複苗字なら苗字＋名前の頭文字）で短縮
function getLotteryResults(sheetName) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return JSON.stringify({ ok: false, error: `「${sheetName}」シートが見つかりません` });

    // 名簿から重複苗字リストを取得（dupName ロジック）
    const membersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MEMBERS);
    const dupNamesArr = membersSheet
      ? membersSheet.getDataRange().getValues()
          .filter(row => row[2] === '重複')
          .map(row => String(row[1]))
      : [];
    const dupName = (name) => {
      const parts = String(name).replace('　', ' ').split(' ');
      return dupNamesArr.includes(parts[0])
        ? parts[0] + (parts[1] ? parts[1][0] : '')
        : parts[0];
    };

    const structure = tournamentSheetStructure_(sheet, false);
    const records = tournamentSheetResponseRecords_(structure, false);

    const stats = {}; // { grade: { winnerNames: [], loserNames: [] } }
    records.forEach(record => {
      const name = record.name;
      if (!name) return;
      const grade = record.grade
        .replace(/[Ａ-Ｅ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
      if (!/^[A-E]$/.test(grade)) return;
      if (!stats[grade]) stats[grade] = { winnerNames: [], loserNames: [] };
      const selectionStatus = record.selection_status;
      if (selectionStatus === '') {
        stats[grade].winnerNames.push(dupName(name));
      } else if (selectionStatus.includes('キャンセル待ち')) {
        stats[grade].loserNames.push(dupName(name));
      }
    });

    const grades = Object.keys(stats).sort().map(g => ({
      grade:       g,
      winnerNames: stats[g].winnerNames,
      loserNames:  stats[g].loserNames,
    }));
    return JSON.stringify({ ok: true, grades });
  } catch (e) {
    return JSON.stringify({ ok: false, error: e.message });
  }
}

// タブ2：Free
function createDraft3(json) {
  try {
    const d = JSON.parse(json);
    const { to, bcc } = getDefaultRecipients_();
    GmailApp.createDraft(to, '', d.body || '', { bcc, name: '慶應かるた会' });
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// タブ3：読手講習会など
function createDraft4(json) {
  try {
    const d = JSON.parse(json);
    const { to, bcc } = getDefaultRecipients_();
    const subject = (d.subject || '') + '　案内';
    GmailApp.createDraft(to, subject, d.body || '', { bcc, name: '慶應かるた会' });
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

function participantContactGradeList_(grades) {
  if (!Array.isArray(grades)) {
    throw new Error('連絡対象の級を選択してください。');
  }
  const seen = {};
  const result = [];
  grades.forEach(value => {
    const grade = String(value || '').trim().toUpperCase();
    if (!/^[A-E]$/.test(grade)) {
      throw new Error('連絡対象の級が正しくありません。');
    }
    if (!seen[grade]) {
      seen[grade] = true;
      result.push(grade);
    }
  });
  if (!result.length) throw new Error('連絡対象の級を選択してください。');
  return result;
}

function participantContactCandidateData_(tournamentName, selectedGrades) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheets = ss.getSheets().filter(sheet =>
    tournamentSheetBaseName_(sheet.getName()) === tournamentName
    && /[A-E]+級$/.test(sheet.getName())
  );
  const ownerByGrade = {};
  sheets.forEach(sheet => {
    tournamentSheetDeclaredGrades_(sheet.getName()).forEach(grade => {
      if (ownerByGrade[grade]) {
        throw new Error(
          tournamentName + grade + '級の回答シートが複数あります。'
        );
      }
      ownerByGrade[grade] = sheet;
    });
  });

  selectedGrades.forEach(grade => {
    if (!ownerByGrade[grade]) {
      throw new Error(tournamentName + grade + '級の回答シートが見つかりません。');
    }
  });

  const snapshot = tournamentDetailSnapshot_(tournamentName);
  const pseudoByReal = tournamentDetailPseudonymMap_();
  const entryByPlayerGrade = {};
  (snapshot.entries || []).forEach(entry => {
    const email = String(entry.player_email || '').trim().toLowerCase();
    const grade = String(entry.grade || '').trim().toUpperCase();
    if (!email || !/^[A-E]$/.test(grade)) return;
    const key = email + '|' + grade;
    const current = entryByPlayerGrade[key];
    const entryIsActive = !entry.canceled_at;
    const currentIsActive = current && !current.canceled_at;
    if (
      !current
      || (entryIsActive && !currentIsActive)
      || (entryIsActive === currentIsActive
        && taikaiCompareIds_(entry.entry_id, current.entry_id) > 0)
    ) {
      entryByPlayerGrade[key] = entry;
    }
  });
  const parsedSheets = {};
  const candidates = [];
  selectedGrades.forEach(grade => {
    const sheet = ownerByGrade[grade];
    const sheetName = sheet.getName();
    if (!parsedSheets[sheetName]) {
      parsedSheets[sheetName] = tournamentSheetResponseRecords_(
        tournamentSheetStructure_(sheet, false), false
      );
    }
    parsedSheets[sheetName].forEach(record => {
      if (record.grade !== grade) return;
      const realKey = String(record.email || '').trim().toLowerCase();
      const dbEmail = isPseudonymousEmail_(realKey)
        ? realKey : String(pseudoByReal[realKey] || realKey);
      const entry = entryByPlayerGrade[dbEmail + '|' + grade] || null;
      const canceledAt = entry ? entry.canceled_at : null;
      const selectionState = tournamentDetailSelectionDisplay_({
        selection_status: record.selection_status,
        canceled_at: canceledAt,
      });
      const canceled = Boolean(canceledAt)
        || selectionState === 'キャンセル';
      let realEmail = '';
      try {
        realEmail = normalizePrivateEmail_(record.email);
      } catch (e) {
        realEmail = '';
      }
      candidates.push({
        candidateId: String(sheet.getSheetId()) + ':' + record.source_row,
        name: String(record.name || ''),
        grade: grade,
        state: canceled ? 'キャンセル' : selectionState,
        canceled: canceled,
        selectable: Boolean(realEmail),
        _realEmail: realEmail,
      });
    });
  });
  candidates.sort((a, b) =>
    a.grade.localeCompare(b.grade, 'ja')
    || a.name.localeCompare(b.name, 'ja')
  );
  return candidates;
}

function getParticipantContactCandidates(json) {
  try {
    const d = JSON.parse(json);
    const name = String(d.tournamentName || '').trim();
    if (!name) throw new Error('大会を選択してください。');
    const grades = participantContactGradeList_(d.grades);
    const candidates = participantContactCandidateData_(name, grades).map(
      candidate => ({
        candidateId: candidate.candidateId,
        name: candidate.name,
        grade: candidate.grade,
        state: candidate.selectable
          ? candidate.state : candidate.state + '（メールアドレス不備）',
        canceled: candidate.canceled,
        selectable: candidate.selectable,
      })
    );
    return JSON.stringify({ ok: true, candidates: candidates });
  } catch (e) {
    return JSON.stringify({ error: participantContactSafeError_(e.message) });
  }
}

// 選択された候補IDをサーバー側で現在の大会・級と再照合する。
// 実メールアドレスは回答シートから取得し、呼び出し元へは返さない。
function participantContactBcc_(tournamentName, grades, selectedCandidateIds) {
  const name = String(tournamentName || '').trim();
  if (!name) throw new Error('大会を選択してください。');
  const selectedGrades = participantContactGradeList_(grades);
  if (!Array.isArray(selectedCandidateIds) || !selectedCandidateIds.length) {
    throw new Error('送信対象を1人以上選択してください。');
  }
  const selected = {};
  selectedCandidateIds.forEach(value => {
    const candidateId = String(value || '').trim();
    if (!/^\d+:\d+$/.test(candidateId)) {
      throw new Error('送信対象の指定が正しくありません。');
    }
    selected[candidateId] = true;
  });

  const candidates = participantContactCandidateData_(name, selectedGrades);
  const selectedCandidates = candidates.filter(candidate =>
    candidate.selectable && selected[candidate.candidateId]
  );
  const matchedIds = {};
  selectedCandidates.forEach(candidate => {
    matchedIds[candidate.candidateId] = true;
  });
  const requestedIds = Object.keys(selected);
  if (
    selectedCandidates.length === 0
    || requestedIds.some(entryId => !matchedIds[entryId])
  ) {
    throw new Error(
      '送信対象が現在の申込情報と一致しません。対象者を再読み込みしてください。'
    );
  }

  const recipients = [];
  const seenRecipients = {};
  selectedCandidates.forEach(candidate => {
    const real = normalizePrivateEmail_(candidate._realEmail);
    if (seenRecipients[real]) return;
    seenRecipients[real] = true;
    recipients.push(real);
  });
  if (!recipients.length) {
    throw new Error('選択した級に、連絡対象となる申込者がいません。');
  }
  return recipients;
}

function participantContactSafeError_(message) {
  return String(message || '下書きの作成に失敗しました。').replace(
    /[A-Z0-9._%+-]+@([A-Z0-9.-]+\.[A-Z]{2,})/gi,
    '***@$1'
  );
}

// 出場者への連絡：個人情報を画面へ返さず、Gmail下書きのBCCだけに設定する。
function createParticipantContactDraft(json) {
  try {
    const d = JSON.parse(json);
    const tournamentName = String(d.tournamentName || '').trim();
    const body = '';
    const bcc = participantContactBcc_(
      tournamentName, d.grades, d.selectedCandidateIds
    );
    const recipients = getDefaultRecipients_();
    const to = String(recipients.to || '').trim();
    if (!to) {
      throw new Error('設定シートの既定Toアドレスが未設定です。');
    }
    const subject = tournamentName + '\u3000参加者への連絡';
    GmailApp.createDraft(to, subject, body, {
      bcc: bcc.join(','),
      name: '慶應かるた会',
    });
    return JSON.stringify({
      ok: true,
      recipientCount: bcc.length,
      subject: subject,
    });
  } catch(e) {
    return JSON.stringify({ error: participantContactSafeError_(e.message) });
  }
}
