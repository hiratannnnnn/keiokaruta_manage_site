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
    const subject = (d.title || '') + (d.grades || '') + 'の案内';
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
// 当選判定: col[N+2] が '' / '済' / '繰越'含む / 'くりこし' → 当選
// それ以外の非空文字列（例: '落選'等）→ 落選
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

    const allData = sheet.getDataRange().getValues();
    let N = null;
    for (const c of allData[0]) {
      if (typeof c === 'number' && Number.isFinite(c) && c >= 1 && c <= 30) { N = c; break; }
    }
    if (N == null) return JSON.stringify({ ok: false, error: 'N が取得できません' });

    let formEndIdx = 1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0] || '').trim() === '') { formEndIdx = i; break; }
      formEndIdx = i + 1;
    }

    const stats = {}; // { grade: { winnerNames: [], loserNames: [] } }
    for (let i = 1; i < formEndIdx; i++) {
      const name = String(allData[i][2] || '').trim();
      if (!name) continue;
      const grade = String(allData[i][4] || '').trim()
        .replace(/[Ａ-Ｅ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
      if (!/^[A-E]$/.test(grade)) continue;
      if (!stats[grade]) stats[grade] = { winnerNames: [], loserNames: [] };
      const pay = String(allData[i][N + 2] || '').trim();
      if (pay === '' || pay === '済') {
        stats[grade].winnerNames.push(dupName(name));
      } else if (pay.includes('キャンセル待ち')) {
        stats[grade].loserNames.push(dupName(name));
      }
    }

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
