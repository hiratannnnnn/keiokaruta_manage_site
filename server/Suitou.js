// ============================================================
// 出納管理・残高記録 更新・取得
// ============================================================

// 出納管理・残高記録の内容を取得
// 残高記録は出納管理の全行から動的に計算する
function getSuitouSheets() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 出納管理 (row 7+、row 6 はヘッダー)
    const suitouSheet = ss.getSheetByName('出納管理');
    let transactions = [];
    const balanceMap = {};

    if (suitouSheet && suitouSheet.getLastRow() >= 7) {
      transactions = suitouSheet.getRange(7, 1, suitouSheet.getLastRow() - 6, 4).getValues()
        .map((r, i) => ({ r, rowNum: i + 7 }))
        .filter(({ r }) => r[0] !== '')
        .map(({ r, rowNum }) => ({
          rowNum,
          name:   String(r[0]),
          amount: typeof r[1] === 'number' ? r[1] : Number(r[1]) || 0,
          reason: String(r[2]),
          date:   r[3] instanceof Date
            ? Utilities.formatDate(r[3], 'JST', 'yyyy/MM/dd')
            : String(r[3] || ''),
        }));

      // 残高を出納管理の全行から集計
      for (const t of transactions) {
        if (!balanceMap[t.name]) balanceMap[t.name] = { balance: 0, tournaments: new Set() };
        balanceMap[t.name].balance += t.amount;
        // デポジット以外の参加費行のみ大会名を収集
        if (t.reason !== 'デポジット') {
          const m = t.reason.match(/^(.+?)　参加費$/);
          if (m) balanceMap[t.name].tournaments.add(shortTournamentName_(m[1]));
        }
      }
    }

    const balances = Object.entries(balanceMap).map(([name, d]) => ({
      name,
      balance:     d.balance,
      tournaments: [...d.tournaments].join(', '),
    }));

    return JSON.stringify({ ok: true, transactions, balances });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 出納管理に1行手動追加
function addSuitouRow(json) {
  try {
    const { name, amount, reason } = JSON.parse(json);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('出納管理');
    if (!sheet) throw new Error('「出納管理」シートが見つかりません');
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
    sheet.insertRowAfter(6);
    sheet.getRange(7, 1, 1, 4).setValues([[name, Number(amount), reason, today]]);
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}



// 指定プレイヤーのデポジット行を出納管理から検索して返す
function getPlayerDeposit(playerName) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('出納管理');
    if (!sheet || sheet.getLastRow() < 7) return JSON.stringify({ ok: true, deposit: null });
    const rows = sheet.getRange(7, 1, sheet.getLastRow() - 6, 3).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (normalizeName_(String(rows[i][0])) === normalizeName_(playerName) &&
          String(rows[i][2]).trim() === 'デポジット' &&
          Number(rows[i][1]) > 0) {
        return JSON.stringify({ ok: true, deposit: { rowNum: i + 7, amount: Number(rows[i][1]) } });
      }
    }
    return JSON.stringify({ ok: true, deposit: null });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 出納管理の指定行を削除
function deleteSuitouRow(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('出納管理');
    if (!sheet) throw new Error('「出納管理」シートが見つかりません');
    if (rowNum < 7 || rowNum > sheet.getLastRow()) throw new Error('行番号が不正です');
    sheet.deleteRow(rowNum);
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 出納管理の指定行の事由を「デポジット」に変更
function convertToDeposit(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('出納管理');
    if (!sheet) throw new Error('「出納管理」シートが見つかりません');
    if (rowNum < 7 || rowNum > sheet.getLastRow()) throw new Error('行番号が不正です');
    sheet.getRange(rowNum, 3).setValue('デポジット');
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

const SUITOU_SKIP_SHEETS = new Set([
  '名簿', 'カレンダー', 'メール管理', '出納管理', '残高記録', 'フォーム作成',
]);

// 大会シートかどうかの判定（ヘッダー行に有効な N が存在するか）
function getSuitouN_(headerRow) {
  for (const c of headerRow) {
    if (typeof c === 'number' && Number.isFinite(c) && c >= 1 && c <= 30) return c;
  }
  return null;
}

// 名前のスペース正規化（全角スペース → 半角、連続スペース → 1つ）
function normalizeName_(name) {
  return String(name).replace(/　/g, ' ').replace(/ +/g, ' ').trim();
}

// 大会シートの下部セクションから参加費テーブルを取得
function getSuitouFeeMap_(allData, formEndIdx) {
  const feeMap = {};
  for (let i = formEndIdx; i < allData.length; i++) {
    const c0 = String(allData[i][0] || '').trim();
    if (/^[A-E]$/.test(c0)) {
      const fee = allData[i][1];
      if (typeof fee === 'number' && fee > 0) feeMap[c0] = fee;
    }
  }
  return feeMap;
}

// 級文字列から参加費合計を計算
function calcFeeFromGrade_(gradeStr, feeMap) {
  const grades = String(gradeStr).replace(/級/g, '').replace(/[Ａ-Ｅ]/g, s =>
    String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
  ).split('').filter(g => /^[A-E]$/.test(g));
  let total = 0;
  for (const g of grades) { if (feeMap[g]) total += feeMap[g]; }
  return total;
}

// 出納管理シートにトランザクション行を1行追加（同名・同事由・同符号が既存なら追加しない）
function appendSuitouTx_(ss, name, amount, reason) {
  const sheet = ss.getSheetByName('出納管理');
  if (!sheet) return;
  if (sheet.getLastRow() >= 7) {
    const existing = sheet.getRange(7, 1, sheet.getLastRow() - 6, 3).getValues();
    for (const row of existing) {
      if (String(row[0]) === name && String(row[2]) === reason &&
          Math.sign(Number(row[1])) === Math.sign(amount)) return;
    }
  }
  const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  sheet.insertRowAfter(6);
  sheet.getRange(7, 1, 1, 4).setValues([[name, amount, reason, today]]);
}

// 出納管理シートから指定事由・マイナス額の行を削除（振込済み取り消し用）
function removeSuitouNegTxByReason_(ss, reason) {
  const sheet = ss.getSheetByName('出納管理');
  if (!sheet || sheet.getLastRow() < 7) return;
  const data     = sheet.getRange(7, 1, sheet.getLastRow() - 6, 3).getValues();
  const toDelete = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][2]) === reason && Number(data[i][1]) < 0) toDelete.push(i + 7);
  }
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
}

// 大会シートの参加者終端行インデックスを取得
function getSuitouFormEndIdx_(allData) {
  let formEndIdx = 1;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][0] || '').trim() === '') { formEndIdx = i; break; }
    formEndIdx = i + 1;
  }
  return formEndIdx;
}

// 大会名の短縮（残高記録の大会欄用）
function shortTournamentName_(name) {
  return name
    .replace(/第\d+回/, '')
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/記念/, '記')
    .slice(0, 4);
}

function updateSuitou() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const suitouSheet  = ss.getSheetByName('出納管理');
  const balanceSheet = ss.getSheetByName('残高記録');
  if (!suitouSheet)  throw new Error('「出納管理」シートが見つかりません');
  if (!balanceSheet) throw new Error('「残高記録」シートが見つかりません');

  const today = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');

  // トランザクション: [name, amount, reason, date]
  const transactions = [];
  // 人別集計: { [name]: { balance: number, tournaments: string[] } }
  const personMap = {};

  // カレンダーで振込済み（payDone=済）の大会はスキップ
  const payDoneNames = new Set();
  const calSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
  if (calSheet && calSheet.getLastRow() >= 3) {
    calSheet.getRange(3, 1, calSheet.getLastRow() - 2, 12).getValues().forEach(r => {
      if (r[0] !== '' && String(r[11]) === '済') payDoneNames.add(String(r[0]));
    });
  }

  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    if (sheet.isSheetHidden()) continue;
    const sheetName = sheet.getName();
    if (SUITOU_SKIP_SHEETS.has(sheetName)) continue;
    if (payDoneNames.has(sheetName)) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 5) continue;

    // ヘッダー行から N を取得
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const N = getSuitouN_(headerRow);
    if (N == null) continue;

    // 全データ取得（最大 N+5 列）
    const readCols = Math.min(lastCol, N + 5);
    const allData  = sheet.getRange(1, 1, lastRow, readCols).getValues();

    // 参加者行の終端を探す（col A が空になった行）
    let formEndIdx = 1;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0] || '').trim() === '') {
        formEndIdx = i;
        break;
      }
      formEndIdx = i + 1;
    }

    // 下部セクションから参加費テーブルを探す
    // col A (index 0) = A〜E, col B (index 1) = 金額（正の数）
    const feeMap = {}; // { 'A': 2500, 'B': 2500, ... }
    for (let i = formEndIdx; i < allData.length; i++) {
      const c0 = String(allData[i][0] || '').trim();
      if (/^[A-E]$/.test(c0)) {
        const fee = allData[i][1];
        if (typeof fee === 'number' && fee > 0) {
          feeMap[c0] = fee;
        }
      }
    }
    // feeMap が空の場合はこのシートはスキップ
    if (Object.keys(feeMap).length === 0) continue;

    // 参加者行を処理（row 2 以降、formEndIdx まで）
    for (let i = 1; i < formEndIdx; i++) {
      // col 3 (1-indexed) = index 2 = 氏名
      const nameRaw = String(allData[i][2] || '').trim();
      if (!nameRaw) continue;
      const name = normalizeName_(nameRaw);
      if (!name) continue;

      // col 5 (1-indexed) = index 4 = 級
      const gradeStr = String(allData[i][4] || '').trim();
      if (!gradeStr) continue;

      // col N+3 (1-indexed) = index N+2 = 振込み済みか
      const payStatus = String(allData[i][N + 2] || '').trim();
      const isPaid = payStatus === '済' || payStatus === '繰越' || payStatus === 'くりこし';
      if (!isPaid) continue;

      // 参加級に応じた参加費を算出
      const grades = gradeStr.replace(/級/g, '').replace(/[Ａ-Ｅ]/g, s =>
        String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
      ).split('').filter(g => /^[A-E]$/.test(g));

      let totalFee = 0;
      for (const g of grades) {
        if (feeMap[g]) totalFee += feeMap[g];
      }
      if (totalFee === 0) continue;

      const reason = sheetName + '　参加費';
      transactions.push([name, totalFee, reason, today]);

      if (!personMap[name]) personMap[name] = { balance: 0, tournaments: [] };
      personMap[name].balance += totalFee;
      personMap[name].tournaments.push(shortTournamentName_(sheetName));
    }
  }

  // 出納管理シート: 7行目以降を書き直し（デポジット行は保持）
  const lastSuitouRow = suitouSheet.getLastRow();
  // 既存のデポジット行を退避
  const depositRows = [];
  if (lastSuitouRow >= 7) {
    suitouSheet.getRange(7, 1, lastSuitouRow - 6, 4).getValues().forEach(r => {
      if (String(r[2]).trim() === 'デポジット') depositRows.push(r);
    });
    suitouSheet.getRange(7, 1, lastSuitouRow - 6, 4).clearContent();
  }
  // デポジット行 → 新規トランザクションの順に書き込む
  const allRows = [...depositRows, ...transactions];
  if (allRows.length > 0) {
    suitouSheet.getRange(7, 1, allRows.length, 4).setValues(allRows);
  }

  // 残高記録: 出納管理の全行（書き込み後）を読んで集計
  const balanceMap = {};
  const writtenLastRow = suitouSheet.getLastRow();
  if (writtenLastRow >= 7) {
    suitouSheet.getRange(7, 1, writtenLastRow - 6, 3).getValues().forEach(row => {
      const n = String(row[0] || '').trim();
      if (!n) return;
      const amt    = typeof row[1] === 'number' ? row[1] : Number(row[1]) || 0;
      const reason = String(row[2] || '');
      if (!balanceMap[n]) balanceMap[n] = { balance: 0, tournaments: new Set() };
      balanceMap[n].balance += amt;
      const m = reason.match(/^(.+?)　参加費$/);
      if (m) balanceMap[n].tournaments.add(shortTournamentName_(m[1]));
    });
  }

  const persons = Object.entries(balanceMap).map(([name, d]) => [
    name,
    d.balance,
    [...d.tournaments].join(', '),
  ]);
  const lastBalanceRow = balanceSheet.getLastRow();
  if (lastBalanceRow >= 1) {
    balanceSheet.getRange(1, 1, lastBalanceRow, 3).clearContent();
  }
  if (persons.length > 0) {
    balanceSheet.getRange(1, 1, persons.length, 3).setValues(persons);
  }

  Logger.log('updateSuitou 完了: トランザクション ' + transactions.length + ' 件, ' + persons.length + ' 人');
  return JSON.stringify({ ok: true, transactionCount: transactions.length, personCount: persons.length });
}
