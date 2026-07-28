// ============================================================
// 出納管理・残高記録 更新・取得
// ============================================================

function suitouFiscalYear_() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, 'JST', 'yyyy'));
  const month = Number(Utilities.formatDate(now, 'JST', 'M'));
  return month < 4 ? year - 1 : year;
}

function suitouApiTimestamp_() {
  return Utilities.formatDate(new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function suitouLedger_() {
  const fiscalYear = suitouFiscalYear_();
  let cursor = null;
  let pages = 0;
  const entries = {};
  const payments = {};
  const deposits = {};
  const depositBalances = {};
  do {
    const page = taikaiApiRequest_(
      'GET', '/admin/accounting-ledger', null, {
        fiscal_year: fiscalYear,
        limit: 500,
        cursor: cursor,
      }, {
        operation: '出納管理の読込',
        outcome: '出納一覧を表示できませんでした',
      }
    ) || {};
    (page.entries || []).forEach(item => {
      entries[String(item.entry_id)] = item;
    });
    (page.payments || []).forEach(item => {
      payments[String(item.id)] = item;
    });
    (page.deposit_transactions || []).forEach(item => {
      deposits[String(item.id)] = item;
    });
    (page.deposit_balances || []).forEach(item => {
      depositBalances[String(item.player_id)] = Number(item.balance_yen || 0);
    });
    cursor = page.page && page.page.has_more
      ? String(page.page.next_cursor || '') : '';
    pages++;
    if (pages > 100) throw new Error('出納一覧のページ数が上限を超えました。');
  } while (cursor);
  return {
    fiscal_year: fiscalYear,
    entries: Object.keys(entries).map(id => entries[id]),
    payments: Object.keys(payments).map(id => payments[id]),
    deposits: Object.keys(deposits).map(id => deposits[id]),
    deposit_balances: depositBalances,
  };
}

function suitouDepositSignedAmount_(transaction) {
  const amount = Number(transaction.amount_yen || 0);
  if (transaction.transaction_type === 'deposit_received'
      || transaction.transaction_type === 'overpayment_to_deposit'
      || (transaction.transaction_type === 'adjustment'
          && transaction.direction === 'credit')) {
    return amount;
  }
  return -amount;
}

// DBの支払い履歴・デポジット取引を正本として取得する。
function getSuitouSheets() {
  try {
    const ledger = suitouLedger_();
    const entryById = {};
    ledger.entries.forEach(entry => {
      entryById[String(entry.entry_id)] = entry;
    });
    const transactions = [];
    ledger.payments.forEach(payment => {
      const entry = entryById[String(payment.entry_id)];
      if (!entry) return;
      transactions.push({
        id: String(payment.id),
        type: 'payment',
        player_id: String(entry.player.id),
        entry_id: String(entry.entry_id),
        name: String(entry.player.name || ''),
        amount: Number(payment.amount_yen || 0),
        reason: String(entry.tournament.name || '') + entry.grade + '級 参加費'
          + '（' + String(payment.method || '') + '）',
        date: String(payment.paid_at || ''),
        note: String(payment.note || ''),
        reversed_at: payment.reversed_at || null,
        reversible: String(payment.method || '') !== 'deposit',
      });
    });
    const playerNames = {};
    ledger.entries.forEach(entry => {
      playerNames[String(entry.player.id)] = String(entry.player.name || '');
    });
    ledger.deposits.forEach(deposit => {
      transactions.push({
        id: String(deposit.id),
        type: 'deposit',
        player_id: String(deposit.player_id),
        entry_id: deposit.entry_id ? String(deposit.entry_id) : null,
        name: playerNames[String(deposit.player_id)] || '選手ID ' + deposit.player_id,
        amount: suitouDepositSignedAmount_(deposit),
        reason: 'デポジット（' + String(deposit.transaction_type || '') + '）',
        date: String(deposit.occurred_at || ''),
        note: String(deposit.note || ''),
        reversed_at: deposit.reversed_at || null,
        reversible: String(deposit.transaction_type || '') !== 'deposit_applied',
      });
    });
    transactions.sort((left, right) =>
      String(right.date).localeCompare(String(left.date))
      || Number(right.id) - Number(left.id)
    );
    const entries = ledger.entries.map(entry => ({
      entry_id: String(entry.entry_id),
      player_id: String(entry.player.id),
      name: String(entry.player.name || ''),
      tournament: String(entry.tournament.name || ''),
      grade: String(entry.grade || ''),
      held_on: String(entry.held_on || ''),
      fee: entry.participation_fee_yen === null
        ? null : Number(entry.participation_fee_yen),
      paid: Number(entry.paid_yen || 0),
      balance: entry.balance_yen === null ? null : Number(entry.balance_yen),
      status: String(entry.status || ''),
      deposit_balance: Number(
        ledger.deposit_balances[String(entry.player.id)] || 0
      ),
      canceled_at: entry.canceled_at || null,
    }));
    return JSON.stringify({
      ok: true,
      fiscalYear: ledger.fiscal_year,
      transactions: transactions,
      entries: entries,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function addSuitouApiTransaction(json) {
  try {
    const input = JSON.parse(json);
    const operation = String(input.operation || '');
    const entryId = String(input.entry_id || '');
    const playerId = String(input.player_id || '');
    const amount = Number(input.amount_yen);
    const key = String(input.idempotency_key || '');
    const note = String(input.note || '').trim() || null;
    if (!Number.isInteger(amount) || amount <= 0) {
      throw new Error('金額は1円以上の整数で入力してください。');
    }
    if (!/^[A-Za-z0-9_-]{16,128}$/.test(key)) {
      throw new Error('再送防止キーが不正です。');
    }
    if (note && note.length > 1000) {
      throw new Error('メモは1000文字以内で入力してください。');
    }
    const timestamp = suitouApiTimestamp_();
    let result;
    if (operation === 'payment') {
      if (!/^\d+$/.test(entryId)) throw new Error('申込IDが不正です。');
      const method = String(input.method || '');
      if (!['bank_transfer', 'cash', 'adjustment'].includes(method)) {
        throw new Error('支払い方法が不正です。');
      }
      result = taikaiApiRequest_(
        'POST', '/entries/' + encodeURIComponent(entryId) + '/payments', {
          idempotency_key: key,
          amount_yen: amount,
          method: method,
          paid_at: timestamp,
          note: note,
        }
      );
    } else if (operation === 'deposit_application') {
      if (!/^\d+$/.test(entryId)) throw new Error('申込IDが不正です。');
      result = taikaiApiRequest_(
        'POST', '/entries/' + encodeURIComponent(entryId) + '/deposit-application', {
          idempotency_key: key,
          amount_yen: amount,
          paid_at: timestamp,
          note: note,
        }
      );
    } else if (operation === 'deposit_received'
        || operation === 'deposit_refunded') {
      if (!/^\d+$/.test(playerId)) throw new Error('選手IDが不正です。');
      result = taikaiApiRequest_(
        'POST', '/players/' + encodeURIComponent(playerId) + '/deposits', {
          idempotency_key: key,
          amount_yen: amount,
          transaction_type: operation,
          occurred_at: timestamp,
          note: note,
        }
      );
    } else {
      throw new Error('操作が不正です。');
    }
    if (result && result.entry_id) {
      try {
        refreshTournamentSheetV2ByEntryId_(String(result.entry_id));
      } catch (writebackError) {
        return JSON.stringify({
          error: '出納APIへの登録は成功しましたが、大会シートの書戻しに失敗しました。'
            + '同じ再送防止キーで再実行してください: ' + writebackError.message,
          partial: true,
          result: result,
        });
      }
    }
    return JSON.stringify({ ok: true, result: result });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function reverseSuitouApiTransaction(type, id, reversedAt) {
  try {
    const normalizedType = String(type || '');
    const normalizedId = String(id || '');
    const timestamp = String(reversedAt || '');
    if (!/^\d+$/.test(normalizedId)) throw new Error('履歴IDが不正です。');
    if (!timestamp) throw new Error('取消日時がありません。');
    const path = normalizedType === 'payment'
      ? '/entry-payments/' + encodeURIComponent(normalizedId) + '/reverse'
      : normalizedType === 'deposit'
        ? '/deposit-transactions/' + encodeURIComponent(normalizedId) + '/reverse'
        : '';
    if (!path) throw new Error('履歴種別が不正です。');
    const result = taikaiApiRequest_(
      'POST', path, { reversed_at: timestamp }
    );
    if (result && result.entry_id) {
      try {
        refreshTournamentSheetV2ByEntryId_(String(result.entry_id));
      } catch (writebackError) {
        return JSON.stringify({
          error: '取消APIへの登録は成功しましたが、大会シートの書戻しに失敗しました。'
            + '同じ取消日時で再実行してください: ' + writebackError.message,
          partial: true,
          result: result,
        });
      }
    }
    return JSON.stringify({ ok: true, result: result });
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



// 指定プレイヤーのデポジット純残高を取得する。
// 使用時はマイナス行を追加するため、正の元行だけでなく正負全行を合算する。
function getPlayerDepositBalance_(rows, playerName) {
  const normalizedName = normalizeName_(playerName);
  return (rows || []).reduce((total, row) => {
    if (normalizeName_(String(row[0])) !== normalizedName
        || String(row[2]).trim() !== 'デポジット') return total;
    const amount = Number(row[1]);
    return total + (Number.isFinite(amount) ? amount : 0);
  }, 0);
}

// 指定プレイヤーの利用可能なデポジット残高を返す
function getPlayerDeposit(playerName) {
  try {
    const player = taikaiFindPlayer_(playerName);
    if (!player) return JSON.stringify({ ok: true, deposit: null });
    const result = taikaiApiRequest_(
      'GET', '/players/' + encodeURIComponent(String(player.id)) + '/deposits'
    ) || {};
    const balance = Number(result.balance_yen || 0);
    return JSON.stringify({
      ok: true,
      deposit: balance > 0 ? { amount: balance } : null,
    });
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

// 名前のスペース正規化（全角スペース → 半角、連続スペース → 1つ）
function normalizeName_(name) {
  return String(name).replace(/　/g, ' ').replace(/ +/g, ' ').trim();
}

// 大会シートの下部セクションから参加費テーブルを取得
function getSuitouFeeMap_(allData, formEndIdx) {
  const feeMap = {};
  const rows = tournamentSheetGradeRows_(allData, formEndIdx, false);
  Object.keys(rows).forEach(grade => {
    const fee = (allData[rows[grade] - 1] || [])[1];
    if (typeof fee === 'number' && fee > 0) feeMap[grade] = fee;
  });
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
  return tournamentSheetResponseEndIndex_(allData);
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

    let structure;
    try {
      structure = tournamentSheetStructure_(sheet, false);
    } catch (e) {
      continue;
    }
    const allData = structure.data;
    const formEndIdx = structure.response_end_index;

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

      const payStatus = String(
        tournamentSheetPaymentStatus_(structure, i + 1) || ''
      ).trim();
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
