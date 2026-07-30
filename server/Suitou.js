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

function suitouDepositReason_(transaction) {
  const type = String(transaction.transaction_type || '');
  const note = String(transaction.note || '');
  if (type === 'overpayment_to_deposit' && note.includes('キャンセル')) {
    return 'キャンセルによるデポジット振替';
  }
  const labels = {
    deposit_received: 'デポジット受入',
    deposit_refunded: 'デポジット返金',
    overpayment_to_deposit: '過払い分のデポジット振替',
    deposit_applied: '参加費へのデポジット充当',
    adjustment: 'デポジット調整',
  };
  return labels[type] || 'デポジット（' + type + '）';
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
        reason: suitouDepositReason_(deposit),
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
    return JSON.stringify({ ok: true, result: result });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// 指定プレイヤーの利用可能なデポジット残高を返す
function getPlayerDeposit(playerId, playerName) {
  try {
    const normalizedId = String(playerId || '').trim();
    if (normalizedId && !/^\d+$/.test(normalizedId)) {
      throw new Error('選手IDが不正です。');
    }
    const player = normalizedId
      ? { id: normalizedId }
      : taikaiFindPlayer_(playerName);
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

// 級文字列から参加費合計を計算
function calcFeeFromGrade_(gradeStr, feeMap) {
  const grades = String(gradeStr).replace(/級/g, '').replace(/[Ａ-Ｅ]/g, s =>
    String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
  ).split('').filter(g => /^[A-E]$/.test(g));
  let total = 0;
  for (const g of grades) { if (feeMap[g]) total += feeMap[g]; }
  return total;
}
