// ============================================================
// キャンセル区分の共通判定
// ============================================================

// 内部コード。キャンセル状態は canceled_at と
// lottery_result_date だけから決め、シートの表示文字列は参照しない。
const CANCELLATION_STATUS = {
  NONE: 'none',
  BEFORE: 'before',
  AFTER: 'after',
  UNKNOWN: 'unknown',
};

const CANCELLATION_LABELS = {
  none: 'キャンセルなし',
  before: '公開前キャンセル',
  after: '公開後キャンセル',
  unknown: '判定不能なキャンセル',
};

function cancellationNormalizeExplicitStatus_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return '';
  const aliases = {
    none: 'none',
    active: 'none',
    not_canceled: 'none',
    not_cancelled: 'none',
    'キャンセルなし': 'none',
    before: 'before',
    before_lottery: 'before',
    pre_lottery: 'before',
    '公開前': 'before',
    '公開前キャンセル': 'before',
    after: 'after',
    after_lottery: 'after',
    post_lottery: 'after',
    '公開後': 'after',
    '公開後キャンセル': 'after',
    unknown: 'unknown',
    undetermined: 'unknown',
    canceled_unknown: 'unknown',
    cancelled_unknown: 'unknown',
    '判定不能': 'unknown',
    '判定不能なキャンセル': 'unknown',
  };
  return aliases[normalized] || '';
}

// DBのDATETIMEはタイムゾーンなしで返ることがあるため、その場合は
// DBが保持する日本時間の日付をそのまま採用する。offset付きISOは瞬間を
// 日本時間へ変換してから日付だけを取り出す。
function cancellationJapanDate_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return cancellationDateFromInstant_(value);
  }
  const text = String(value).trim();
  if (!text) return '';
  const dateOnly = text.match(/^(\d{4}-\d{2}-\d{2})(?:$|[ T])/);
  const hasOffset = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(text);
  if (dateOnly && !hasOffset) return dateOnly[1];
  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return '';
  return cancellationDateFromInstant_(parsed);
}

function cancellationDateFromInstant_(date) {
  // JSTは固定UTC+09:00のため、Apps ScriptのUtilitiesに依存せず同じ値を得る。
  return new Date(date.getTime() + 9 * 60 * 60 * 1000)
    .toISOString().slice(0, 10);
}

function cancellationStatusFromRecord_(record) {
  const item = record && typeof record === 'object' ? record : {};
  const canceledAt = item.canceled_at;
  if (canceledAt === null || canceledAt === undefined
      || String(canceledAt).trim() === '') {
    return CANCELLATION_STATUS.NONE;
  }

  // APIが明示的な区分を返す場合はそれを利用する。ただし、
  // 「canceled」のような汎用ラベルは判定済み値ではないため無視する。
  const explicit = cancellationNormalizeExplicitStatus_(
    item.cancellation_status || item.cancellation_category
  );
  if (explicit && explicit !== CANCELLATION_STATUS.NONE) return explicit;

  const canceledDate = cancellationJapanDate_(canceledAt);
  const lotteryDate = cancellationJapanDate_(item.lottery_result_date);
  if (!canceledDate || !lotteryDate) return CANCELLATION_STATUS.UNKNOWN;
  return canceledDate < lotteryDate
    ? CANCELLATION_STATUS.BEFORE : CANCELLATION_STATUS.AFTER;
}

function cancellationLabelFromStatus_(status) {
  return CANCELLATION_LABELS[
    cancellationNormalizeExplicitStatus_(status)
      || CANCELLATION_STATUS.UNKNOWN
  ];
}

function cancellationLabelFromRecord_(record) {
  return cancellationLabelFromStatus_(cancellationStatusFromRecord_(record));
}

// DB/APIと同じ参加実績ルール。判定不能なキャンセルは推測せず対象外。
function cancellationCountsAsParticipant_(record) {
  const status = cancellationStatusFromRecord_(record);
  return status === CANCELLATION_STATUS.NONE
    || status === CANCELLATION_STATUS.AFTER;
}

function cancellationIsCanceled_(record) {
  return cancellationStatusFromRecord_(record) !== CANCELLATION_STATUS.NONE;
}
