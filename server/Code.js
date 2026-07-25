// ============================================================
// エントリポイント・共通ユーティリティ
// 静的設定値は config.js の CONFIG を参照
// ============================================================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setTitle('かるた部 管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTML 内で <?!= include('ファイル名') ?> として使う
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Date オブジェクトは yyyy/MM/dd に整形、それ以外はそのまま文字列化
function formatCell(val) {
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, 'JST', 'yyyy/MM/dd');
  }
  return String(val === null || val === undefined ? '' : val);
}

// Googleフォーム回答の同一人物重複を、最新回答だけにまとめる。
// 原則メールアドレスで識別し、メールがない旧形式では氏名を使用する。
function latestFormRowsByPlayer_(rows) {
  const latest = {};
  (rows || []).forEach((row, index) => {
    const email = String(row[1] || '').trim().toLowerCase();
    const name = String(row[2] || '').replace(/[ 　]+/g, '').trim().toLowerCase();
    const key = email ? 'email:' + email : 'name:' + name;
    if (!email && !name) return;
    const timestamp = row[0] && typeof row[0].getTime === 'function' && !isNaN(row[0].getTime())
      ? row[0].getTime()
      : -Infinity;
    const current = latest[key];
    if (!current || timestamp > current.timestamp ||
        (timestamp === current.timestamp && index > current.index)) {
      latest[key] = { row: row, timestamp: timestamp, index: index };
    }
  });
  return Object.keys(latest).map(key => latest[key]).sort((a, b) => a.index - b.index)
    .map(item => item.row);
}
