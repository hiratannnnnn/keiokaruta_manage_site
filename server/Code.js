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
