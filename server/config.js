// ============================================================
// 静的設定値（定数）
// GASサービスを呼ばない純粋な定数のみ置く → ロード順の影響なし
// ============================================================

const CONFIG = {

  // スプレッドシートID（固定値・非機密なので直接記載）
  SPREADSHEET_ID:       '1KfxkE6RdW-u0AciuqGUMH9P1zeuOkBVa5Y6iTOzy6Hc',
  FORM_FOLDER_TO:       '1uoJJyxsjb63ewkjsv3NzZHwQ1l_DOkf8',
  FORM_TEMPLATE_ID:     '1eDEcwPhpMJu3nDtqoa-zy1t4eQTQE7hzPQ0snL5XRtw',
  TRASH_SPREADSHEET_ID: '1PYFau8GFIaHHq0FQP8yViteyG5ewI08thDsHcbb93Hs',  // ゴミ箱用スプレッドシートID (ids[0])

  // スプレッドシート内のシート名
  SHEET_NAMES: {
    MEMBERS:  '名簿',
    CALENDAR: 'カレンダー',
    MAIL:     'メール管理',
  },

  // 大会結果検索API
  KARUTA_SEARCH_URL: 'http://keiokarutakai.atwebpages.com/search_results2.php',

};
