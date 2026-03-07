// ============================================================
// 掲示板
// ============================================================
// シート「掲示板」構造:
//   row 1: ヘッダー（ID / 日時 / 役職 / 投稿者名 / 内容）
//   row 2+: データ

var BOARD_SHEET_NAME_  = '掲示板';
var BOARD_SPREADSHEET_ID_ = '1FmndV7ZmlKKMDU4ElExHHZWqVrDHRf4Veg7K4QrJDzg';

function getBoardPosts() {
  try {
    const ss    = SpreadsheetApp.openById(BOARD_SPREADSHEET_ID_);
    let sheet   = ss.getSheetByName(BOARD_SHEET_NAME_);
    if (!sheet) {
      // シートがなければ作成
      sheet = ss.insertSheet(BOARD_SHEET_NAME_);
      sheet.getRange(1, 1, 1, 5).setValues([['ID', '日時', '役職', '投稿者名', '内容']]);
      return JSON.stringify({ ok: true, posts: [] });
    }
    const last = sheet.getLastRow();
    if (last < 2) return JSON.stringify({ ok: true, posts: [] });

    const data = sheet.getRange(2, 1, last - 1, 5).getValues();
    const posts = data
      .filter(r => r[0] !== '')
      .map(r => ({
        id:      String(r[0]),
        datetime: r[1] instanceof Date
          ? Utilities.formatDate(r[1], 'JST', 'yyyy/MM/dd HH:mm')
          : String(r[1] || ''),
        role:    String(r[2]),
        name:    String(r[3]),
        content: String(r[4]),
      }))
      .reverse(); // 新着順

    return JSON.stringify({ ok: true, posts });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

function addBoardPost(json) {
  try {
    const { content, role } = JSON.parse(json);
    if (!content || !content.trim()) return JSON.stringify({ error: '内容が空です' });

    const ss    = SpreadsheetApp.openById(BOARD_SPREADSHEET_ID_);
    let sheet   = ss.getSheetByName(BOARD_SHEET_NAME_);
    if (!sheet) {
      sheet = ss.insertSheet(BOARD_SHEET_NAME_);
      sheet.getRange(1, 1, 1, 5).setValues([['ID', '日時', '役職', '投稿者名', '内容']]);
    }

    // 役職者名を PropertiesService から取得
    const namesStored = PropertiesService.getScriptProperties().getProperty('ROLE_NAMES');
    const names = namesStored ? JSON.parse(namesStored) : {};
    const roleNameMap = { 会長: names['会長'] || '', 副会長: names['副会長'] || '', 申込係: names['申込係'] || '', 振込係: names['振込係'] || '', 案内係: names['案内係'] || '' };
    const posterName = roleNameMap[role] || role;

    const now     = new Date();
    const id      = String(now.getTime());
    const datetimeStr = Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm');
    sheet.appendRow([id, datetimeStr, role, posterName, content.trim()]);

    return JSON.stringify({ ok: true, id, datetime: datetimeStr, role, name: posterName });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// ============================================================
// ヘルプ（掲示板スプレッドシートの「ヘルプ」シートで管理）
// ============================================================
// シート構造: ページキー | ページタイトル | セクション見出し | 内容
var HELP_SHEET_NAME_    = 'ヘルプ';
var HELP_DEFAULT_ROWS_  = [
  ['calendar',        '大会一覧',       '「完了」ボタンについて',     '「完了」を押すと大会が一覧から非表示になり、メール管理・出納管理の関連データも整理されます。'],
  ['calendar',        '大会一覧',       '「完了」ボタンについて',     '詳細な情報（参加者詳細など）を確認したい場合は、元のスプレッドシートを直接参照してください。'],
  ['player',          '選手大会照会',   '出場回数について',           'その年度に出場した公認大会の数をデータベースに記録しています。'],
  ['player',          '選手大会照会',   '出場回数について',           '「公認大会」の定義は「全日本かるた協会公認（主催は除く）」の大会です。全日本選手権・名人戦・新春大会などは含まれません（定義は変わる可能性があります）。'],
  ['player',          '選手大会照会',   '出場回数について',           '出場回数の対象は「その大会の申込開始日までに、出場する権利を有した公認大会」の数です。'],
  ['form-create',     'フォーム作成',   '大会名の入力',               '一意性を保つため「第〇回」を入れることをお勧めします。'],
  ['form-create',     'フォーム作成',   '大会名の入力',               '名前は長くなくて構いません。関係者が判別できる程度に短くまとめてください。'],
  ['form-create',     'フォーム作成',   '級の選択',                   '申込期限が異なる場合は、級ごとにフォームを分けて作成してください。'],
  ['form-create',     'フォーム作成',   '質問項目の追加',             '後から項目を付け足すとバグを引き起こす可能性があります。なるべく最初にすべて追加してください。'],
  ['make-email',      '案内作成',       '参加表明締め切りについて',   '設定した締め切り日の前日に、リマインダーが自動送信される設定になっています。'],
  ['mail-management', 'メール管理',     'リマインダーをすぐ送りたい場合', '「設定」→「再設定」を押し、表示されるプレビューをコピーして手動で送信してください。'],
  ['suitou',          '出納管理',       '使い方',                     '大会詳細で「済」を入力すると参加費がプラス計上されます。'],
  ['suitou',          '出納管理',       '使い方',                     '大会一覧で「振込済」を入力すると参加費がマイナス計上され、大会への支払い完了を記録します。'],
  ['suitou',          '出納管理',       '使い方',                     '「シートから再計算」を使うと、振込済みでない大会の参加費をすべて集計し直します。'],
  ['board',           '掲示板',         '使い方',                     'メンバー全員への連絡事項や共有事項を投稿できます。'],
  ['board',           '掲示板',         '使い方',                     '「大会を挿入」ボタンで大会名を本文中に引用できます。'],
  ['board',           '掲示板',         '使い方',                     '削除は会長副会長（全件）または同じ役職のメンバー（自分の投稿のみ）が可能です。'],
];

function getHelpContent() {
  try {
    const ss  = SpreadsheetApp.openById(BOARD_SPREADSHEET_ID_);
    let sheet = ss.getSheetByName(HELP_SHEET_NAME_);

    // シートがなければ作成してデフォルトデータを投入
    if (!sheet) {
      sheet = ss.insertSheet(HELP_SHEET_NAME_);
      sheet.getRange(1, 1, 1, 4).setValues([['ページキー', 'ページタイトル', 'セクション見出し', '内容']]);
      if (HELP_DEFAULT_ROWS_.length) {
        sheet.getRange(2, 1, HELP_DEFAULT_ROWS_.length, 4).setValues(HELP_DEFAULT_ROWS_);
      }
    }

    const last = sheet.getLastRow();
    if (last < 2) return JSON.stringify({ ok: true, content: {} });

    const data = sheet.getRange(2, 1, last - 1, 4).getValues();

    // { pageKey: { title, sections: { heading: [items] } } } に集約
    const content = {};
    data.forEach(r => {
      const [pageKey, pageTitle, heading, item] = r.map(String);
      if (!pageKey) return;
      if (!content[pageKey]) content[pageKey] = { title: pageTitle, sections: {} };
      if (!content[pageKey].sections[heading]) content[pageKey].sections[heading] = [];
      if (item) content[pageKey].sections[heading].push(item);
    });

    // sections を配列形式に変換
    Object.keys(content).forEach(key => {
      content[key].sections = Object.entries(content[key].sections)
        .map(([heading, items]) => ({ heading, items }));
    });

    return JSON.stringify({ ok: true, content });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

function deleteBoardPost(postId) {
  try {
    const ss    = SpreadsheetApp.openById(BOARD_SPREADSHEET_ID_);
    const sheet = ss.getSheetByName(BOARD_SHEET_NAME_);
    if (!sheet) return JSON.stringify({ error: 'シートが見つかりません' });

    const last = sheet.getLastRow();
    if (last < 2) return JSON.stringify({ error: '投稿が見つかりません' });

    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(postId)) {
        sheet.deleteRow(i + 2);
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ error: '投稿が見つかりません' });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}
