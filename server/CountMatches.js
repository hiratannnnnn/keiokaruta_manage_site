// ============================================================
// 出場回数カウント
// ============================================================

// 外部 API から選手の出場履歴を取得し filterMatches を適用して返す
// beforeDate : 大会前日（この日より前の試合のみ対象）
// 返り値     : "date：location：raffleDate" 形式の文字列配列
function fetchCountMatches_(playerName, beforeDate) {
  const url = CONFIG.KARUTA_SEARCH_URL
    + '?name=' + encodeURIComponent(playerName.replace('　', ' '))
    + '&date=' + encodeURIComponent(Utilities.formatDate(beforeDate, 'JST', 'yyyy-MM-dd'));
  try {
    const res     = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    const results = JSON.parse(res.getContentText());
    if (!Array.isArray(results)) return [];

    const y           = beforeDate.getFullYear();
    const fiscalYear  = (beforeDate.getMonth() + 1) < 4 ? y - 1 : y;
    const fiscalStart = new Date(fiscalYear,     3,  1);
    const fiscalEnd   = new Date(fiscalYear + 1, 2, 31);

    return results
      .filter(item => {
        const d   = new Date(item.date);
        const loc = String(item.location || '');
        return d >= fiscalStart && d <= fiscalEnd && d < beforeDate
          && !loc.includes('団体') && !loc.includes('職域') && !loc.includes('非公認');
      })
      .map(item => `${item.date}：${item.location}：${item.raffleDate}`);
  } catch (e) {
    return [];
  }
}

// 出場回数を計算し、col[N+3]/col[N+4] に直接書き込む
function runCountMatches(name) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`「${name}」シートが見つかりません`);

    const data  = sheet.getDataRange().getValues();
    const count = data[0].find(c => typeof c === 'number');
    if (count == null) throw new Error('カラム数 (N) が取得できません');

    const gradeRegex     = /^[A-E]$/;
    const formerDates    = {};
    let   moshikomiStart = '';

    sheet.getRange(1, 1, sheet.getLastRow(), count + 5).getValues().forEach(row => {
      if (gradeRegex.test(String(row[0]))) {
        const d = row[count + 2];
        if (d instanceof Date) formerDates[row[0]] = new Date(d.getTime() - 24 * 60 * 60 * 1000);
      }
      if (row[count + 3] === '申込開始日') moshikomiStart = row[count + 4];
      if (typeof moshikomiStart === 'string' && row[count + 3] === 'リマインダー') {
        moshikomiStart = row[count + 4];
      }
    });
    const moshikomiDate = moshikomiStart ? new Date(moshikomiStart) : null;

    // 「公認大会出場回数」列のインデックスを取得
    let kouninCount = 0;
    for (let j = 0; j < data[0].length; j++) {
      if (typeof data[0][j] === 'string' && data[0][j].includes('公認大会出場回数')) {
        kouninCount = j;
      }
    }

    for (let i = 0; i < data.length; i++) {
      if (typeof data[i][2] === 'number') break;
      if (typeof data[i][2] !== 'string' || data[i][2] === '') continue;
      if (!data[i][2].includes(' ') && !data[i][2].includes('　')) continue;

      const grade      = String(data[i][4]);
      const formerDate = formerDates[grade];
      if (!formerDate) continue;

      let matchesList = fetchCountMatches_(String(data[i][2]), formerDate);
      if (moshikomiDate) {
        matchesList = matchesList.filter(item => new Date(item.split('：')[2]) <= moshikomiDate);
      }

      const inputVal = kouninCount > 0 ? data[i][kouninCount] : null;
      const num = (inputVal === null || Number(inputVal) === matchesList.length)
        ? matchesList.length
        : `（入力：${inputVal}）${matchesList.length}`;

      const historyList = matchesList.map(item => {
        const p = item.split('：');
        return p[0] + '：' + p[1];
      });

      sheet.getRange(i + 1, count + 4, 1, 2).setValues([[num, historyList.join(',')]]);

      const bg = (inputVal !== null && Number(inputVal) !== matchesList.length) ? 'yellow' : 'white';
      sheet.getRange(i + 1, count + 4).setBackground(bg);

      if (data[i][count + 3] === '催促メール設定（「☆大会フォーム」から）') {
        sheet.getRange(i, count + 4).setValue('大会開催日までの、申込開始日時点で抽選に通っているもの');
      }
    }

    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
