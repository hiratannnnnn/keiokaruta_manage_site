// ============================================================
// 出場回数カウント
// ============================================================

// taikai_manage API から選手の出場履歴を取得し filterMatches を適用して返す
// beforeDate : 大会前日（この日より前の試合のみ対象）
// 返り値     : "date：location：raffleDate" 形式の文字列配列
function fetchCountMatches_(playerName, beforeDate) {
  try {
    const y           = beforeDate.getFullYear();
    const fiscalYear  = (beforeDate.getMonth() + 1) < 4 ? y - 1 : y;
    const fiscalStart = new Date(fiscalYear,     3,  1);
    const fiscalEnd   = new Date(fiscalYear + 1, 2, 31);

    return taikaiGetParticipations_(String(playerName).replace(/　/g, ' '), beforeDate)
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

    const structure = tournamentSheetStructure_(sheet, true);
    const data = structure.data;
    const paymentStatusIndex = structure.layout.payment_status_column - 1;
    const formIdIndex = structure.layout.form_id_column - 1;
    const editUrlIndex = structure.layout.edit_url_column - 1;

    const gradeRegex     = /^[A-E]$/;
    const formerDates    = {};
    let   moshikomiStart = '';

    data.slice(structure.response_end_index).forEach(row => {
      if (gradeRegex.test(String(row[0]))) {
        const d = row[paymentStatusIndex];
        if (d instanceof Date) formerDates[row[0]] = new Date(d.getTime() - 24 * 60 * 60 * 1000);
      }
      if (row[formIdIndex] === '申込開始日') moshikomiStart = row[editUrlIndex];
      if (typeof moshikomiStart === 'string' && row[formIdIndex] === 'リマインダー') {
        moshikomiStart = row[editUrlIndex];
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

    for (let i = 1; i < structure.response_end_index; i++) {
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

      sheet.getRange(
        i + 1, structure.layout.form_id_column, 1, 2
      ).setValues([[num, historyList.join(',')]]);

      const bg = (inputVal !== null && Number(inputVal) !== matchesList.length) ? 'yellow' : 'white';
      sheet.getRange(i + 1, structure.layout.form_id_column).setBackground(bg);

    }

    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
