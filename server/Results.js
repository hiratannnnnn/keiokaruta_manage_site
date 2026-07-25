// ============================================================
// 出場大会履歴検索（taikai_manage API）
// ============================================================

// name   : 選手名（全角スペースは半角に変換）
// dateStr: 日付文字列 yyyy-MM-dd（省略可）
function getTournamentResults(name, dateStr) {
  const date = dateStr || Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  try {
    const results = taikaiGetParticipations_(String(name).replace(/　/g, ' '), new Date(date + 'T23:59:59+09:00'));
    return JSON.stringify({
      url:    'taikai_manage:/players/{id}/participations',
      status: 200,
      body:   JSON.stringify(results.map(item => ({
        date: item.date,
        location: item.location,
        raffleDate: item.raffleDate,
        isOfficial: item.isOfficial,
      }))),
    });
  } catch (e) {
    return JSON.stringify({ url: 'taikai_manage:/players/{id}/participations', error: e.message });
  }
}
