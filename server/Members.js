// ============================================================
// 名簿関連
// ============================================================

// 選手マスタを全件取得する。通常の名簿画面にはメールアドレスを返さない。
function getMembers() {
  try {
    const players = [];
    let offset = 0;
    let total = 0;
    do {
      const page = taikaiApiRequest_(
        'GET', '/admin/database/players', null, { limit: 100, offset: offset }
      ) || {};
      const rows = Array.isArray(page.rows) ? page.rows : [];
      total = Number(page.total || 0);
      players.push.apply(players, rows);
      offset += rows.length;
      if (!rows.length) break;
    } while (offset < total);

    if (players.length < total) {
      throw new Error(
        '選手一覧を最後まで取得できませんでした（'
        + players.length + '/' + total + '件）。'
      );
    }

    const members = players.map(player => ({
      id: String(player.id),
      name: [player.family_name, player.given_name].filter(Boolean).join(' '),
      ruby: String(player.ruby || ''),
      club: String(player.club || ''),
    }));
    return JSON.stringify(members);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
