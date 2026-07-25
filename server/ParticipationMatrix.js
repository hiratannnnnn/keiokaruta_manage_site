// ============================================================
// 選手 × 大会 出場一覧
// ============================================================

function getParticipationMatrix() {
  try {
    const players = taikaiApiRequest_('GET', '/players') || [];
    const tournaments = taikaiApiRequest_('GET', '/tournaments') || [];
    const schedules = taikaiApiRequest_('GET', '/schedules') || [];
    const entries = taikaiApiRequest_('GET', '/entries') || [];

    const scheduleTournament = {};
    const tournamentHeldOn = {};
    schedules.forEach(schedule => {
      const tournamentId = String(schedule.tournament_id);
      scheduleTournament[String(schedule.id)] = tournamentId;
      const heldOn = String(schedule.held_on || '');
      if (heldOn && (!tournamentHeldOn[tournamentId] || heldOn < tournamentHeldOn[tournamentId])) {
        tournamentHeldOn[tournamentId] = heldOn;
      }
    });

    // 同一選手・同一大会に複数登録がある場合は、後から登録された
    // entry（IDが最大のもの）だけを現在の状態として採用する。
    const latestEntries = {};
    entries.forEach(entry => {
      const tournamentId = scheduleTournament[String(entry.schedule_id)];
      if (!tournamentId) return;
      const key = String(entry.player_id) + ':' + tournamentId;
      const current = latestEntries[key];
      if (!current || compareTaikaiIds_(entry.id, current.id) > 0) {
        latestEntries[key] = entry;
      }
    });
    const entryKeySet = {};
    Object.keys(latestEntries).forEach(key => {
      if (!latestEntries[key].canceled_at) entryKeySet[key] = true;
    });

    const visibleTournaments = tournaments.map(tournament => ({
      id: String(tournament.id),
      name: String(tournament.name || ''),
      heldOn: tournamentHeldOn[String(tournament.id)] || '',
    })).sort((a, b) =>
      String(a.heldOn || '9999-99-99').localeCompare(String(b.heldOn || '9999-99-99')) ||
      a.name.localeCompare(b.name, 'ja')
    );

    const visiblePlayers = players.map(player => ({
      id: String(player.id),
      name: [player.family_name, player.given_name].filter(Boolean).join(' '),
      ruby: String(player.ruby || ''),
    })).sort((a, b) =>
      a.name.localeCompare(b.name, 'ja')
    );

    return JSON.stringify({
      players: visiblePlayers,
      tournaments: visibleTournaments,
      entryKeys: Object.keys(entryKeySet),
      truncated: players.length >= 100 || tournaments.length >= 100 ||
        schedules.length >= 100 || entries.length >= 100,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function compareTaikaiIds_(left, right) {
  const leftText = String(left === undefined || left === null ? '' : left);
  const rightText = String(right === undefined || right === null ? '' : right);
  if (/^\d+$/.test(leftText) && /^\d+$/.test(rightText)) {
    if (leftText.length !== rightText.length) return leftText.length - rightText.length;
  }
  return leftText.localeCompare(rightText);
}
