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

    const entryKeySet = {};
    entries.forEach(entry => {
      if (entry.canceled_at) return;
      const tournamentId = scheduleTournament[String(entry.schedule_id)];
      if (!tournamentId) return;
      entryKeySet[String(entry.player_id) + ':' + tournamentId] = true;
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
