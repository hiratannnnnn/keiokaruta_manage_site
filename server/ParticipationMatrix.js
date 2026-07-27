// ============================================================
// 選手 × 大会 出場一覧
// ============================================================

function getParticipationMatrix() {
  try {
    const players = taikaiApiRequest_('GET', '/players') || [];
    const tournaments = taikaiApiRequest_('GET', '/tournaments') || [];
    const schedules = taikaiApiRequest_('GET', '/schedules') || [];
    const entries = taikaiApiRequest_('GET', '/entries') || [];

    const scheduleInfo = {};
    const tournamentSchedules = {};
    schedules.forEach(schedule => {
      const tournamentId = String(schedule.tournament_id);
      const info = {
        id: String(schedule.id),
        tournament_id: tournamentId,
        held_on: String(schedule.held_on || ''),
        grade: String(schedule.grade || '').toUpperCase(),
        is_sanctioned: schedule.is_sanctioned === true,
      };
      scheduleInfo[info.id] = info;
      if (!tournamentSchedules[tournamentId]) tournamentSchedules[tournamentId] = [];
      tournamentSchedules[tournamentId].push(info);
    });

    // 同一選手・同一日程の再登録は、IDが最大のレコードだけを現在状態とする。
    // これにより、再登録後にキャンセルされた古い申込を復活表示しない。
    const latestEntries = {};
    entries.forEach(entry => {
      const scheduleId = String(entry.schedule_id);
      if (!scheduleInfo[scheduleId]) return;
      const key = String(entry.player_id) + ':' + scheduleId;
      const current = latestEntries[key];
      if (!current || taikaiCompareIds_(entry.id, current.id) > 0) {
        latestEntries[key] = entry;
      }
    });

    // 空欄セルは返さず、有効な選手×大会だけを大会単位に集約する。
    const participationMap = {};
    Object.keys(latestEntries).forEach(entryKey => {
      const entry = latestEntries[entryKey];
      if (entry.canceled_at) return;
      const schedule = scheduleInfo[String(entry.schedule_id)];
      const playerId = String(entry.player_id);
      const key = playerId + ':' + schedule.tournament_id;
      if (!participationMap[key]) {
        participationMap[key] = {
          player_id: playerId,
          tournament_id: schedule.tournament_id,
          schedule_ids: [],
          has_sanctioned: false,
          has_unsanctioned: false,
        };
      }
      const item = participationMap[key];
      if (!item.schedule_ids.includes(schedule.id)) item.schedule_ids.push(schedule.id);
      if (schedule.is_sanctioned) item.has_sanctioned = true;
      else item.has_unsanctioned = true;
    });
    const participations = Object.keys(participationMap).map(key => {
      const item = participationMap[key];
      return {
        player_id: item.player_id,
        tournament_id: item.tournament_id,
        schedule_ids: item.schedule_ids.sort(taikaiCompareIds_),
        mark: item.has_sanctioned && item.has_unsanctioned
          ? 'mixed'
          : (item.has_sanctioned ? 'sanctioned' : 'unsanctioned'),
      };
    });

    const visibleTournaments = tournaments.map(tournament => {
      const id = String(tournament.id);
      const related = tournamentSchedules[id] || [];
      const gradeSet = {};
      const heldOnSet = {};
      related.forEach(schedule => {
        if (/^[A-E]$/.test(schedule.grade)) gradeSet[schedule.grade] = true;
        if (schedule.held_on) heldOnSet[schedule.held_on] = true;
      });
      const name = String(tournament.name || '');
      const shortName = name
        .replace(/^第\s*[0-9０-９一二三四五六七八九十百千〇零]+\s*回\s*/, '')
        .replace(/大会$/, '')
        .trim() || name;
      return {
        id: id,
        name: name,
        shortName: shortName,
        grades: ['A', 'B', 'C', 'D', 'E'].filter(grade => gradeSet[grade]),
        heldOn: Object.keys(heldOnSet).sort(),
      };
    }).sort((a, b) =>
      String(a.heldOn[0] || '9999-99-99').localeCompare(
        String(b.heldOn[0] || '9999-99-99')
      ) || a.name.localeCompare(b.name, 'ja')
    );

    const counts = {};
    participations.forEach(item => {
      if (!counts[item.player_id]) {
        counts[item.player_id] = {
          sanctioned: 0,
          unsanctioned: 0,
          mixed: 0,
          allTournaments: 0,
        };
      }
      counts[item.player_id].allTournaments++;
      counts[item.player_id][item.mark]++;
    });
    const visiblePlayers = players.map(player => {
      const id = String(player.id);
      const count = counts[id] || {
        sanctioned: 0, unsanctioned: 0, mixed: 0, allTournaments: 0,
      };
      return {
        id: id,
        name: [player.family_name, player.given_name].filter(Boolean).join(' '),
        ruby: String(player.ruby || ''),
        // 混在大会も公認日程への出場があるため、公認回数へ1回加算する。
        totalCount: count.sanctioned + count.mixed,
        sanctionedCount: count.sanctioned + count.mixed,
        unsanctionedCount: count.unsanctioned + count.mixed,
        mixedCount: count.mixed,
        allTournamentCount: count.allTournaments,
      };
    }).sort((a, b) => a.name.localeCompare(b.name, 'ja'));

    return JSON.stringify({
      players: visiblePlayers,
      tournaments: visibleTournaments,
      participations: participations,
      truncated: players.length >= 100 || tournaments.length >= 100 ||
        schedules.length >= 100 || entries.length >= 100,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
