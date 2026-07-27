// ============================================================
// 選手 × 大会 出場一覧
// ============================================================

function participationMatrixFiscalYear_() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, 'JST', 'yyyy'));
  const month = Number(Utilities.formatDate(now, 'JST', 'M'));
  return month < 4 ? year - 1 : year;
}

function participationMatrixShortName_(name) {
  const fullName = String(name || '');
  return fullName
    .replace(/^第\s*[0-9０-９一二三四五六七八九十百千〇零]+\s*回\s*/, '')
    .replace(/大会$/, '')
    .trim() || fullName;
}

function getParticipationMatrix() {
  try {
    const fiscalYear = participationMatrixFiscalYear_();
    const source = taikaiApiRequest_(
      'GET',
      '/admin/participation-matrix',
      null,
      { fiscal_year: fiscalYear }
    ) || {};
    const participations = (source.participations || []).map(item => ({
      player_id: String(item.player_id),
      tournament_id: String(item.tournament_id),
      mark: String(item.mark || ''),
      schedule_ids: (item.schedule_ids || []).map(String),
      sanctioned_schedule_ids: (item.sanctioned_schedule_ids || []).map(String),
      unsanctioned_schedule_ids: (item.unsanctioned_schedule_ids || []).map(String),
    }));

    const countsByPlayer = {};
    participations.forEach(item => {
      if (!countsByPlayer[item.player_id]) {
        countsByPlayer[item.player_id] = { mixed: 0, all: 0 };
      }
      countsByPlayer[item.player_id].all++;
      if (item.mark === 'mixed') countsByPlayer[item.player_id].mixed++;
    });

    const players = (source.players || []).map(player => {
      const id = String(player.id);
      const counts = countsByPlayer[id] || { mixed: 0, all: 0 };
      const sanctionedCount = Number(player.sanctioned_count || 0);
      const unsanctionedCount = Number(player.unsanctioned_count || 0);
      return {
        id: id,
        name: [player.family_name, player.given_name].filter(Boolean).join(' '),
        ruby: String(player.ruby || ''),
        totalCount: sanctionedCount,
        sanctionedCount: sanctionedCount,
        unsanctionedCount: unsanctionedCount,
        mixedCount: counts.mixed,
        allTournamentCount: counts.all,
      };
    }).sort((left, right) => left.name.localeCompare(right.name, 'ja'));

    const tournaments = (source.tournaments || []).map(tournament => {
      const name = String(tournament.name || '');
      return {
        id: String(tournament.id),
        name: name,
        shortName: participationMatrixShortName_(name),
        grades: (tournament.grades || []).map(grade => String(grade).toUpperCase()),
        heldOn: (tournament.held_on || []).map(String).sort(),
      };
    }).sort((left, right) =>
      String(left.heldOn[0] || '9999-99-99').localeCompare(
        String(right.heldOn[0] || '9999-99-99')
      ) || left.name.localeCompare(right.name, 'ja')
    );

    return JSON.stringify({
      fiscalYear: Number(source.fiscal_year || fiscalYear),
      players: players,
      tournaments: tournaments,
      participations: participations,
      total: source.total || {
        players: players.length,
        tournaments: tournaments.length,
        participations: participations.length,
      },
      truncated: source.truncated === true,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
