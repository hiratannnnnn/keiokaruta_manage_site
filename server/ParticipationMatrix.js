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

function participationMatrixNotes_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PLAYER_NOTES);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const notes = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(row => {
    const playerId = String(row[0] || '').trim();
    const memo = String(row[1] || '');
    if (playerId && memo) notes[playerId] = memo;
  });
  return notes;
}

function participationMatrixNotesSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PLAYER_NOTES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.PLAYER_NOTES);
    sheet.getRange(1, 1, 1, 3).setValues([[
      'player_id', 'memo', 'updated_at',
    ]]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function saveParticipationPlayerNote(playerId, memo) {
  const normalizedId = String(playerId || '').trim();
  const normalizedMemo = String(memo || '').trim();
  if (!/^\d+$/.test(normalizedId)) {
    return JSON.stringify({ error: '選手IDが不正です。' });
  }
  if (normalizedMemo.length > 1000) {
    return JSON.stringify({ error: 'メモは1000文字以内で入力してください。' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = participationMatrixNotesSheet_(ss);
    const lastRow = sheet.getLastRow();
    const ids = lastRow >= 2
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => String(row[0]))
      : [];
    const index = ids.indexOf(normalizedId);
    if (!normalizedMemo) {
      if (index >= 0) sheet.deleteRow(index + 2);
    } else {
      const targetRow = index >= 0 ? index + 2 : sheet.getLastRow() + 1;
      sheet.getRange(targetRow, 1, 1, 3).setValues([[
        normalizedId,
        normalizedMemo,
        Utilities.formatDate(new Date(), 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX"),
      ]]);
    }
    return JSON.stringify({
      ok: true,
      player_id: normalizedId,
      memo: normalizedMemo,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
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

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const notes = participationMatrixNotes_(ss);
    const visiblePlayerIds = {};
    players.forEach(player => { visiblePlayerIds[player.id] = true; });
    const visibleNotes = {};
    Object.keys(notes).forEach(playerId => {
      if (visiblePlayerIds[playerId]) visibleNotes[playerId] = notes[playerId];
    });

    return JSON.stringify({
      fiscalYear: Number(source.fiscal_year || fiscalYear),
      players: players,
      tournaments: tournaments,
      participations: participations,
      playerNotes: visibleNotes,
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
