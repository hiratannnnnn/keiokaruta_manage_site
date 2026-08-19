// ============================================================
// 選手 × 大会 出場一覧
// ============================================================

function participationMatrixFiscalYear_() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, 'JST', 'yyyy'));
  const month = Number(Utilities.formatDate(now, 'JST', 'M'));
  return month < 4 ? year - 1 : year;
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

function participationMatrixAllDatabaseRows_(resource) {
  const rows = [];
  let offset = 0;
  let total = 0;
  do {
    const page = taikaiApiRequest_(
      'GET', '/admin/database/' + resource, null, { limit: 100, offset: offset }
    ) || {};
    const pageRows = Array.isArray(page.rows) ? page.rows : [];
    total = Number(page.total || 0);
    rows.push.apply(rows, pageRows);
    offset += pageRows.length;
    if (!pageRows.length) break;
  } while (offset < total);
  return rows;
}

function participationMatrixInFiscalYear_(heldOn, fiscalYear) {
  const date = String(heldOn || '').slice(0, 10);
  return date >= fiscalYear + '-04-01'
    && date <= (fiscalYear + 1) + '-03-31';
}

function participationMatrixAllTournaments_(sourceTournaments, fiscalYear) {
  const tournaments = taikaiApiRequest_('GET', '/tournaments', null, {}) || [];
  const schedules = taikaiApiRequest_('GET', '/schedules', null, {}) || [];
  const schedulesByTournament = {};
  (Array.isArray(schedules) ? schedules : []).forEach(schedule => {
    const id = String(schedule.tournament_id || '');
    if (!schedulesByTournament[id]) schedulesByTournament[id] = [];
    schedulesByTournament[id].push(schedule);
  });

  const byId = {};
  (sourceTournaments || []).forEach(tournament => {
    byId[String(tournament.id)] = Object.assign({}, tournament);
  });
  (Array.isArray(tournaments) ? tournaments : []).forEach(tournament => {
    const id = String(tournament.id || '');
    const tournamentSchedules = schedulesByTournament[id] || [];
    const belongsToFiscalYear = Number(tournament.fiscal_year) === fiscalYear
      || tournamentSchedules.some(schedule =>
        participationMatrixInFiscalYear_(schedule.held_on, fiscalYear)
      );
    if (!id || !belongsToFiscalYear) return;
    const current = byId[id] || {};
    const grades = (current.grades || [])
      .concat(tournamentSchedules.map(schedule => schedule.grade))
      .filter(Boolean)
      .filter((grade, index, values) => values.indexOf(grade) === index);
    const heldOn = (current.held_on || [])
      .concat(tournamentSchedules.map(schedule => schedule.held_on))
      .filter(Boolean)
      .filter((date, index, values) => values.indexOf(date) === index);
    byId[id] = Object.assign({}, tournament, current, {
      grades: grades,
      held_on: heldOn,
    });
  });
  return Object.keys(byId).map(id => byId[id]);
}

function participationMatrixCancellationStatuses_() {
  const schedules = taikaiApiRequest_('GET', '/schedules', null, {}) || [];
  const scheduleById = {};
  (Array.isArray(schedules) ? schedules : []).forEach(schedule => {
    scheduleById[String(schedule.id)] = schedule;
  });
  const latestEntries = {};
  participationMatrixAllDatabaseRows_('entries').forEach(entry => {
    const scheduleId = String(entry.schedule_id || '');
    const key = String(entry.player_id || '') + ':' + scheduleId;
    const current = latestEntries[key];
    const entryId = entry.entry_id || entry.id;
    const currentId = current && (current.entry_id || current.id);
    if (!current || taikaiCompareIds_(entryId, currentId) > 0) {
      latestEntries[key] = entry;
    }
  });

  const states = {};
  Object.keys(latestEntries).forEach(key => {
    const entry = latestEntries[key];
    const schedule = scheduleById[String(entry.schedule_id || '')] || {};
    const status = cancellationStatusFromRecord_(Object.assign({}, entry, {
      lottery_result_date: schedule.lottery_result_date,
    }));
    if (status !== CANCELLATION_STATUS.NONE
        && status !== CANCELLATION_STATUS.AFTER) return;
    const participationKey = String(entry.player_id || '') + ':'
      + String(schedule.tournament_id || '');
    if (!states[participationKey]) states[participationKey] = [];
    states[participationKey].push(status);
  });

  const result = {};
  Object.keys(states).forEach(key => {
    result[key] = states[key].includes(CANCELLATION_STATUS.NONE)
      ? CANCELLATION_STATUS.NONE : CANCELLATION_STATUS.AFTER;
  });
  return result;
}

function getParticipationMatrix(fiscalYearInput) {
  try {
    const requestedFiscalYear = Number(fiscalYearInput);
    const fiscalYear = Number.isInteger(requestedFiscalYear)
      && requestedFiscalYear >= 2000 && requestedFiscalYear <= 2200
      ? requestedFiscalYear : participationMatrixFiscalYear_();
    const source = taikaiApiRequest_(
      'GET',
      '/admin/participation-matrix',
      null,
      { fiscal_year: fiscalYear }
    ) || {};
    const cancellationStatuses = participationMatrixCancellationStatuses_();
    const participations = (source.participations || []).map(item => {
      const playerId = String(item.player_id);
      const tournamentId = String(item.tournament_id);
      return {
        player_id: playerId,
        tournament_id: tournamentId,
        mark: String(item.mark || ''),
        cancellation_status: cancellationNormalizeExplicitStatus_(
          item.cancellation_status || item.cancellation_timing
        ) || cancellationStatuses[playerId + ':' + tournamentId]
          || CANCELLATION_STATUS.NONE,
        schedule_ids: (item.schedule_ids || []).map(String),
        sanctioned_schedule_ids: (item.sanctioned_schedule_ids || []).map(String),
        unsanctioned_schedule_ids: (item.unsanctioned_schedule_ids || []).map(String),
      };
    });

    const countsByPlayer = {};
    participations.forEach(item => {
      if (!countsByPlayer[item.player_id]) {
        countsByPlayer[item.player_id] = { mixed: 0, all: 0 };
      }
      countsByPlayer[item.player_id].all++;
      if (item.mark === 'mixed') countsByPlayer[item.player_id].mixed++;
    });

    const databasePlayers = participationMatrixAllDatabaseRows_('players');
    const playerSortOrders = {};
    databasePlayers.forEach(player => {
      playerSortOrders[String(player.id)] = player.sort_order;
    });
    const players = (source.players || []).map((player, sourceIndex) => {
      const id = String(player.id);
      const counts = countsByPlayer[id] || { mixed: 0, all: 0 };
      const sanctionedCount = Number(player.sanctioned_count || 0);
      const unsanctionedCount = Number(player.unsanctioned_count || 0);
      const rawSortOrder = player.sort_order !== undefined
        ? player.sort_order : playerSortOrders[id];
      const sortOrder = rawSortOrder === null || rawSortOrder === ''
        ? NaN : Number(rawSortOrder);
      return {
        id: id,
        name: [player.family_name, player.given_name].filter(Boolean).join(' '),
        ruby: String(player.ruby || ''),
        totalCount: sanctionedCount,
        sanctionedCount: sanctionedCount,
        unsanctionedCount: unsanctionedCount,
        mixedCount: counts.mixed,
        allTournamentCount: counts.all,
        sort_order: rawSortOrder === undefined ? null : rawSortOrder,
        _sortOrder: Number.isFinite(sortOrder) ? sortOrder : null,
        _sourceIndex: sourceIndex,
      };
    }).sort((left, right) =>
      left._sortOrder !== null && right._sortOrder !== null
        ? left._sortOrder - right._sortOrder
          || left._sourceIndex - right._sourceIndex
        : left._sortOrder !== null ? -1
          : right._sortOrder !== null ? 1
            : left._sourceIndex - right._sourceIndex
    ).map(player => {
      delete player._sortOrder;
      delete player._sourceIndex;
      return player;
    });

    const tournaments = participationMatrixAllTournaments_(
      source.tournaments || [], fiscalYear
    ).map(tournament => {
      const name = String(tournament.name || '');
      return {
        id: String(tournament.id),
        name: name,
        grades: (tournament.grades || [])
          .map(grade => String(grade).toUpperCase())
          .filter((grade, index, values) => values.indexOf(grade) === index)
          .sort(),
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
