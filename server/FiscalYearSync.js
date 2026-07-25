// ============================================================
// カレンダー・大会シートから今年度DBを完全同期
// ============================================================

function fiscalYearForDate_(date) {
  return date.getMonth() < 3 ? date.getFullYear() - 1 : date.getFullYear();
}

function fiscalSyncDate_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return '';
  return Utilities.formatDate(value, 'JST', 'yyyy-MM-dd');
}

function fiscalSyncInYear_(value, fiscalYear) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return false;
  const text = fiscalSyncDate_(value);
  return text >= fiscalYear + '-04-01' && text <= (fiscalYear + 1) + '-03-31';
}

function fiscalSyncPlayer_(name) {
  const parts = String(name || '').trim().split(/[ 　]+/).filter(Boolean);
  if (parts.length < 2) throw new Error('氏名を名字と名前に分けられません: ' + name);
  return { family_name: parts.shift(), given_name: parts.join(' ') };
}

function fiscalSyncProgressKey_(operationId) {
  return 'fiscal_sync_progress_' + String(operationId || '').replace(/[^a-zA-Z0-9_-]/g, '');
}

function fiscalSyncSetProgress_(operationId, progress) {
  if (!operationId) return;
  CacheService.getUserCache().put(
    fiscalSyncProgressKey_(operationId),
    JSON.stringify(progress),
    600
  );
}

function getFiscalYearSyncProgress(operationId, password) {
  try {
    databaseAdminAuthenticate_(password);
    const raw = CacheService.getUserCache().get(fiscalSyncProgressKey_(operationId));
    return raw || JSON.stringify({ phase: 'starting', processed: 0, total: 0 });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function buildFiscalYearDatabaseSnapshot_(operationId) {
  const now = new Date();
  const fiscalYear = fiscalYearForDate_(now);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const calendar = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
  if (!calendar) throw new Error('カレンダーシートが見つかりません。');

  const calendarRows = calendar.getRange(1, 1, calendar.getLastRow(), 16).getValues().slice(2);
  const grouped = {};
  const errors = [];
  let entrySourceOrder = 0;
  fiscalSyncSetProgress_(operationId, {
    phase: 'scanning',
    processed: 0,
    total: calendarRows.length,
    current: 'カレンダーを読み込みました',
  });

  calendarRows.forEach((calendarRow, rowIndex) => {
    const sheetName = String(calendarRow[0] || '').trim();
    fiscalSyncSetProgress_(operationId, {
      phase: 'scanning',
      processed: rowIndex,
      total: calendarRows.length,
      current: sheetName || '空行',
    });
    if (!sheetName) return;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      if (fiscalSyncInYear_(calendarRow[2], fiscalYear)) {
        errors.push(sheetName + ': 大会シートが見つかりません。');
      }
      return;
    }

    const data = sheet.getDataRange().getValues();
    const count = (data[0] || []).find(value => typeof value === 'number' && Number.isFinite(value));
    if (count == null) {
      errors.push(sheetName + ': カラム数Nを取得できません。');
      return;
    }

    const gradeDates = {};
    let isSanctioned = true;
    data.forEach((row, index) => {
      const grade = String(row[0] || '').trim();
      if (/^[A-E]$/.test(grade) && row[count + 2] instanceof Date) {
        gradeDates[grade] = row[count + 2];
      }
      if (String(row[0] || '').trim() === 'registerDatabase' && data[index + 1]) {
        isSanctioned = String(data[index + 1][1] || '').trim() === '';
      }
    });

    const sheetGradesMatch = sheetName.match(/([A-E]+)級$/);
    const declaredGrades = sheetGradesMatch ? sheetGradesMatch[1].split('') : Object.keys(gradeDates);
    const currentFiscalGrades = declaredGrades.filter(grade =>
      fiscalSyncInYear_(gradeDates[grade], fiscalYear)
    );
    // 年度判定に使うのは開催日だけ。申込・抽選・振込期限は年度をまたいでよい。
    if (!currentFiscalGrades.length) return;

    const baseName = sheetName.replace(/[A-E]+級$/, '');
    const applicationDeadline = fiscalSyncDate_(calendarRow[5]);
    if (!applicationDeadline) errors.push(sheetName + ': 申込期限が未設定です。');

    const schedules = [];
    currentFiscalGrades.forEach(grade => {
      const heldOn = fiscalSyncDate_(gradeDates[grade]);
      schedules.push({
        held_on: heldOn,
        grade: grade,
        application_deadline: applicationDeadline,
        payment_deadline: fiscalSyncDate_(calendarRow[10]) || null,
        payment_timing: null,
        lottery_result_date: fiscalSyncDate_(calendarRow[7]) || null,
        venue: null,
        reception_ends_at: null,
        is_sanctioned: isSanctioned,
      });
    });

    const header = data[0] || [];
    const rubyIndex = header.findIndex(value => String(value || '').includes('ふりがな'));
    const clubIndex = header.findIndex(value => String(value || '').includes('所属'));
    const entries = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (typeof row[2] === 'number' || String(row[0] || '').trim() === 'registerDatabase') break;
      const playerName = String(row[2] || '').trim();
      if (!playerName || !/[ 　]/.test(playerName)) continue;

      const payStatus = String(row[count + 2] || '').trim();
      const isTarget = payStatus === '' || payStatus === '済'
        || (payStatus.includes('繰') && payStatus.includes('越'))
        || payStatus === 'くりこし';

      const email = String(row[1] || '').trim();
      const grade = String(row[4] || '').replace(/級/g, '').trim().toUpperCase();
      if (!currentFiscalGrades.includes(grade)) continue;
      const heldOn = fiscalSyncDate_(gradeDates[grade]);
      if (!email) {
        errors.push(sheetName + ': ' + playerName + 'のメールアドレスがありません。');
        continue;
      }
      if (!/^[A-E]$/.test(grade) || !heldOn) {
        errors.push(sheetName + ': ' + playerName + 'の級または開催日を特定できません。');
        continue;
      }
      try {
        const player = fiscalSyncPlayer_(playerName);
        entries.push({
          family_name: player.family_name,
          given_name: player.given_name,
          ruby: rubyIndex >= 0 ? String(row[rubyIndex] || '').trim() || null : null,
          email: email,
          club: clubIndex >= 0 ? String(row[clubIndex] || '').trim() || null : null,
          grade: grade,
          held_on: heldOn,
          is_paid: payStatus === '済'
            || (payStatus.includes('繰') && payStatus.includes('越'))
            || payStatus === 'くりこし',
          source_sheet: sheetName,
          source_row: i + 1,
          // 同一人物・同一大会の重複時は、回答時刻が新しい行を正とする。
          // 時刻がない場合も後から読み込んだ行を採用できるよう順序を保持する。
          registered_at_ms: row[0] && typeof row[0].getTime === 'function' && !isNaN(row[0].getTime())
            ? row[0].getTime()
            : null,
          source_order: entrySourceOrder++,
          is_sync_target: isTarget,
        });
      } catch (e) {
        errors.push(sheetName + ': ' + e.message);
      }
    }

    if (!grouped[baseName]) {
      grouped[baseName] = {
        name: baseName,
        registration_completed: true,
        payment_completed: true,
        schedules: [],
        entries: [],
        sheet_names: [],
      };
    }
    const tournament = grouped[baseName];
    tournament.registration_completed =
      tournament.registration_completed && String(calendarRow[6] || '').trim() === '済';
    tournament.payment_completed =
      tournament.payment_completed && String(calendarRow[11] || '').trim() === '済';
    tournament.schedules = tournament.schedules.concat(schedules);
    tournament.entries = tournament.entries.concat(entries);
    tournament.sheet_names.push(sheetName);
  });

  const tournaments = Object.keys(grouped).map(name => grouped[name]);
  tournaments.forEach(tournament => {
    tournament.entries = fiscalSyncLatestEntries_(tournament.entries)
      .filter(entry => entry.is_sync_target)
      .map(entry => {
        const clean = Object.assign({}, entry);
        delete clean.registered_at_ms;
        delete clean.source_order;
        delete clean.is_sync_target;
        return clean;
      });

    const scheduleKeys = {};
    tournament.schedules.forEach(schedule => {
      const key = schedule.grade + '|' + schedule.held_on;
      if (scheduleKeys[key]) {
        errors.push(tournament.name + ': 同じ級・開催日の日程が重複しています（' + key + '）。');
      }
      scheduleKeys[key] = true;
    });
  });
  fiscalSyncSetProgress_(operationId, {
    phase: 'complete',
    processed: calendarRows.length,
    total: calendarRows.length,
    current: '検査完了',
  });
  return {
    fiscal_year: fiscalYear,
    fiscal_start: fiscalYear + '-04-01',
    fiscal_end: (fiscalYear + 1) + '-03-31',
    tournaments: tournaments,
    errors: errors,
    summary: {
      tournament_count: tournaments.length,
      schedule_count: tournaments.reduce((sum, item) => sum + item.schedules.length, 0),
      entry_count: tournaments.reduce((sum, item) => sum + item.entries.length, 0),
    },
  };
}

function fiscalSyncLatestEntries_(entries) {
  const latestByEmail = {};
  (entries || []).forEach(entry => {
    const key = String(entry.email || '').trim().toLowerCase();
    const current = latestByEmail[key];
    const entryTime = entry.registered_at_ms === null ? -Infinity : Number(entry.registered_at_ms);
    const currentTime = !current || current.registered_at_ms === null
      ? -Infinity
      : Number(current.registered_at_ms);
    if (!current || entryTime > currentTime ||
        (entryTime === currentTime && Number(entry.source_order) > Number(current.source_order))) {
      latestByEmail[key] = entry;
    }
  });
  return Object.keys(latestByEmail).map(key => latestByEmail[key]);
}

function previewFiscalYearDatabaseSync(password, operationId) {
  try {
    databaseAdminAuthenticate_(password);
    return JSON.stringify(buildFiscalYearDatabaseSnapshot_(operationId));
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function syncFiscalYearDatabase(password) {
  try {
    databaseAdminAuthenticate_(password);
    const snapshot = buildFiscalYearDatabaseSnapshot_();
    if (snapshot.errors.length) {
      throw new Error('同期前検証でエラーがあります。プレビューを確認してください。');
    }
    const result = taikaiApiRequest_('POST', '/admin/fiscal-year-sync', {
      fiscal_year: snapshot.fiscal_year,
      tournaments: snapshot.tournaments.map(item => ({
        name: item.name,
        registration_completed: item.registration_completed,
        payment_completed: item.payment_completed,
        schedules: item.schedules,
        entries: item.entries,
      })),
    });
    return JSON.stringify({ ok: true, preview: snapshot.summary, result: result });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
