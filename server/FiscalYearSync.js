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

function fiscalSyncPayment_(value, heldOn) {
  if (String(value || '').trim() === '当日支払い') {
    return {
      payment_deadline: heldOn,
      payment_timing: 'on_tournament_day',
    };
  }
  return {
    payment_deadline: fiscalSyncDate_(value) || null,
    payment_timing: null,
  };
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

function getFiscalYearSyncProgress(operationId) {
  try {
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
  const warnings = [];
  const pendingTournaments = taikaiPendingTournaments_();
  Object.keys(pendingTournaments).forEach(name => {
    warnings.push(
      name + ': API障害時に作成されたためDB未同期です。'
      + '完全同期の成功後にこの状態は解除されます。'
    );
  });
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

    let structure;
    try {
      structure = tournamentSheetStructure_(sheet, true);
    } catch (e) {
      errors.push(sheetName + ': ' + e.message);
      return;
    }
    const data = structure.data;
    const paymentStatusIndex = structure.layout.payment_status_column - 1;

    const gradeDates = {};
    let isSanctioned = true;
    Object.keys(structure.grade_rows).forEach(grade => {
      const row = data[structure.grade_rows[grade] - 1] || [];
      if (row[paymentStatusIndex] instanceof Date) {
        gradeDates[grade] = row[paymentStatusIndex];
      }
    });
    if (structure.register_database_row) {
      const statusRow = data[structure.register_database_row] || [];
      isSanctioned = String(statusRow[1] || '').trim() === '';
    }

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
    const paymentDeadline = fiscalSyncDate_(calendarRow[10]);
    const isTournamentDayPayment = String(calendarRow[10] || '').trim() === '当日支払い';
    const paymentCompleted = String(calendarRow[11] || '').trim() === '済';
    if (paymentCompleted && !paymentDeadline && !isTournamentDayPayment) {
      warnings.push(
        sheetName + ': 振込完了が「済」ですが、本振込期限が未設定または日付として認識できません。'
      );
    }
    const feeResult = taikaiGradeFeesFromSheetData_(data, sheetName, currentFiscalGrades);
    errors.push.apply(errors, feeResult.errors);

    const schedules = [];
    currentFiscalGrades.forEach(grade => {
      const heldOn = fiscalSyncDate_(gradeDates[grade]);
      const payment = fiscalSyncPayment_(calendarRow[10], heldOn);
      schedules.push({
        held_on: heldOn,
        grade: grade,
        application_deadline: applicationDeadline,
        payment_deadline: payment.payment_deadline,
        payment_timing: payment.payment_timing,
        participation_fee_yen: Object.prototype.hasOwnProperty.call(feeResult.fees, grade)
          ? feeResult.fees[grade]
          : null,
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
    for (let i = 1; i < structure.response_end_index; i++) {
      const row = data[i];
      const playerName = String(row[2] || '').trim();
      if (!playerName || !/[ 　]/.test(playerName)) continue;

      const payStatus = String(row[paymentStatusIndex] || '').trim();
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
          // プレビュー内にも実メールを出さない。対応表への保存は実行時だけ行う。
          email: pseudonymousEmailFor_(email),
          real_email: normalizePrivateEmail_(email),
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
        // API送信用の公開スナップショットには実メールを残さない。
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
    warnings: warnings,
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

function previewFiscalYearDatabaseSync(operationId) {
  try {
    const snapshot = buildFiscalYearDatabaseSnapshot_(operationId);
    snapshot.mail_sync = buildMailManagementSnapshotSyncPlan_(snapshot);
    if (snapshot.mail_sync.errors.length) {
      snapshot.errors.push(
        'メール管理シートに' + snapshot.mail_sync.errors.length
        + '件の同期エラーがあります。'
      );
    }
    snapshot.tournaments.forEach(tournament => {
      tournament.entries.forEach(entry => { delete entry.real_email; });
    });
    return JSON.stringify(snapshot);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function syncFiscalYearDatabase() {
  try {
    const snapshot = buildFiscalYearDatabaseSnapshot_();
    const mailSnapshotPlan = buildMailManagementSnapshotSyncPlan_(snapshot);
    if (mailSnapshotPlan.errors.length) {
      throw new Error(
        'メール管理シートの同期前検証でエラーがあります。プレビューを確認してください。'
      );
    }
    if (snapshot.errors.length) {
      throw new Error('同期前検証でエラーがあります。プレビューを確認してください。');
    }
    rememberSnapshotEmailMappings_(snapshot);
    const result = taikaiApiRequest_('POST', '/admin/fiscal-year-sync', {
      fiscal_year: snapshot.fiscal_year,
      tournaments: snapshot.tournaments.map(item => ({
        name: item.name,
        registration_completed: item.registration_completed,
        payment_completed: item.payment_completed,
        schedules: item.schedules,
        entries: item.entries.map(entry => {
          const clean = Object.assign({}, entry);
          delete clean.real_email;
          return clean;
        }),
      })),
    }, null, {
      operation: '今年度のシート→DB完全同期',
      outcome: '同期処理を中断',
    });
    let mailResult;
    let gmailPlan;
    let gmailLinked = [];
    try {
      const mailPlan = resolveMailManagementSnapshotPlan_(mailSnapshotPlan);
      mailResult = executeMailManagementDatabaseSyncPlan_(mailPlan);
      gmailPlan = buildAnnouncementGmailLinkPlan_();
      gmailPlan.candidates.forEach(item => {
        taikaiApiRequest_(
          'PATCH',
          '/announcements/' + encodeURIComponent(item.announcement_id),
          { gmail_message_id: item.gmail_message_id }
        );
        gmailLinked.push({
          announcement_id: item.announcement_id,
          subject: item.subject,
        });
      });
    } catch (mailError) {
      return JSON.stringify({
        error: '大会・申込の差分同期は成功しましたが、メール同期に失敗しました。'
          + '同じ完全同期を再実行してください: ' + mailError.message,
        partial: true,
        result: result,
      });
    }
    const pendingCleared = taikaiClearPendingTournaments_(
      snapshot.tournaments.map(item => item.name)
    );
    const warnings = [];
    if (!pendingCleared) {
      warnings.push('DB未同期状態の解除に失敗しました。');
    }
    if (gmailPlan.errors.length) {
      warnings.push(
        'Gmail IDを一意に確認できない案内が'
        + gmailPlan.errors.length + '件残っています。'
      );
    }
    return JSON.stringify({
      ok: true,
      preview: snapshot.summary,
      result: result,
      mail_sync: mailResult,
      gmail_links: {
        linked_count: gmailLinked.length,
        already_linked_count: gmailPlan.skipped.length,
        unresolved_count: gmailPlan.errors.length,
        unresolved: gmailPlan.errors,
      },
      warning: warnings.join(' '),
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
