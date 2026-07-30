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

function fiscalSyncCurrentGradesForSheet_(sheetName, structure, fiscalYear) {
  const gradeDates = {};
  Object.keys(structure.grade_rows).forEach(grade => {
    const value = tournamentSheetGradeDate_(structure, grade);
    if (value instanceof Date || fiscalSyncDate_(value)) {
      gradeDates[grade] = value;
    }
  });
  const match = String(sheetName || '').match(/([A-E]+)級$/);
  const declaredGrades = match ? match[1].split('') : Object.keys(gradeDates);
  return declaredGrades.filter(grade =>
    fiscalSyncInYear_(gradeDates[grade], fiscalYear)
  );
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

function fiscalSyncCalendarColumn_(headerRows, label) {
  const columns = [];
  (headerRows || []).forEach(row => {
    (row || []).forEach((value, index) => {
      if (String(value || '').replace(/\s+/g, '') === label.replace(/\s+/g, '')) {
        if (!columns.includes(index)) columns.push(index);
      }
    });
  });
  if (columns.length !== 1) {
    throw new Error(
      'カレンダーの「' + label + '」列を一意に特定できません'
      + '（候補' + columns.length + '件）。'
    );
  }
  return columns[0];
}

function fiscalSyncOptionalDate_(value) {
  const raw = String(value || '').trim();
  if (!raw) return { value: null, valid: true };
  const formatted = fiscalSyncDate_(value);
  return { value: formatted || null, valid: Boolean(formatted) };
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

  const calendarData = calendar.getRange(1, 1, calendar.getLastRow(), 16).getValues();
  const calendarRows = calendarData.slice(2);
  const grouped = {};
  const errors = [];
  const warnings = [];
  try {
    tournamentSheetValidateGradeOwnership_(
      calendarRows.map(row => String(row[0] || '').trim()).filter(Boolean)
    );
  } catch (ownershipError) {
    errors.push(ownershipError.message);
  }
  let internalPaymentDeadlineIndex = null;
  try {
    internalPaymentDeadlineIndex = fiscalSyncCalendarColumn_(
      calendarData.slice(0, 2), '振込開始'
    );
  } catch (e) {
    errors.push(e.message + ' 会内振込期限を安全に同期できません。');
  }
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
    const isSanctioned = tournamentSheetIsSanctioned_(structure);
    const currentFiscalGrades = fiscalSyncCurrentGradesForSheet_(
      sheetName, structure, fiscalYear
    );
    // 年度判定に使うのは開催日だけ。申込・抽選・振込期限は年度をまたいでよい。
    if (!currentFiscalGrades.length) return;
    const gradeDates = {};
    currentFiscalGrades.forEach(grade => {
      gradeDates[grade] = tournamentSheetGradeDate_(structure, grade);
    });

    const baseName = sheetName.replace(/[A-E]+級$/, '');
    const applicationDeadline = fiscalSyncDate_(calendarRow[5]);
    if (structure.version === 1 && !applicationDeadline) {
      errors.push(sheetName + ': 申込期限が未設定です。');
    }
    const paymentSetting = structure.version === 2
      ? tournamentSheetManagementValue_(structure, '本振込期限')
      : calendarRow[10];
    const paymentDeadline = fiscalSyncDate_(paymentSetting);
    const paymentTiming = structure.version === 2
      ? tournamentSheetPaymentTimingFromStructure_(structure)
      : fiscalSyncPayment_(calendarRow[10], '').payment_timing;
    const paymentCompleted = String(calendarRow[11] || '').trim() === '済';
    if (paymentCompleted && !paymentDeadline
        && paymentTiming !== 'on_tournament_day'
        && paymentTiming !== 'after_tournament') {
      warnings.push(
        sheetName + ': 振込完了が「済」ですが、本振込期限が未設定または日付として認識できません。'
      );
    }
    const feeResult = taikaiGradeFeesFromStructure_(
      structure, sheetName, currentFiscalGrades
    );
    errors.push.apply(errors, feeResult.errors);
    let paymentInstructions;
    let hasPaymentInstructions = true;
    if (structure.version === 1) {
      try {
        paymentInstructions = tournamentSheetPaymentInstructionsFromStructure_(structure);
      } catch (e) {
        hasPaymentInstructions = false;
        errors.push(
          sheetName + ': ' + e.message
          + ' 振込先を安全に同期できません。'
        );
      }
    }
    let internalPaymentDeadline = { value: null, valid: true };
    if (structure.version === 1 && internalPaymentDeadlineIndex !== null) {
      internalPaymentDeadline = fiscalSyncOptionalDate_(
        calendarRow[internalPaymentDeadlineIndex]
      );
      if (!internalPaymentDeadline.valid) {
        errors.push(
          sheetName + ': 会内振込期限「'
          + String(calendarRow[internalPaymentDeadlineIndex])
          + '」を日付として認識できません。'
        );
      }
    }

    const schedules = [];
    currentFiscalGrades.forEach(grade => {
      const heldOn = fiscalSyncDate_(gradeDates[grade]);
      const payment = fiscalSyncPayment_(calendarRow[10], heldOn);
      const v2ApplicationDeadline = structure.version === 2
        ? fiscalSyncDate_(
          tournamentSheetManagementValue_(structure, '本申込期限')
        )
        : applicationDeadline;
      if (!v2ApplicationDeadline) {
        errors.push(sheetName + ': ' + grade + '級の申込期限が未設定です。');
      }
      const schedule = {
        held_on: heldOn,
        grade: grade,
        application_deadline: v2ApplicationDeadline,
        payment_deadline: structure.version === 2
          ? (fiscalSyncDate_(
            tournamentSheetManagementValue_(structure, '本振込期限')
          ) || null)
          : payment.payment_deadline,
        payment_timing: structure.version === 2
          ? tournamentSheetPaymentTimingFromStructure_(structure)
          : payment.payment_timing,
        participation_fee_yen: Object.prototype.hasOwnProperty.call(feeResult.fees, grade)
          ? feeResult.fees[grade]
          : null,
        lottery_result_date: structure.version === 2
          ? (fiscalSyncDate_(
            tournamentSheetManagementValue_(structure, '抽選日')
          ) || null)
          : (fiscalSyncDate_(calendarRow[7]) || null),
        is_sanctioned: isSanctioned,
      };
      if (structure.version === 2) {
        schedule.payment_instructions =
          tournamentSheetPaymentInstructionsFromStructure_(structure);
      } else {
        schedule.venue = null;
        schedule.reception_ends_at = null;
        if (internalPaymentDeadlineIndex !== null && internalPaymentDeadline.valid) {
          schedule.internal_payment_deadline = internalPaymentDeadline.value;
        }
        if (hasPaymentInstructions) {
          schedule.payment_instructions = paymentInstructions;
        }
      }
      schedules.push(schedule);
    });

    const header = data[0] || [];
    let responseColumns;
    try {
      responseColumns = tournamentSheetResponseColumns_(structure);
    } catch (e) {
      errors.push(sheetName + ': ' + e.message);
      return;
    }
    const rubyIndex = header.findIndex(value => String(value || '').includes('ふりがな'));
    const clubIndex = header.findIndex(value => String(value || '').includes('所属'));
    const entries = [];
    const responseRecords = tournamentSheetResponseRecords_(structure, false);
    responseRecords.forEach(record => {
      const row = record.raw_values;
      const playerName = record.name;
      if (!playerName || !/[ 　]/.test(playerName)) return;

      const selectionStatus = record.selection_status;
      const isCarriedOver = structure.version === 1
        && ((record.raw_sheet_status.includes('繰')
            && record.raw_sheet_status.includes('越'))
          || record.raw_sheet_status === 'くりこし');
      const isTarget = selectionStatus === ''
        || (structure.version === 1 && isCarriedOver);

      const email = record.email;
      const grade = record.grade;
      if (!currentFiscalGrades.includes(grade)) return;
      const heldOn = fiscalSyncDate_(gradeDates[grade]);
      if (!email) {
        errors.push(sheetName + ': ' + playerName + 'のメールアドレスがありません。');
        return;
      }
      if (!/^[A-E]$/.test(grade) || !heldOn) {
        errors.push(sheetName + ': ' + playerName + 'の級または開催日を特定できません。');
        return;
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
          is_paid: record.is_paid,
          payment_method: isCarriedOver ? 'carried_over' : 'bank_transfer',
          // v2の支払状態はAPIからの書戻し値。完全同期で履歴を再生成しない。
          sync_payment: structure.version !== 2,
          source_sheet: sheetName,
          source_row: record.source_row,
          // 同一人物・同一大会の重複時は、回答時刻が新しい行を正とする。
          // 時刻がない場合も後から読み込んだ行を採用できるよう順序を保持する。
          registered_at_ms: tournamentResponseTimestampMs_(
            row[responseColumns.timestamp]
          ),
          source_order: entrySourceOrder++,
          is_sync_target: isTarget,
        });
      } catch (e) {
        errors.push(sheetName + ': ' + e.message);
      }
    });

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
  const latestByEmail = tournamentLatestByEmail_(entries);
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
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return JSON.stringify({ error: '別の年度完全同期が実行中です。完了後に再実行してください。' });
  }
  try {
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
    const warnings = [];
    if (gmailPlan.errors.length) {
      warnings.push(
        'Gmail IDを一意に確認できない案内が'
        + gmailPlan.errors.length + '件残っています。'
      );
    }
    const pendingCleared = gmailPlan.errors.length
      ? false
      : taikaiClearPendingTournaments_(
        snapshot.tournaments.map(item => item.name)
      );
    if (!gmailPlan.errors.length && !pendingCleared) {
      return JSON.stringify({
        error: 'DB/API・Gmailの同期は完了しましたが、'
          + 'DB未同期状態の解除に失敗しました。同じ完全同期を再実行してください。',
        partial: true,
        result: result,
        mail_sync: mailResult,
        gmail_links: {
          linked_count: gmailLinked.length,
          already_linked_count: gmailPlan.skipped.length,
          unresolved_count: 0,
          unresolved: [],
        },
      });
    }
    if (gmailPlan.errors.length) {
      return JSON.stringify({
        error: 'DB/APIと大会シートの同期は完了しましたが、Gmail IDを一意に'
          + '確認できない案内が' + gmailPlan.errors.length
          + '件あります。同じ完全同期を再実行してください。',
        partial: true,
        result: result,
        mail_sync: mailResult,
        gmail_links: {
          linked_count: gmailLinked.length,
          already_linked_count: gmailPlan.skipped.length,
          unresolved_count: gmailPlan.errors.length,
          unresolved: gmailPlan.errors,
        },
      });
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
  } finally {
    lock.releaseLock();
  }
}
