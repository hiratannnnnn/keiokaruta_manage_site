// ============================================================
// メール管理シート操作
// ============================================================

// メール管理シートのデータを取得（行6以降）
function getMailManagement() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (!sheet) throw new Error('メール管理シートが見つかりません');

    const lastRow = sheet.getLastRow();
    if (lastRow < 6) return JSON.stringify([]);

    const rows = sheet.getRange(6, 1, lastRow - 5, 9).getValues()
      .map((row, i) => ({
        rowNum:         i + 6,
        tournamentName: String(row[0] || ''),
        grades:         String(row[1] || ''),
        sendDateTime:   row[2] instanceof Date
          ? Utilities.formatDate(row[2], 'JST', 'yyyy-MM-dd HH:mm:ss')
          : String(row[2] || ''),
        mailType:       String(row[3] || ''),
        threadTitle:    String(row[4] || ''),
        formLink:       String(row[5] || ''),
        reminderSet:    String(row[6] || ''),
        sent:           String(row[7] || ''),
        includeNotPaid: !!row[8],
      }))
      .filter(r => r.tournamentName);

    return JSON.stringify(rows);
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

function mailManagementSyncDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const text = String(value || '').trim();
  if (!text) return null;
  const parsed = new Date(text.replace(' ', 'T'));
  return isNaN(parsed.getTime()) ? null : parsed;
}

function mailManagementSyncDateTime_(value) {
  const date = mailManagementSyncDate_(value);
  return date ? Utilities.formatDate(date, 'JST', "yyyy-MM-dd'T'HH:mm:ssXXX") : '';
}

function mailManagementScheduleIds_(tournament, grades, schedules) {
  const gradeList = String(grades || '').replace(/級/g, '').replace(/\s+/g, '')
    .toUpperCase().split('').filter(Boolean);
  if (!gradeList.length || gradeList.some(grade => !/^[A-E]$/.test(grade))) {
    throw new Error('級をA〜Eで特定できません。');
  }
  const ids = [];
  gradeList.forEach(grade => {
    const matches = schedules.filter(schedule =>
      String(schedule.tournament_id) === String(tournament.id)
      && String(schedule.grade || '').toUpperCase() === grade
    );
    if (matches.length !== 1) {
      throw new Error(
        grade + '級の日程が' + matches.length + '件あり、一意に特定できません。'
      );
    }
    ids.push(String(matches[0].id));
  });
  return ids.sort((left, right) => taikaiCompareIds_(left, right));
}

function buildMailManagementDatabaseSyncPlan_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
  if (!sheet) throw new Error('メール管理シートが見つかりません。');

  const lastRow = sheet.getLastRow();
  const rows = lastRow < 6 ? [] : sheet.getRange(6, 1, lastRow - 5, 9).getValues();
  const tournaments = taikaiApiRequest_('GET', '/tournaments', null, {}) || [];
  const schedules = taikaiApiRequest_('GET', '/schedules', null, {}) || [];
  const now = new Date();
  const candidates = [];
  const skipped = [];
  const errors = [];

  rows.forEach((row, index) => {
    const rowNum = index + 6;
    const tournamentName = String(row[0] || '').trim();
    if (!tournamentName) return;
    const sent = String(row[7] || '').trim();
    if (sent.includes('済')) {
      skipped.push({ row: rowNum, reason: '送信済み' });
      return;
    }
    const scheduledDate = mailManagementSyncDate_(row[2]);
    const scheduledAt = mailManagementSyncDateTime_(scheduledDate);
    if (!scheduledAt) {
      skipped.push({ row: rowNum, reason: '送信予定日時が未設定または日付ではありません' });
      return;
    }
    if (scheduledDate.getTime() <= now.getTime()) {
      skipped.push({ row: rowNum, reason: '送信予定日時が過去です' });
      return;
    }

    const mailTypeText = String(row[3] || '').trim();
    const mailType = mailTypeText === 'リマインダー'
      ? 'reminder'
      : (mailTypeText === '振込確認' ? 'payment_confirmation' : '');
    if (!mailType) {
      errors.push({ row: rowNum, reason: 'メール種別を特定できません' });
      return;
    }

    const tournamentMatches = tournaments.filter(tournament =>
      String(tournament.name || '').trim() === tournamentName
    );
    if (tournamentMatches.length !== 1) {
      errors.push({
        row: rowNum,
        reason: '大会が' + tournamentMatches.length + '件あり、一意に特定できません',
      });
      return;
    }

    try {
      const scheduleIds = mailManagementScheduleIds_(
        tournamentMatches[0],
        String(row[1] || ''),
        schedules
      );
      candidates.push({
        row: rowNum,
        tournament_name: tournamentName,
        grades: String(row[1] || '').trim(),
        scheduled_at: scheduledAt,
        mail_type: mailType,
        subject: String(row[4] || '').trim()
          || tournamentName + String(row[1] || '').trim() + '\u3000案内',
        form_url: String(row[5] || '').trim() || null,
        schedule_ids: scheduleIds,
      });
    } catch (e) {
      errors.push({ row: rowNum, reason: e.message });
    }
  });

  return {
    candidates: candidates,
    skipped: skipped,
    errors: errors,
    summary: {
      candidate_count: candidates.length,
      skipped_count: skipped.length,
      error_count: errors.length,
    },
  };
}

function buildMailManagementSnapshotSyncPlan_(snapshot) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
  if (!sheet) throw new Error('メール管理シートが見つかりません。');

  const lastRow = sheet.getLastRow();
  const rows = lastRow < 6 ? [] : sheet.getRange(6, 1, lastRow - 5, 9).getValues();
  const now = new Date();
  const candidates = [];
  const skipped = [];
  const errors = [];
  const tournaments = snapshot.tournaments || [];

  rows.forEach((row, index) => {
    const rowNum = index + 6;
    const tournamentName = String(row[0] || '').trim();
    if (!tournamentName) return;
    if (String(row[7] || '').trim().includes('済')) {
      skipped.push({ row: rowNum, reason: '送信済み' });
      return;
    }
    const scheduledDate = mailManagementSyncDate_(row[2]);
    const scheduledAt = mailManagementSyncDateTime_(scheduledDate);
    if (!scheduledAt) {
      skipped.push({ row: rowNum, reason: '送信予定日時が未設定または日付ではありません' });
      return;
    }
    if (scheduledDate.getTime() <= now.getTime()) {
      skipped.push({ row: rowNum, reason: '送信予定日時が過去です' });
      return;
    }

    const mailTypeText = String(row[3] || '').trim();
    const mailType = mailTypeText === 'リマインダー'
      ? 'reminder'
      : (mailTypeText === '振込確認' ? 'payment_confirmation' : '');
    if (!mailType) {
      errors.push({ row: rowNum, reason: 'メール種別を特定できません' });
      return;
    }
    const tournamentMatches = tournaments.filter(item =>
      String(item.name || '').trim() === tournamentName
    );
    if (tournamentMatches.length !== 1) {
      errors.push({
        row: rowNum,
        reason: '同期スナップショット内の大会が' + tournamentMatches.length
          + '件あり、一意に特定できません',
      });
      return;
    }
    const grades = String(row[1] || '').replace(/級/g, '').replace(/\s+/g, '')
      .toUpperCase().split('').filter(Boolean);
    if (!grades.length || grades.some(grade => !/^[A-E]$/.test(grade))) {
      errors.push({ row: rowNum, reason: '級をA〜Eで特定できません' });
      return;
    }
    const targets = [];
    grades.forEach(grade => {
      const matches = (tournamentMatches[0].schedules || []).filter(schedule =>
        String(schedule.grade || '').toUpperCase() === grade
      );
      if (matches.length !== 1) {
        errors.push({
          row: rowNum,
          reason: grade + '級の日程が' + matches.length + '件あり、一意に特定できません',
        });
        return;
      }
      targets.push({ grade: grade, held_on: String(matches[0].held_on || '') });
    });
    if (targets.length !== grades.length) return;

    const subject = String(row[4] || '').trim()
      || tournamentName + String(row[1] || '').trim() + '\u3000案内';
    candidates.push({
      row: rowNum,
      tournament_name: tournamentName,
      grades: String(row[1] || '').trim(),
      scheduled_at: scheduledAt,
      mail_type: mailType,
      subject: subject,
      form_url: String(row[5] || '').trim() || null,
      targets: targets,
    });
  });

  return {
    candidates: candidates,
    skipped: skipped,
    errors: errors,
    summary: {
      candidate_count: candidates.length,
      skipped_count: skipped.length,
      error_count: errors.length,
    },
  };
}

function resolveMailManagementSnapshotPlan_(snapshotPlan) {
  const resolved = {
    candidates: [],
    skipped: snapshotPlan.skipped || [],
    errors: [],
  };
  (snapshotPlan.candidates || []).forEach(item => {
    const tournaments = taikaiApiRequest_('GET', '/tournaments', null, {
      name: item.tournament_name,
    }) || [];
    const exact = tournaments.filter(tournament =>
      String(tournament.name || '').trim() === item.tournament_name
    );
    if (exact.length !== 1) {
      resolved.errors.push({
        row: item.row,
        reason: '同期後の大会が' + exact.length + '件あり、一意に特定できません',
      });
      return;
    }
    const schedules = taikaiApiRequest_(
      'GET',
      '/tournaments/' + encodeURIComponent(String(exact[0].id)) + '/schedules'
    ) || [];
    const scheduleIds = [];
    (item.targets || []).forEach(target => {
      const matches = schedules.filter(schedule =>
        String(schedule.grade || '').toUpperCase() === target.grade
        && String(schedule.held_on || '') === target.held_on
      );
      if (matches.length !== 1) {
        resolved.errors.push({
          row: item.row,
          reason: target.grade + '級・' + target.held_on + 'の日程が'
            + matches.length + '件あり、一意に特定できません',
        });
        return;
      }
      scheduleIds.push(String(matches[0].id));
    });
    if (scheduleIds.length !== (item.targets || []).length) return;
    resolved.candidates.push(Object.assign({}, item, {
      schedule_ids: scheduleIds.sort((left, right) => taikaiCompareIds_(left, right)),
    }));
  });
  resolved.summary = {
    candidate_count: resolved.candidates.length,
    skipped_count: resolved.skipped.length,
    error_count: resolved.errors.length,
  };
  return resolved;
}

function mailManagementSameIds_(left, right) {
  const normalize = values => (values || []).map(String)
    .sort((a, b) => taikaiCompareIds_(a, b)).join(',');
  return normalize(left) === normalize(right);
}

function ensureMailManagementAnnouncement_(item) {
  const found = taikaiApiRequest_('GET', '/announcements', null, {
    subject: item.subject,
  }) || [];
  const exact = found.filter(announcement =>
    String(announcement.subject || '') === item.subject
  );
  if (exact.length > 1) {
    throw new Error('同じ件名の案内が複数あります。');
  }
  if (!exact.length) {
    return taikaiApiRequest_('POST', '/announcements', {
      subject: item.subject,
      form_url: item.form_url,
      gmail_message_id: null,
      schedule_ids: item.schedule_ids,
    });
  }

  let announcement = exact[0];
  if (String(announcement.form_url || '') !== String(item.form_url || '')) {
    announcement = taikaiApiRequest_(
      'PATCH',
      '/announcements/' + encodeURIComponent(String(announcement.id)),
      { form_url: item.form_url }
    );
  }
  if (!mailManagementSameIds_(announcement.schedule_ids, item.schedule_ids)) {
    announcement = taikaiApiRequest_(
      'PUT',
      '/announcements/' + encodeURIComponent(String(announcement.id)) + '/targets',
      { schedule_ids: item.schedule_ids }
    );
  }
  return announcement;
}

function previewMailManagementDatabaseSync() {
  try {
    return JSON.stringify(buildMailManagementDatabaseSyncPlan_());
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function executeMailManagementDatabaseSyncPlan_(plan) {
  if (plan.errors.length) {
    throw new Error('同期前検証でエラーがあります。プレビューを確認してください。');
  }
  const results = [];
  plan.candidates.forEach(item => {
      const announcement = item.mail_type === 'reminder'
        ? ensureMailManagementAnnouncement_(item)
        : null;
      const job = taikaiApiRequest_('POST', '/email-jobs', {
        scheduled_at: item.scheduled_at,
        mail_type: item.mail_type,
        announcement_id: announcement ? String(announcement.id) : null,
        schedule_ids: item.schedule_ids,
      });
      results.push({
        row: item.row,
        job_id: String(job.id),
        mail_type: item.mail_type,
        schedule_count: item.schedule_ids.length,
      });
  });
  return {
    ok: true,
    created_or_existing_count: results.length,
    skipped_count: plan.skipped.length,
    results: results,
  };
}

function syncMailManagementDatabase() {
  try {
    const plan = buildMailManagementDatabaseSyncPlan_();
    const result = executeMailManagementDatabaseSyncPlan_(plan);
    return JSON.stringify({
      ok: result.ok,
      created_or_existing_count: result.created_or_existing_count,
      skipped_count: result.skipped_count,
      results: result.results,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function mailManagementEscapeGmailQuery_(text) {
  return String(text || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function mailManagementMessageMatchesAnnouncement_(message, announcement) {
  if (!message) return false;
  const subject = String(announcement.subject || '').trim();
  if (!subject || String(message.getSubject() || '') !== subject) return false;
  const formUrl = String(announcement.form_url || '').trim();
  if (!formUrl) return true;
  const plain = String(message.getPlainBody ? message.getPlainBody() : '');
  const html = String(message.getBody ? message.getBody() : '');
  return plain.includes(formUrl) || html.includes(formUrl);
}

function mailManagementAnnouncementMessages_(announcement) {
  const subject = String(announcement.subject || '').trim();
  if (!subject) return [];
  const threads = GmailApp.search(
    'in:sent subject:"' + mailManagementEscapeGmailQuery_(subject) + '"',
    0,
    50
  );
  const matches = {};
  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      if (!mailManagementMessageMatchesAnnouncement_(message, announcement)) return;
      matches[String(message.getId())] = message;
    });
  });
  return Object.keys(matches).map(id => matches[id]);
}

function mailManagementGmailMessageById_(messageId) {
  if (!messageId) return null;
  try {
    return GmailApp.getMessageById(String(messageId)) || null;
  } catch (e) {
    return null;
  }
}

function getAnnouncementGmailAudit() {
  try {
    const source = taikaiApiRequest_('GET', '/announcement-audit-source') || {};
    const announcements = (source.announcements || []).map(announcement => {
      const currentId = String(announcement.gmail_message_id || '');
      const currentMessage = mailManagementGmailMessageById_(currentId);
      let currentStatus = currentId ? 'missing' : 'unlinked';
      if (currentMessage) {
        currentStatus = mailManagementMessageMatchesAnnouncement_(
          currentMessage,
          announcement
        ) ? 'valid' : 'mismatch';
      }
      const candidates = mailManagementAnnouncementMessages_(announcement).map(message => ({
        gmail_message_id: String(message.getId()),
        sent_at: Utilities.formatDate(message.getDate(), 'JST', 'yyyy-MM-dd HH:mm:ss'),
      }));
      return {
        id: String(announcement.id),
        subject: String(announcement.subject || ''),
        form_url: String(announcement.form_url || ''),
        schedule_ids: (announcement.schedule_ids || []).map(String),
        gmail_message_id: currentId || null,
        current_status: currentStatus,
        candidates: candidates,
      };
    });
    return JSON.stringify({
      announcements: announcements,
      uncovered_schedules: (source.uncovered_schedules || []).map(schedule => ({
        schedule_id: String(schedule.schedule_id),
        tournament_id: String(schedule.tournament_id),
        tournament_name: String(schedule.tournament_name || ''),
        grade: String(schedule.grade || ''),
        held_on: String(schedule.held_on || ''),
      })),
      summary: {
        announcement_count: announcements.length,
        valid_count: announcements.filter(item => item.current_status === 'valid').length,
        unresolved_count: announcements.filter(item => item.current_status !== 'valid').length,
        uncovered_schedule_count: (source.uncovered_schedules || []).length,
      },
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function setAnnouncementGmailLinkManual(json) {
  try {
    const request = JSON.parse(json);
    const announcementId = String(request.announcement_id || '').trim();
    const messageId = String(request.gmail_message_id || '').trim();
    if (!/^\d+$/.test(announcementId)) throw new Error('案内IDが不正です。');
    if (!messageId) throw new Error('Gmail IDを入力してください。');

    const announcement = taikaiApiRequest_(
      'GET',
      '/announcements/' + encodeURIComponent(announcementId)
    );
    const expected = request.expected_gmail_message_id == null
      ? null
      : String(request.expected_gmail_message_id);
    const current = announcement.gmail_message_id == null
      ? null
      : String(announcement.gmail_message_id);
    if (current !== expected) {
      throw new Error('Gmail IDが確認後に変更されています。監査結果を再読込してください。');
    }
    const verifiedCandidates = mailManagementAnnouncementMessages_(announcement);
    const message = verifiedCandidates.find(item =>
      String(item.getId()) === messageId
    );
    if (!message) {
      throw new Error(
        '指定されたGmail IDは、送信済みメールの件名・フォームURL完全一致候補ではありません。'
      );
    }
    const updated = taikaiApiRequest_(
      'PUT',
      '/announcements/' + encodeURIComponent(announcementId) + '/gmail-message',
      {
        gmail_message_id: messageId,
        expected_gmail_message_id: expected,
      }
    );
    return JSON.stringify({
      ok: true,
      announcement_id: announcementId,
      gmail_message_id: String(updated.gmail_message_id || messageId),
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function buildAnnouncementGmailLinkPlan_() {
  const announcements = taikaiApiRequest_('GET', '/announcements', null, {}) || [];
  const candidates = [];
  const skipped = [];
  const errors = [];
  announcements.forEach(announcement => {
    const id = String(announcement.id);
    const subject = String(announcement.subject || '').trim();
    if (announcement.gmail_message_id) {
      skipped.push({ announcement_id: id, subject: subject, reason: '紐付け済み' });
      return;
    }
    const messages = mailManagementAnnouncementMessages_(announcement);
    if (messages.length !== 1) {
      errors.push({
        announcement_id: id,
        subject: subject,
        reason: messages.length
          ? '一致する送信済みGmailが複数あります'
          : '一致する送信済みGmailが見つかりません',
        match_count: messages.length,
      });
      return;
    }
    candidates.push({
      announcement_id: id,
      subject: subject,
      gmail_message_id: String(messages[0].getId()),
    });
  });
  return {
    candidates: candidates,
    skipped: skipped,
    errors: errors,
    summary: {
      candidate_count: candidates.length,
      skipped_count: skipped.length,
      error_count: errors.length,
    },
  };
}

function announcementGmailLinkPublicPlan_(plan) {
  return {
    candidates: plan.candidates.map(item => ({
      announcement_id: item.announcement_id,
      subject: item.subject,
      match_count: 1,
    })),
    skipped: plan.skipped,
    errors: plan.errors,
    summary: plan.summary,
  };
}

function previewAnnouncementGmailLinks() {
  try {
    return JSON.stringify(
      announcementGmailLinkPublicPlan_(buildAnnouncementGmailLinkPlan_())
    );
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function linkAnnouncementGmailMessages() {
  try {
    const plan = buildAnnouncementGmailLinkPlan_();
    if (plan.errors.length) {
      throw new Error('紐付け前検証でエラーがあります。プレビューを確認してください。');
    }
    const results = plan.candidates.map(item => {
      taikaiApiRequest_(
        'PATCH',
        '/announcements/' + encodeURIComponent(item.announcement_id),
        { gmail_message_id: item.gmail_message_id }
      );
      return {
        announcement_id: item.announcement_id,
        subject: item.subject,
        linked: true,
      };
    });
    return JSON.stringify({
      ok: true,
      linked_count: results.length,
      already_linked_count: plan.skipped.length,
      results: results,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

// メール管理シートの行を削除
function deleteMailManagementRow(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    if (!sheet) throw new Error('メール管理シートが見つかりません');
    sheet.deleteRow(rowNum);
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// メール管理シートに行を追加
function addMailManagementRow(json) {
  try {
    const d     = JSON.parse(json);
    let emailJob = null;
    if (d.mailType === '振込確認') {
      // DB登録に失敗した場合はシートへ成功行を作らず、同じ入力で再試行できるようにする。
      emailJob = taikaiRegisterPaymentEmailJob_(
        d.tournamentName,
        d.grades,
        d.sendDateTime,
        d.paymentDeadline
      );
    }
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    const lastRow = Math.max(sheet.getLastRow(), 5);
    sheet.getRange(lastRow + 1, 1, 1, 9).setValues([[
      d.tournamentName || '',
      d.grades         || '',
      d.sendDateTime   || '',
      d.mailType       || 'リマインダー',
      d.threadTitle    || '',
      d.formLink       || '',
      d.mailType === '振込確認' ? '済' : '',
      '',
      false,
    ]]);

    // 振込確認の場合、カレンダーシートにも振込確認送信日を記録（col 9）
    if (d.mailType === '振込確認' && d.tournamentName && d.sendDateTime) {
      try {
        const calSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
        if (calSheet) {
          const fullName = (d.tournamentName || '') + (d.grades || '');
          const calData  = calSheet.getRange(1, 1, calSheet.getLastRow(), 1).getValues();
          for (let i = 2; i < calData.length; i++) {
            if (String(calData[i][0]) === fullName) {
              const dateStr = d.sendDateTime.slice(0, 10).replace(/-/g, '/');
              calSheet.getRange(i + 1, 9).setValue(dateStr);
              break;
            }
          }
        }
      } catch(calErr) {
        // カレンダー更新のエラーは無視して続行
      }
    }

    return JSON.stringify({ ok: true, emailJob: emailJob });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// ------------------------------------------------
// 内部: 参加者リスト構築（makeParticipantList 相当）
// ------------------------------------------------
function makeParticipantList_(tournamentName, grades, includeNotPaid) {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tournamentName + grades);
  if (!sheet) return '（シートが見つかりません）\n';

  let structure;
  try {
    structure = tournamentSheetStructure_(sheet, false);
  } catch (e) {
    return '（大会シート構造を特定できません）\n';
  }
  const responseData = tournamentSheetResponseRowsWithStatus_(structure);
  const paymentStatusIndex = responseData.payment_status_index;
  const data = latestFormRowsByPlayer_(responseData.rows)
    .filter(row => row[2] !== '');

  const gradesArray = grades.replace('級', '').split('');
  const byGrade = {};
  gradesArray.forEach(g => { byGrade[g + '級'] = []; });
  const paidList = [];

  data.forEach(row => {
    const name      = String(row[2]).replace('　', ' ');
    const gradeStr  = String(row[4]);
    const isPaid    = String(row[paymentStatusIndex] || '').trim() === '済';

    if (isPaid && !paidList.includes(name)) paidList.push(name);

    gradeStr.replace('級', '').split('').forEach(g => {
      const key = g + '級';
      if (byGrade[key] && !byGrade[key].includes(name)) byGrade[key].push(name);
    });
  });

  const total = Object.values(byGrade).reduce((s, a) => s + a.length, 0);
  let result = '';
  if (total === 0) {
    result = '現在、どなたからも申し込みは頂いておりません。\n';
  } else {
    result += '現在、以下の方々からご連絡を頂いております。(敬称略)\n';
    for (const grade in byGrade) {
      result += grade + '：' + byGrade[grade].join('、') + '\n';
    }
    if (includeNotPaid && paidList.length > 0) {
      result +=
        '\nそのうち、振込の完了を確認できている方：\n' + paidList.join('、') + '\n\n' +
        '確認漏れがございましたら申し訳ありません。\n' +
        '明日までに振込を確認できない場合はキャンセルとせざるを得ない場合がありますので、ご了承ください。\n' +
        'ご相談等がある場合は、このメールに返信するか、会長 / 副会長までご連絡ください。\n';
    }
  }
  return result;
}

// 内部: Gmailから件名で最新スレッドを引用（quoteEmail 相当）
function quoteEmail_(subject) {
  if (!subject) return '';
  try {
    const normalizeSubject = value => String(value || '')
      .replace(/^(?:(?:Fwd|Re):\s*)+/i, '')
      .replace(/【[^】]*】/g, '')
      .replace(/(?:　案内|の案内)/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    const expected = normalizeSubject(subject);
    const searchText = expected.replace(/"/g, '');
    const threads = GmailApp.search('subject:"' + searchText + '"', 0, 20);
    for (let i = 0; i < threads.length; i++) {
      const messages = threads[i].getMessages();
      const mentioned = messages.find(message =>
        normalizeSubject(message.getSubject()) === expected
      );
      if (!mentioned) continue;
      return mentioned.getPlainBody().split('\n').map(line => '> ' + line).join('\n');
    }
    return '';
  } catch(e) {
    return '';
  }
}

// 内部: メール本文生成（createEmailBody 相当）
function createEmailBody_(tournamentName, grades, participantList, tomorrowDateStr, formLink, mailType, quotedContent) {
  let body = 'おはようございます。\n\nこのメールは、慶應かるた会のメールシステムにより自動で送信されています。\n\n';

  if (mailType === 'リマインダー') {
    body +=
      'こちらは、' + tournamentName + grades + 'のリマインダーです。\n\n' +
      participantList + '\n' +
      'これから参加表明をされる方は、明日' + tomorrowDateStr + 'までに以下のリンクからフォームにご回答ください。\n' +
      formLink + '\n\n' +
      'よろしくお願いします。';
  } else {
    body +=
      'こちらは、' + tournamentName + grades + 'の振込確認です。\n\n' +
      participantList + '\n' +
      'まだお済みでない方は、お早めにお振込みください。\n\n' +
      'よろしくお願いします。';
  }

  if (quotedContent) {
    body += '\n\n以下、先日の本件に関するメールのコピーです。\n' + quotedContent;
  }
  return body;
}

// リマインダーメール本文プレビュー取得
function getReminderPreview(json) {
  try {
    const { rowNum, includeNotPaid, sendDateTime } = JSON.parse(json);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    const row   = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];

    const tournamentName = String(row[0]);
    const grades         = String(row[1]);
    const sheetDateStr   = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'JST', 'yyyy-MM-dd HH:mm:ss')
      : String(row[2] || '');
    const sendDateStr = sendDateTime || sheetDateStr;
    const mailType    = String(row[3]);
    const threadTitle = String(row[4]);
    const formLink    = String(row[5]);

    if (!sendDateStr) return JSON.stringify({ error: '送信予定日時が未設定です' });

    const targetDate = new Date(sendDateStr);
    if (targetDate - new Date() < 60 * 60 * 1000) return JSON.stringify({ error: '送信予定日時は現在時刻から1時間以上後に設定してください' });

    const participantList = makeParticipantList_(tournamentName, grades, includeNotPaid);

    const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];
    const tomorrow = new Date(targetDate.getTime());
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowDateStr = (tomorrow.getMonth() + 1) + '月' + tomorrow.getDate() + '日（' + WEEKDAYS[tomorrow.getDay()] + '）';

    const quotedContent = quoteEmail_(threadTitle);
    const body          = createEmailBody_(tournamentName, grades, participantList, tomorrowDateStr, formLink, mailType, quotedContent);
    const sendDateFormatted = Utilities.formatDate(targetDate, 'JST', 'yyyy年MM月dd日 HH時mm分');

    return JSON.stringify({ ok: true, body, sendDateFormatted, tournamentName, grades, mailType });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// リマインダートリガー設定（setReminders 相当）
function setReminderTrigger(json) {
  try {
    const { rowNum, includeNotPaid, sendDateTime } = JSON.parse(json);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIL);
    const row   = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];

    // モーダルで指定された日時があればシートに上書き
    if (sendDateTime) sheet.getRange(rowNum, 3).setValue(sendDateTime);

    const sendDateStr = sendDateTime || (row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'JST', 'yyyy-MM-dd HH:mm:ss')
      : String(row[2] || ''));
    const targetDate = new Date(sendDateStr);
    if (targetDate - new Date() < 60 * 60 * 1000) return JSON.stringify({ error: '送信予定日時は現在時刻から1時間以上後に設定してください' });

    sheet.getRange(rowNum, 7, 1, 3).setValues([['済', '', includeNotPaid]]);

    const tournamentName = String(row[0]);
    const grades         = String(row[1]);
    const mailType       = String(row[3]);

    // リマインダーの場合: フォーム説明更新 + カレンダー更新
    if (mailType === 'リマインダー') {
      try {
        const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];
        const tomorrow = new Date(targetDate.getTime());
        tomorrow.setDate(tomorrow.getDate() + 1);
        const tomorrowStr = (tomorrow.getMonth() + 1) + '月' + tomorrow.getDate() + '日（' + WEEKDAYS[tomorrow.getDay()] + '）';

        const tSheet = ss.getSheetByName(tournamentName + grades);
        if (tSheet) {
          const structure = tournamentSheetStructure_(tSheet, false);
          const formId = tournamentSheetFormId_(structure);
          FormApp.openById(formId).setDescription(
            'こちらは' + tournamentName + grades + 'の参加表明フォームです。\n' +
            '該当項目に回答の上、送信してください。\n' +
            'このフォームの回答期限は【' + tomorrowStr + '】の23:59までです。\n' +
            '回答が正しく送信されている場合、入力いただいたメールアドレスに回答のコピーが届きますのでご確認ください。'
          );
        }

        const calSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CALENDAR);
        if (calSheet) {
          const calData = calSheet.getRange(1, 1, calSheet.getLastRow(), 1).getValues();
          for (let i = 1; i < calData.length; i++) {
            if (String(calData[i][0]) === tournamentName + grades) {
              calSheet.getRange(i + 1, 4).setValue(Utilities.formatDate(targetDate, 'JST', 'yyyy/MM/dd'));
              const moshikomiStart = new Date(targetDate.getTime() + 48 * 60 * 60 * 1000);
              calSheet.getRange(i + 1, 5).setValue(Utilities.formatDate(moshikomiStart, 'JST', 'yyyy/MM/dd'));
              break;
            }
          }
        }
      } catch(innerErr) {
        // フォーム/カレンダー更新のエラーは無視して続行
      }
    }

    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}
