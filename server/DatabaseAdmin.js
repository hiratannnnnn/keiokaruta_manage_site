// ============================================================
// DB管理画面用 API コンソール
// ============================================================

function databaseAdminAllowedPath_(method, path) {
  const rules = {
    GET: [
      /^\/admin\/database\/[a-z-]+$/,
      /^\/tournaments$/,
      /^\/tournaments\/\d+$/,
      /^\/tournaments\/\d+\/schedules$/,
      /^\/schedules$/,
      /^\/schedules\/\d+$/,
      /^\/players$/,
      /^\/players\/\d+$/,
      /^\/players\/\d+\/participation-summary$/,
      /^\/players\/\d+\/participations$/,
      /^\/entries$/,
      /^\/entries\/\d+$/,
      /^\/announcements$/,
      /^\/announcements\/\d+$/,
      /^\/recipient-groups$/,
      /^\/recipient-groups\/\d+$/,
      /^\/email-jobs$/,
      /^\/email-jobs\/\d+$/,
      /^\/email-jobs\/\d+\/dispatch-context$/,
      /^\/email-deliveries$/,
      /^\/email-deliveries\/\d+$/,
    ],
    POST: [
      /^\/tournaments$/,
      /^\/tournaments\/\d+\/schedules$/,
      /^\/players$/,
      /^\/registrations$/,
      /^\/entries$/,
      /^\/announcements$/,
      /^\/recipient-groups$/,
      /^\/email-jobs$/,
    ],
    PATCH: [
      /^\/tournaments\/\d+$/,
      /^\/schedules\/\d+$/,
      /^\/players\/\d+$/,
      /^\/entries\/\d+$/,
      /^\/announcements\/\d+$/,
      /^\/recipient-groups\/\d+$/,
      /^\/email-jobs\/\d+$/,
    ],
    PUT: [
      /^\/entries\/\d+\/payment$/,
      /^\/entries\/\d+\/cancellation$/,
      /^\/announcements\/\d+\/targets$/,
      /^\/recipient-groups\/\d+\/members$/,
    ],
    DELETE: [
      /^\/tournaments\/\d+$/,
    ],
  };
  return (rules[method] || []).some(pattern => pattern.test(path));
}

function databaseAdminAuthenticate_(password) {
  const stored = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!stored || String(password || '') !== stored) {
    throw new Error('DB管理権限の認証に失敗しました。');
  }
}

function databaseAdminRequest(json) {
  try {
    const input = JSON.parse(json);
    databaseAdminAuthenticate_(input.password);

    const method = String(input.method || 'GET').toUpperCase();
    const path = String(input.path || '').trim();
    if (!databaseAdminAllowedPath_(method, path)) {
      throw new Error('このメソッドとAPIパスはDB管理画面から実行できません。');
    }

    const query = input.query && typeof input.query === 'object' ? input.query : {};
    const body = input.body === undefined ? null : input.body;
    const data = taikaiApiRequest_(method, path, body, query);
    return JSON.stringify({ ok: true, data: data });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
