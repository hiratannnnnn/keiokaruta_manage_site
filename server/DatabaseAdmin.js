// ============================================================
// DB管理画面用 API コンソール
// ============================================================

function databaseAdminAllowedPath_(method, path) {
  const rules = {
    GET: [
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
      /^\/admin\/fiscal-year-sync$/,
      /^\/admin\/player-email-migration$/,
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

function databaseAdminPseudonymizePlayerEmail_(player) {
  if (!player || typeof player !== 'object' ||
      !Object.prototype.hasOwnProperty.call(player, 'email')) {
    return player;
  }
  const email = String(player.email || '').trim();
  if (!email || isPseudonymousEmail_(email)) return player;

  const clean = Object.assign({}, player);
  const playerName = [
    String(clean.family_name || '').trim(),
    String(clean.given_name || '').trim(),
  ].filter(Boolean).join(' ');
  clean.email = rememberPseudonymousEmail_(email, playerName);
  return clean;
}

function databaseAdminProtectPlayerEmail_(method, path, body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return body;

  if ((method === 'POST' && path === '/players') ||
      (method === 'PATCH' && /^\/players\/\d+$/.test(path))) {
    return databaseAdminPseudonymizePlayerEmail_(body);
  }
  if (method === 'POST' && path === '/registrations' &&
      body.player && typeof body.player === 'object') {
    const clean = Object.assign({}, body);
    clean.player = databaseAdminPseudonymizePlayerEmail_(body.player);
    return clean;
  }
  return body;
}

function databaseAdminRequest(json) {
  try {
    const input = JSON.parse(json);
    const method = String(input.method || 'GET').toUpperCase();
    const path = String(input.path || '').trim();
    if (!databaseAdminAllowedPath_(method, path)) {
      throw new Error('このメソッドとAPIパスはDB管理画面から実行できません。');
    }

    const query = input.query && typeof input.query === 'object' ? input.query : {};
    const body = databaseAdminProtectPlayerEmail_(
      method,
      path,
      input.body === undefined ? null : input.body
    );
    const data = taikaiApiRequest_(method, path, body, query);
    return JSON.stringify({ ok: true, data: data });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
