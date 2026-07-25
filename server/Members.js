// ============================================================
// 名簿関連
// ============================================================

// 名簿取得（taikai_manage API）
// フォーム回答・大会登録で共有する選手マスタをDBから取得する。
function getMembers() {
  try {
    const players = taikaiApiRequest_('GET', '/players', null, {}) || [];
    const members = players.map(player => ({
      '氏名': String(player.family_name || '') + ' ' + String(player.given_name || ''),
      'ふりがな': player.ruby || '',
      'メールアドレス': player.email || '',
      '所属': player.club || '',
    }));
    return JSON.stringify(members);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
