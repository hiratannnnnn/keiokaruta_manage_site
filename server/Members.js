// ============================================================
// 名簿関連
// ============================================================

// 選手マスタはDBから取得するが、画面に出す実メールは非公開対応表から逆引きする。
function getMembers() {
  try {
    const players = taikaiApiRequest_('GET', '/players', null, {}) || [];
    const members = players.map(player => ({
      '氏名': String(player.family_name || '') + ' ' + String(player.given_name || ''),
      'ふりがな': player.ruby || '',
      'メールアドレス': realEmailForPseudonym_(player.email),
      '所属': player.club || '',
    }));
    return JSON.stringify(members);
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
