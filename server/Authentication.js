// ============================================================
// 認証候補の確認
// ============================================================

function authenticationNormalizedName_(value) {
  return String(value || '').replace(/[\s\u3000]+/g, '').trim();
}

function authenticationDuplicateGroups_(rows, nameIndex, emailIndex, firstDataRow) {
  const byName = {};
  (rows || []).forEach((row, index) => {
    const displayName = String(row[nameIndex] || '').trim();
    const normalizedName = authenticationNormalizedName_(displayName);
    const email = String(row[emailIndex] || '').trim().toLowerCase();
    if (!normalizedName || !email) return;
    if (!byName[normalizedName]) {
      byName[normalizedName] = {
        name: displayName,
        emails: {},
        rows: [],
      };
    }
    byName[normalizedName].emails[email] = true;
    byName[normalizedName].rows.push(index + firstDataRow);
  });
  return Object.keys(byName).map(key => byName[key])
    .filter(group => Object.keys(group.emails).length > 1)
    .map(group => ({
      name: group.name,
      emails: Object.keys(group.emails).sort(),
      rows: group.rows,
    }))
    .sort((left, right) => left.name.localeCompare(right.name, 'ja'));
}

function getDuplicateNameEmailGroups() {
  try {
    const sheet = emailMapSheet_(false);
    if (!sheet) {
      throw new Error('DBメール対応表が見つかりません。');
    }
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      .map(value => String(value || '').trim());
    const nameIndex = headers.indexOf('氏名');
    const emailIndex = headers.indexOf('実メールアドレス');
    if (nameIndex < 0 || emailIndex < 0) {
      throw new Error('DBメール対応表に「氏名」または「実メールアドレス」列がありません。');
    }
    const lastRow = sheet.getLastRow();
    const rows = lastRow < 2
      ? []
      : sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
    return JSON.stringify({
      groups: authenticationDuplicateGroups_(rows, nameIndex, emailIndex, 2),
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
