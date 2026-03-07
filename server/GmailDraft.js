// ============================================================
// Gmail 下書き作成
// ============================================================

// タブ0：案内メール
function createDraft1(json) {
  try {
    const d = JSON.parse(json);
    const subject = (d.title || '') + (d.grades || '') + 'の案内';
    GmailApp.createDraft('', subject, d.body || '');
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// タブ1：振込案内
function createDraft2(json) {
  try {
    const d = JSON.parse(json);
    const subject = (d.title || '') + (d.grades || '') + '\u3000出場者確定のお知らせ';
    GmailApp.createDraft('', subject, d.body || '');
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// タブ2：Free
function createDraft3(json) {
  try {
    const d = JSON.parse(json);
    GmailApp.createDraft('', '', d.body || '');
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// タブ3：読手講習会など
function createDraft4(json) {
  try {
    const d = JSON.parse(json);
    const subject = (d.subject || '') + '　案内';
    GmailApp.createDraft('', subject, d.body || '');
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}
