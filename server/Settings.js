// ============================================================
// 権限設定・パスワード認証
// ============================================================

var DEFAULT_PERMISSIONS_ = {
  applyDoneInput:          ['申込係', '会長副会長'],
  payDoneInput:            ['振込係', '会長副会長'],
  completeTournament:      ['会長副会長'],
  payDetailButtons:        ['振込係', '会長副会長'],
  operationPanel:          ['申込係', '振込係', '案内係', '会長副会長'],
  announcementDoubleCheck: ['案内係', '会長副会長'],
  deleteTournament:        ['会長副会長'],
};

// 権限設定取得（PropertiesService の 'PERMISSIONS' キー）
function getPermissions() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty('PERMISSIONS');
    if (stored) {
      // 新規キーはデフォルト値で補完しつつ、保存済み値を優先
      const parsed = JSON.parse(stored);
      const result = {};
      Object.keys(DEFAULT_PERMISSIONS_).forEach(key => {
        result[key] = key in parsed ? parsed[key] : DEFAULT_PERMISSIONS_[key];
      });
      return JSON.stringify(result);
    }
    return JSON.stringify(DEFAULT_PERMISSIONS_);
  } catch(e) {
    Logger.log('getPermissions error: ' + e.message);
    return JSON.stringify(DEFAULT_PERMISSIONS_);
  }
}

// 権限設定保存
function savePermissions(permJson) {
  try {
    PropertiesService.getScriptProperties().setProperty('PERMISSIONS', permJson);
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

// 役職者苗字の取得（PropertiesService の 'ROLE_NAMES' キー）
function getRoleNames() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty('ROLE_NAMES');
    const defaults = { 会長: '', 副会長: '', 申込係: '', 振込係: '', 案内係: '' };
    return JSON.stringify(Object.assign(defaults, stored ? JSON.parse(stored) : {}));
  } catch(e) {
    return JSON.stringify({ 会長: '', 副会長: '', 申込係: '', 振込係: '', 案内係: '' });
  }
}

// 役職者苗字の保存
function saveRoleNames(json) {
  try {
    PropertiesService.getScriptProperties().setProperty('ROLE_NAMES', json);
    return JSON.stringify({ ok: true });
  } catch(e) {
    return JSON.stringify({ error: e.message });
  }
}

function getEnvironmentConfigurationStatus() {
  try {
    const props = PropertiesService.getScriptProperties();
    const resourceKeys = [
      'MAIN_SPREADSHEET_ID',
      'FORM_FOLDER_ID',
      'FORM_TEMPLATE_ID',
      'TRASH_SPREADSHEET_ID',
      'BOARD_SPREADSHEET_ID',
    ];
    const requiredKeys = resourceKeys.concat([
      'TAIKAI_API_BASE_URL',
      'TAIKAI_API_ALERT_EMAIL',
      'PSEUDO_EMAIL_SECRET',
    ]);
    const items = requiredKeys.map(key => ({
      key: key,
      configured: Boolean(String(props.getProperty(key) || '').trim()),
      fallback: false,
    }));
    const configuredResources = items
      .filter(item => resourceKeys.includes(item.key) && item.configured).length;
    const warnings = [];
    if (configuredResources < resourceKeys.length) {
      warnings.push('未登録の外部リソースIDがあります。登録するまで関連機能は実行できません。');
    }
    return JSON.stringify({
      items: items,
      warnings: warnings,
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}
