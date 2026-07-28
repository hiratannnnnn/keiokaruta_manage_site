// ============================================================
// 環境別設定。ローカル.envと同名のScript Propertiesから取得する。
// ============================================================

function configValue_(propertyName) {
  const value = PropertiesService.getScriptProperties().getProperty(propertyName);
  const normalized = String(value || '').trim();
  if (!normalized) {
    throw new Error(propertyName + ' がScript Propertiesに設定されていません。');
  }
  return normalized;
}

const CONFIG = {
  get SPREADSHEET_ID() {
    return configValue_('MAIN_SPREADSHEET_ID');
  },
  get FORM_FOLDER_TO() {
    return configValue_('FORM_FOLDER_ID');
  },
  get FORM_TEMPLATE_ID() {
    return configValue_('FORM_TEMPLATE_ID');
  },
  get TRASH_SPREADSHEET_ID() {
    return configValue_('TRASH_SPREADSHEET_ID');
  },
  get BOARD_SPREADSHEET_ID() {
    return configValue_('BOARD_SPREADSHEET_ID');
  },

  // スプレッドシート内のシート名
  SHEET_NAMES: {
    MEMBERS:  '名簿',
    EMAIL_MAP: 'DBメール対応表',
    CALENDAR: 'カレンダー',
    MAIL:     'メール管理',
    SETTINGS: '設定用',
  },

};
