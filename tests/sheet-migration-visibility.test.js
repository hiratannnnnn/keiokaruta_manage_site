const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const page = fs.readFileSync(path.join(root, 'pages/settings.html'), 'utf8');
const settings = fs.readFileSync(path.join(root, 'scripts/settings.html'), 'utf8');
const utils = fs.readFileSync(path.join(root, 'scripts/utils.html'), 'utf8');

assert.match(index, /id="nav-sheet-migration-btn" style="display:none;"/);
assert.match(page, /id="settings-show-sheet-migration"/);
assert.match(page, /タブを開くと再び非表示になります/);
assert.match(page, /実行機能の安全ゲートは別途維持されます/);
assert.match(settings, /const SHEET_MIGRATION_NAV_KEY = 'karuta_show_sheet_migration'/);
assert.match(settings, /currentRole === '会長副会長'\s*&& isSheetMigrationNavigationEnabled\(\)/);
assert.match(settings, /sessionStorage\.removeItem\(SHEET_MIGRATION_NAV_KEY\)/);
assert.match(settings, /function consumeSheetMigrationNavigationActivation/);
assert.match(utils, /applySheetMigrationNavigationVisibility\(\)/);
assert.match(utils, /initSheetMigration\(\);\s*[\s\S]*?consumeSheetMigrationNavigationActivation\(\)/);
assert.doesNotMatch(
  utils,
  /nav-sheet-migration-btn'\)\.style\.display\s*=\s*currentRole/,
  '役職だけでシート移行を表示してはいけません'
);

console.log('Sheet migration visibility regression checks passed.');
