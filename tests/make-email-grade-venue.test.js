const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const page = fs.readFileSync(
  path.join(root, 'pages/make-email.html'), 'utf8'
);
const script = fs.readFileSync(
  path.join(root, 'scripts/make-email.html'), 'utf8'
);

['A', 'B', 'C', 'D', 'E'].forEach(grade => {
  assert.match(page, new RegExp('id="me-venue-row-' + grade + '"'));
  assert.match(page, new RegExp('id="me-venue-' + grade + '"'));
});
assert.match(page, /id="me-venue-bulk"/);
assert.doesNotMatch(page, /id="me-venue"/);
assert.match(script, /function meBulkVenueChange\(/);
assert.match(script, /\['date', 'venue', 'cap', 'elig', 'fee'\]/);
assert.match(script, /const venueLines = active\.map/);
assert.match(script, /`場所：\\n\$\{venueLines\}\\n\\n`/);
assert.doesNotMatch(script, /getElementById\('me-venue'\)/);

console.log('Make-email grade venue regression checks passed.');
