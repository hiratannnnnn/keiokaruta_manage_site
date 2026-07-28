const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const utils = fs.readFileSync(path.join(root, 'scripts/utils.html'), 'utf8');
const members = fs.readFileSync(path.join(root, 'scripts/members.html'), 'utf8');
const calendar = fs.readFileSync(path.join(root, 'scripts/calendar.html'), 'utf8');
const suitou = fs.readFileSync(path.join(root, 'scripts/suitou.html'), 'utf8');
const playerPage = fs.readFileSync(path.join(root, 'pages/player.html'), 'utf8');

const initAppBody = utils.match(/function initApp\(\) \{([\s\S]*?)\n\}/);
assert.ok(initAppBody, 'initApp must exist');
assert.doesNotMatch(initAppBody[1], /loadMembers\(/);
assert.match(utils, /if \(calendarLoaded\) \{\s*renderCalendar\(allTournaments\);\s*\} else \{\s*loadCalendar\(\)/);
assert.match(utils, /let calendarLoading\s*=\s*false/);
assert.match(calendar, /if \(calendarLoading\) return/);
assert.match(calendar, /function onCalendarLoaded[\s\S]*?calendarLoading = false/);
assert.match(calendar, /function onCalendarLoaded[\s\S]*?calendarLoaded = true/);
assert.match(utils, /if \(page === 'player'\) ensureMembersLoaded\(\)/);
assert.match(members, /function ensureMembersLoaded/);
assert.match(members, /if \(!membersLoading\) loadMembers\(false\)/);
assert.match(calendar, /ensureMembersLoaded\(\(\) => goToPlayerSearchWithMembers\(name\)\)/);
assert.match(suitou, /function suShowAddModal[\s\S]*?ensureMembersLoaded\(\)/);
assert.match(playerPage, /onclick="loadMembers\(true\)"/);

console.log('Lazy loading regression checks passed.');
