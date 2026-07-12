const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const vm = require('node:vm');

const python = process.env.PYTHON || 'python3';
const html = execFileSync(
  python,
  ['-c', "import local_app; print(local_app.render_index_html())"],
  { encoding: 'utf8' },
);
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
assert.ok(script.includes('`${safeName}_Timeline.xlsx`'), 'download suffix must use capital Timeline');
assert.ok(script.includes('calendar-day-note'), 'date picker must visibly mark adjusted workdays');
const calendarStart = script.indexOf('const CHINA_PUBLIC_HOLIDAYS');
const calendarEnd = script.indexOf('\n    function daysLabel', calendarStart);
assert.notEqual(calendarStart, -1, 'calendar workday rules must be present');
assert.notEqual(calendarEnd, -1, 'calendar workday rules must end before labels');
const calendarContext = {};
vm.runInNewContext(script.slice(calendarStart, calendarEnd), calendarContext);
assert.equal(calendarContext.isTimelineWorkday(new Date(2026, 8, 20)), true);
assert.equal(calendarContext.countWorkdays('2026-09-18', '2026-09-21'), 3);
assert.equal(calendarContext.isTimelineWorkday(new Date(2026, 8, 25)), false);
const start = script.indexOf('function reorderStageTaskBlocks');
const end = script.indexOf('\n    function clearStageDragIndicators', start);
assert.notEqual(start, -1, 'stage reorder helper must be present');
assert.notEqual(end, -1, 'stage reorder helper must end before DOM bindings');

const context = {
  normalizeStage(value) {
    return String(value || '').trim() || 'Uncategorized';
  },
};
vm.runInNewContext(script.slice(start, end), context);

const source = [
  { stage: 'Requirement', name: 'scope' },
  { stage: 'Development', name: 'build-a' },
  { stage: 'Proposal', name: 'concept' },
  { stage: 'Development', name: 'build-b' },
  { stage: 'Requirement', name: 'sign-off' },
];
const reordered = context.reorderStageTaskBlocks(source, ['Development', 'Proposal', 'Requirement']);

assert.deepEqual(
  reordered.map(task => `${task.stage}:${task.name}`),
  ['Development:build-a', 'Development:build-b', 'Proposal:concept', 'Requirement:scope', 'Requirement:sign-off'],
);
console.log('stage tab reorder: PASS');
