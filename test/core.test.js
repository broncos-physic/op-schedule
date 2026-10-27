const test = require('node:test');
const assert = require('node:assert');
const XLSX = require('xlsx');
const C = require('../core.js');

// A small synthetic workbook in the same layout as the real schedule.
function sheet(rows) { return XLSX.utils.aoa_to_sheet(rows); }
function workbook() {
  const days = Array.from({ length: 31 }, (_, i) => i + 1);
  const dow = ['TH', 'F', 'SA', 'SU', 'M', 'T', 'W'];
  const oct = [
    ['junk header text', ...days],
    [' ', ...days.map((_, i) => dow[i % 7])],
    ['Pat  Q', 'A', 'A', 0, 'M', 'L', 'X', 'S', '3:30-8', '4', ' ', '=-B7'],
    [' '],
    ['Sam', 'E', 'D', '2', 0, 'A', 'A'],
    [null, 'A=3:30-12'], // legend line: no name in column A
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet(oct), 'Oct 26 A M');
  XLSX.utils.book_append_sheet(wb, sheet([['x']]), 'OCT 26-T');
  XLSX.utils.book_append_sheet(wb, sheet(oct), 'June-26 A and M');
  return wb;
}

test('only A M sheets are read, month comes from the sheet name', () => {
  const d = C.parseWorkbook(XLSX, workbook(), C.DEFAULT_SHEET_PATTERN);
  assert.deepStrictEqual(d.months.map((m) => [m.sheet, m.year, m.month]), [['June-26 A and M', 2026, 5], ['Oct 26 A M', 2026, 9]]);
});

test('names are normalized, legend and spacer rows are ignored, formulas are empty', () => {
  const d = C.parseWorkbook(XLSX, workbook(), C.DEFAULT_SHEET_PATTERN);
  assert.deepStrictEqual([...d.people.keys()].sort(), ['pat q', 'sam']);
  const pat = d.people.get('pat q');
  assert.strictEqual(pat.name, 'Pat Q');
  assert.strictEqual(pat.days.get('2026-10-01').raw, 'A');
  assert.strictEqual(pat.days.has('2026-10-11'), false); // blank cell
  assert.strictEqual(pat.days.has('2026-10-12'), false); // formula cell
});

test('cell codes resolve to shifts', () => {
  const r = (x) => C.resolveCell(x, C.DEFAULT_SHIFTS);
  assert.strictEqual(r('0'), null);
  assert.strictEqual(r(''), null);
  assert.strictEqual(r('O'), null);
  assert.strictEqual(r('a').startMin, 15 * 60 + 30);
  assert.strictEqual(r('A').endMin, 1440); // 12 AM
  assert.strictEqual(r('M').endMin, 8 * 60 + 1440); // runs into the next morning
  assert.strictEqual(r('2').endMin, 22 * 60 + 30);
  assert.strictEqual(r('E').startMin, 6 * 60);
  assert.strictEqual(r('L').kind, 'timeoff');
  assert.strictEqual(r('x').kind, 'sick');
  assert.strictEqual(r('S').kind, 'allday');
  assert.strictEqual(r('4').kind, 'unknown');
  const custom = r('3:30-8');
  assert.deepStrictEqual([custom.startMin, custom.endMin], [15 * 60 + 30, 20 * 60]);
});

test('ics: overnight shift ends next day, DST-safe TZID times, escaping and CRLF', () => {
  const d = C.parseWorkbook(XLSX, workbook(), C.DEFAULT_SHEET_PATTERN);
  const ev = C.buildPersonEvents(d.people.get('pat q'), C.DEFAULT_SHIFTS, { from: '2026-10-01', to: '2026-10-31', prefix: 'Work' });
  assert.strictEqual(ev.length, 6); // A, A, M, S (all-day), custom 3:30-8, lone '4'
  const ics = C.buildICS(ev, { calName: 'Pat Q', location: '2263 Fuller Ct, Ann Arbor, MI', reminderMin: 60, now: new Date('2026-10-26T12:00:00Z') });
  assert.match(ics, /DTSTART;TZID=America\/Detroit:20261004T220000\r\nDTEND;TZID=America\/Detroit:20261005T080000/);
  assert.match(ics, /DTSTART;TZID=America\/Detroit:20261001T153000\r\nDTEND;TZID=America\/Detroit:20261002T000000/);
  assert.match(ics, /DTSTART;VALUE=DATE:20261007\r\nDTEND;VALUE=DATE:20261008/);
  assert.match(ics, /LOCATION:2263 Fuller Ct\\, Ann Arbor\\, MI/);
  assert.match(ics, /TRIGGER:-PT60M/);
  assert.ok(ics.endsWith('END:VCALENDAR\r\n'));
  assert.ok(ics.split('\r\n').every((l) => l.length <= 75));
  assert.strictEqual(new Set(ev.map((e) => e.uid)).size, ev.length);
});

test('time off only exports when asked', () => {
  const d = C.parseWorkbook(XLSX, workbook(), C.DEFAULT_SHEET_PATTERN);
  const p = d.people.get('pat q');
  assert.strictEqual(C.buildPersonEvents(p, C.DEFAULT_SHIFTS, { from: '2026-10-01', to: '2026-10-31', includeOff: true }).length, 8);
});

test('a lone number is a start time that keeps the person\'s usual end', () => {
  const A = { normalCode: 'A', dow: 3 };
  const r = C.resolveCell('4', C.DEFAULT_SHIFTS, A);
  assert.deepStrictEqual([r.kind, r.startMin, r.endMin], ['shift', 16 * 60, 1440]); // 4 PM to 12 AM
  const m = C.resolveCell('10', C.DEFAULT_SHIFTS, { normalCode: 'M', dow: 3 });
  assert.deepStrictEqual([m.startMin, m.endMin], [22 * 60, 8 * 60 + 1440]); // 10 PM to 8 AM
  assert.strictEqual(C.resolveCell('4', C.DEFAULT_SHIFTS, {}).kind, 'unknown'); // no usual shift to borrow an end from
});

test('M starts at 10 PM on Saturday and Sunday', () => {
  const at = (dow) => C.resolveCell('M', C.DEFAULT_SHIFTS, { dow });
  assert.strictEqual(at(5).startMin, 20 * 60); // Friday
  assert.strictEqual(at(6).startMin, 22 * 60); // Saturday
  assert.strictEqual(at(0).startMin, 22 * 60); // Sunday
  assert.strictEqual(at(6).endMin, 8 * 60 + 1440);
  assert.strictEqual(at(1).startMin, 20 * 60); // Monday
});

test('name aliases merge two spellings into one person', () => {
  const wb = XLSX.utils.book_new();
  const days = Array.from({ length: 30 }, (_, i) => i + 1);
  const mk = (n) => [['x', ...days], ['', ...days.map(() => 'M')], [n, 'M', 'M', 'M', 'M']];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mk('Bernard')), 'June-26 A and M');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mk('Bernard K')), 'Sept 26 - A M');
  const d = C.parseWorkbook(XLSX, wb, C.DEFAULT_SHEET_PATTERN, C.parseAliases('Bernard = Bernard K'));
  assert.deepStrictEqual([...d.people.keys()], ['bernard k']);
  assert.strictEqual(d.people.get('bernard k').days.size, 8);
});
