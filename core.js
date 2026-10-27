/*
 * core.js: spreadsheet parsing, shift resolution and .ics generation.
 * No DOM access, so it runs in the browser (window.ShiftCore) and in Node (tests).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ShiftCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Sheets whose name contains "A M", "A and M" or "A & M" (e.g. "Oct 26 A M").
  const DEFAULT_SHEET_PATTERN = '\\bA\\s*(?:and|&)?\\s*M\\b';

  // Times for each code. The legend in the spreadsheet only says "M=8-8" and
  // "D = DAYS", so these are editable in the app's settings.
  const DEFAULT_SHIFTS = {
    A: { label: 'Shift A', start: '15:30', end: '00:00' }, // A=3:30-12
    M: { label: 'Shift M', start: '20:00', end: '08:00', weekend: { start: '22:00', end: '08:00' } }, // M=8-8 overnight; Sat/Sun starts are 10p-8a
    '2': { label: 'Shift 2', start: '14:00', end: '22:30' }, // 2=2-10:30
    E: { label: 'Shift E', start: '06:00', end: '14:00' }, // E=6-2
    D: { label: 'Days', start: '08:00', end: '16:30' }, // D=DAYS (hours assumed)
  };

  const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  const WEEKDAY = { SU: 0, M: 1, T: 2, W: 3, TH: 4, F: 5, SA: 6 };

  const pad = (n) => String(n).padStart(2, '0');
  const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
  const dateKey = (y, m, d) => `${y}-${pad(m + 1)}-${pad(d)}`;
  const parseKey = (k) => {
    const [y, m, d] = k.split('-').map(Number);
    return { y, m: m - 1, d };
  };
  const slug = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

  function cloneShifts(src) {
    const out = {};
    const base = src || DEFAULT_SHIFTS;
    for (const k of Object.keys(base)) {
      out[k] = { ...base[k] };
      if (base[k].weekend) out[k].weekend = { ...base[k].weekend };
    }
    return out;
  }

  /* ---------- cell and name cleanup ---------- */

  function normCell(v) {
    if (v == null || v instanceof Date) return '';
    const s = String(v).replace(/\u00a0/g, ' ').trim();
    return s.startsWith('=') ? '' : s; // stray formulas (broken external links) count as empty
  }
  const normName = (v) => normCell(v).replace(/\s+/g, ' ');
  const personKey = (name) => name.toLowerCase();

  /* ---------- time helpers ---------- */

  const toMin = (hhmm) => {
    const [h, m] = hhmm.split(':').map(Number);
    return h * 60 + (m || 0);
  };
  const toHHMM = (min) => `${pad(Math.floor(min / 60) % 24)}:${pad(min % 60)}`;

  // 15:30 -> "3:30 PM", 0 -> "12 AM"
  function fmtTime(min, short) {
    const t = ((min % 1440) + 1440) % 1440;
    const h24 = Math.floor(t / 60), m = t % 60;
    const h = h24 % 12 === 0 ? 12 : h24 % 12;
    const ap = h24 < 12 ? (short ? 'a' : ' AM') : short ? 'p' : ' PM';
    return `${h}${m ? ':' + pad(m) : ''}${ap}`;
  }
  const fmtRange = (s, e, short) => `${fmtTime(s, short)}\u2013${fmtTime(e, short)}`;

  /* ---------- parsing ---------- */

  // "Bernard = Bernard K" lines -> { 'bernard': 'Bernard K' }
  function parseAliases(text) {
    const map = {};
    for (const line of String(text || '').split(/\r?\n/)) {
      const m = line.split(/=|->/);
      if (m.length !== 2) continue;
      const from = m[0].trim().replace(/\s+/g, ' '), to = m[1].trim().replace(/\s+/g, ' ');
      if (from && to) map[from.toLowerCase()] = to;
    }
    return map;
  }

  // The shift code a person works most often, used to fill in a missing end time.
  function modalCode(person) {
    const counts = {};
    for (const { raw } of person.days.values()) {
      const u = raw.toUpperCase();
      if (Object.prototype.hasOwnProperty.call(DEFAULT_SHIFTS, u)) counts[u] = (counts[u] || 0) + 1;
    }
    let best = null;
    for (const k of Object.keys(counts)) if (best == null || counts[k] > counts[best]) best = k;
    return best;
  }

  // Context that resolveCell needs for one person on one day.
  function ctxFor(person, key) {
    const { y, m, d } = parseKey(key);
    return { dow: new Date(y, m, d).getDay(), normalCode: person.normalCode };
  }

  // Month and year come from the sheet name ("Sept 26 - A M"); cell A1 is the fallback.
  function parseSheetMonth(sheetName, a1) {
    const m = sheetName.toLowerCase().match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b[^0-9]*(\d{2}|\d{4})\b/);
    if (m) {
      const yr = Number(m[2]);
      return { year: yr < 100 ? 2000 + yr : yr, month: MONTHS[m[1]] };
    }
    if (a1 instanceof Date && !isNaN(a1)) return { year: a1.getFullYear(), month: a1.getMonth() };
    return null;
  }

  /**
   * Reads every sheet whose name matches `patternSrc`.
   * Layout: row 1 = day numbers (from column B), row 2 = weekday letters,
   * then one row per person (name in column A, one code per day).
   * Returns { people: Map(key -> {name, days: Map('YYYY-MM-DD' -> {raw, sheet})}), months, warnings }.
   */
  function parseWorkbook(XLSX, wb, patternSrc, aliases) {
    const alias = aliases || {};
    let re;
    try { re = new RegExp(patternSrc || DEFAULT_SHEET_PATTERN, 'i'); }
    catch (e) { re = new RegExp(DEFAULT_SHEET_PATTERN, 'i'); }

    const out = { people: new Map(), months: [], warnings: [] };

    for (const sheetName of wb.SheetNames) {
      if (!re.test(sheetName)) continue;
      const ws = wb.Sheets[sheetName];
      if (!ws || !ws['!ref']) continue;

      const range = XLSX.utils.decode_range(ws['!ref']);
      range.s.r = 0; range.s.c = 0; // sheets that start at A3 etc. still index from A1
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range, raw: true, defval: null, blankrows: true });
      if (!rows.length) continue;

      const ym = parseSheetMonth(sheetName, rows[0][0]);
      if (!ym) {
        out.warnings.push(`Skipped "${sheetName}": couldn't tell which month it is.`);
        continue;
      }
      const dim = daysInMonth(ym.year, ym.month);

      const dayCols = [];
      (rows[0] || []).forEach((v, c) => {
        if (c === 0 || v == null || v === '') return;
        const n = Number(String(v).trim());
        if (Number.isInteger(n) && n >= 1 && n <= dim) dayCols.push({ c, day: n });
      });
      if (!dayCols.length) {
        out.warnings.push(`Skipped "${sheetName}": no day numbers found in row 1.`);
        continue;
      }

      // Sanity check: the weekday letters in row 2 should agree with the month we inferred.
      let mismatches = 0;
      for (const { c, day } of dayCols) {
        const label = normCell((rows[1] || [])[c]).toUpperCase();
        if (label in WEEKDAY && WEEKDAY[label] !== new Date(ym.year, ym.month, day).getDay()) mismatches++;
      }
      if (mismatches > 2) {
        out.warnings.push(`"${sheetName}": weekday letters don't match the dates for ${mismatches} days. Check the sheet name.`);
      }

      const names = [];
      for (let r = 2; r < rows.length; r++) {
        const row = rows[r] || [];
        let name = normName(row[0]);
        if (!name || !/[A-Za-z]/.test(name)) continue;
        if (Object.prototype.hasOwnProperty.call(alias, name.toLowerCase())) name = alias[name.toLowerCase()];
        const cells = dayCols.map(({ c, day }) => ({ day, raw: normCell(row[c]) })).filter((x) => x.raw !== '');
        if (cells.length < 3) continue; // legend lines and spacer rows

        const key = personKey(name);
        if (!out.people.has(key)) out.people.set(key, { name, key, days: new Map(), months: new Set() });
        const person = out.people.get(key);
        person.months.add(dateKey(ym.year, ym.month, 1).slice(0, 7)); // they have a row this month, so blank days are days off
        for (const { day, raw } of cells) {
          person.days.set(dateKey(ym.year, ym.month, day), { raw, sheet: sheetName });
        }
        names.push(name);
      }
      out.months.push({ sheet: sheetName, year: ym.year, month: ym.month, people: names });
    }

    out.months.sort((a, b) => a.year - b.year || a.month - b.month);
    for (const p of out.people.values()) p.normalCode = modalCode(p);
    return out;
  }

  /* ---------- turning a cell code into a shift ---------- */

  const OFF = new Set(['0', '-', '--', 'O', 'OFF']);
  const RANGE_RE = /^(\d{1,2})(?::(\d{2}))?\s*-\s*(\d{1,2})(?::(\d{2}))?$/;

  // "3:30-8" -> 15:30 to 20:00. Start hours 1-5 are PM, 6-11 are AM; the end is the
  // first clock time after the start.
  function parseCustomRange(raw) {
    const m = raw.match(RANGE_RE);
    if (!m) return null;
    let h1 = Number(m[1]);
    const m1 = Number(m[2] || 0);
    let h2 = Number(m[3]);
    const m2 = Number(m[4] || 0);
    if (h1 > 12 || h2 > 24 || m1 > 59 || m2 > 59) return null;
    if (h1 >= 1 && h1 <= 5) h1 += 12;
    const start = h1 * 60 + m1;
    let end = null;
    for (const cand of [h2 * 60 + m2, (h2 + 12) * 60 + m2, (h2 + 24) * 60 + m2]) {
      if (cand > start) { end = cand; break; }
    }
    if (end == null) return null;
    return { startMin: start, endMin: end };
  }

  function timedShift(code, label, startMin, endMin) {
    return { kind: 'shift', code, label, startMin, endMin, crossesMidnight: endMin > 1440 || endMin === 1440 };
  }

  /**
   * Meaning of one cell. Returns null for days off / blank cells, otherwise
   * { kind: 'shift' | 'allday' | 'timeoff' | 'sick' | 'unknown', code, label, startMin?, endMin? }
   * where endMin is minutes after midnight of the start day (>1440 means it ends the next day).
   */
  function resolveCell(raw, shifts, ctx) {
    const cx = ctx || {};
    if (!raw) return null;
    const s = String(raw).trim();
    const up = s.toUpperCase();
    if (!s || OFF.has(up) || /^0+$/.test(s)) return null;
    if (up === 'L') return { kind: 'timeoff', code: 'L', label: 'Time off (requested)' };
    if (up === 'X') return { kind: 'sick', code: 'X', label: 'Called in sick' };
    if (up === 'S') return { kind: 'allday', code: 'S', label: 'Special shift' };

    const defs = shifts || DEFAULT_SHIFTS;
    if (Object.prototype.hasOwnProperty.call(defs, up)) {
      const base = defs[up];
      const d = (cx.dow === 0 || cx.dow === 6) && base.weekend ? base.weekend : base;
      let sm = toMin(d.start), em = toMin(d.end);
      if (em <= sm) em += 1440;
      return timedShift(up, base.label || `Shift ${up}`, sm, em);
    }
    const custom = parseCustomRange(s);
    if (custom) return { ...timedShift(s, 'Custom shift', custom.startMin, custom.endMin), custom: true };

    // A lone time such as "4" means "starts at 4, finishes at their usual end time".
    const bare = s.match(/^(\d{1,2})(?::(\d{2}))?$/);
    if (bare && cx.normalCode && Object.prototype.hasOwnProperty.call(defs, cx.normalCode)) {
      const h = Number(bare[1]), mi = Number(bare[2] || 0);
      if (h <= 23 && mi <= 59) {
        const usual = defs[cx.normalCode];
        const cands = h >= 1 && h <= 12 ? [(h % 12) * 60 + mi, (h % 12) * 60 + mi + 720] : [h * 60 + mi];
        const target = toMin(usual.start);
        const start = cands.reduce((a, b) => (Math.abs(b - target) < Math.abs(a - target) ? b : a));
        let end = toMin(usual.end);
        while (end <= start) end += 1440;
        return { ...timedShift(s, 'Custom start time', start, end), custom: true };
      }
    }
    return { kind: 'unknown', code: s, label: `Unrecognized code "${s}"` };
  }

  /* ---------- events and .ics ---------- */

  function addDaysKey(key, n) {
    const { y, m, d } = parseKey(key);
    const t = new Date(y, m, d + n);
    return dateKey(t.getFullYear(), t.getMonth(), t.getDate());
  }
  const compactDate = (key) => key.replace(/-/g, '');
  function localStamp(key, minutes) {
    const k = addDaysKey(key, Math.floor(minutes / 1440));
    const rem = minutes % 1440;
    return `${compactDate(k)}T${pad(Math.floor(rem / 60))}${pad(rem % 60)}00`;
  }

  /**
   * One person's schedule as event objects.
   * opts: { from, to (date keys, inclusive, optional), includeOff, prefix, midnightAs2359 }
   */
  function buildPersonEvents(person, shifts, opts) {
    const o = opts || {};
    const prefix = o.prefix || 'Work';
    const events = [];
    const keys = [...person.days.keys()].sort();
    for (const k of keys) {
      if (o.from && k < o.from) continue;
      if (o.to && k > o.to) continue;
      const cell = person.days.get(k);
      const r = resolveCell(cell.raw, shifts, ctxFor(person, k));
      if (!r) continue;
      const uidBase = `${slug(person.name)}-${compactDate(k)}-${slug(r.code)}@shift-calendar`;
      if (r.kind === 'shift') {
        let endMin = r.endMin;
        if (o.midnightAs2359 && endMin % 1440 === 0) endMin -= 1; // 12 AM -> 11:59 PM the same night
        events.push({
          uid: uidBase, dateKey: k, allDay: false,
          startMin: r.startMin, endMin,
          summary: prefix,
          description: `${r.label} (code ${r.code}): ${fmtRange(r.startMin, r.endMin, false).replace('\u2013', '-')}. From sheet "${cell.sheet}"`,
        });
      } else if (r.kind === 'allday') {
        events.push({
          uid: uidBase, dateKey: k, allDay: true,
          summary: prefix,
          description: `${r.label}. Time not listed in the schedule. From sheet "${cell.sheet}"`,
        });
      } else if (o.includeOff && (r.kind === 'timeoff' || r.kind === 'sick')) {
        events.push({
          uid: uidBase, dateKey: k, allDay: true, transparent: true,
          summary: r.kind === 'timeoff' ? 'Off (requested)' : 'Called in sick',
          description: `From sheet "${cell.sheet}"`,
        });
      }
    }
    return events;
  }

  const icsEscape = (s) =>
    String(s).replace(/\\/g, '\\\\').replace(/\r?\n/g, '\\n').replace(/;/g, '\\;').replace(/,/g, '\\,');

  function fold(line) {
    if (line.length <= 70) return line;
    const parts = [line.slice(0, 70)];
    for (let i = 70; i < line.length; i += 69) parts.push(' ' + line.slice(i, i + 69));
    return parts.join('\r\n');
  }

  const VTIMEZONE = [
    'BEGIN:VTIMEZONE',
    'TZID:America/Detroit',
    'BEGIN:DAYLIGHT',
    'TZOFFSETFROM:-0500',
    'TZOFFSETTO:-0400',
    'TZNAME:EDT',
    'DTSTART:19700308T020000',
    'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU',
    'END:DAYLIGHT',
    'BEGIN:STANDARD',
    'TZOFFSETFROM:-0400',
    'TZOFFSETTO:-0500',
    'TZNAME:EST',
    'DTSTART:19701101T020000',
    'RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU',
    'END:STANDARD',
    'END:VTIMEZONE',
  ];

  /**
   * opts: { calName, location, reminderMin (0 = none), now (Date, for tests) }
   * Times are written as America/Detroit wall-clock times with a VTIMEZONE block.
   */
  function buildICS(events, opts) {
    const o = opts || {};
    const now = o.now || new Date();
    const stamp = `${now.getUTCFullYear()}${pad(now.getUTCMonth() + 1)}${pad(now.getUTCDate())}T${pad(now.getUTCHours())}${pad(now.getUTCMinutes())}${pad(now.getUTCSeconds())}Z`;
    const L = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//shift-calendar//EN',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      `X-WR-CALNAME:${icsEscape(o.calName || 'Work schedule')}`,
      'X-WR-TIMEZONE:America/Detroit',
      ...VTIMEZONE,
    ];
    for (const e of events) {
      L.push('BEGIN:VEVENT', `UID:${e.uid}`, `DTSTAMP:${stamp}`);
      if (e.allDay) {
        L.push(`DTSTART;VALUE=DATE:${compactDate(e.dateKey)}`, `DTEND;VALUE=DATE:${compactDate(addDaysKey(e.dateKey, 1))}`);
      } else {
        L.push(`DTSTART;TZID=America/Detroit:${localStamp(e.dateKey, e.startMin)}`, `DTEND;TZID=America/Detroit:${localStamp(e.dateKey, e.endMin)}`);
      }
      L.push(`SUMMARY:${icsEscape(e.summary)}`);
      if (o.location) L.push(`LOCATION:${icsEscape(o.location)}`);
      if (e.description) L.push(`DESCRIPTION:${icsEscape(e.description)}`);
      L.push(e.transparent ? 'TRANSP:TRANSPARENT' : 'TRANSP:OPAQUE');
      if (o.reminderMin > 0 && !e.allDay) {
        L.push('BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:${icsEscape(e.summary)}`, `TRIGGER:-PT${o.reminderMin}M`, 'END:VALARM');
      }
      L.push('END:VEVENT');
    }
    L.push('END:VCALENDAR');
    return L.map(fold).join('\r\n') + '\r\n';
  }

  return {
    DEFAULT_SHEET_PATTERN, DEFAULT_SHIFTS, cloneShifts,
    parseWorkbook, parseSheetMonth, parseAliases, ctxFor, resolveCell, parseCustomRange,
    buildPersonEvents, buildICS,
    fmtTime, fmtRange, toMin, toHHMM, dateKey, parseKey, addDaysKey, daysInMonth, slug,
  };
});
