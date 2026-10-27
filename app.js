(function () {
  'use strict';

  const C = window.ShiftCore;
  const XLSX = window.XLSX;

  const PREF_KEY = 'shiftcal.prefs.v1';
  const DB_NAME = 'shift-calendar';
  const STORE = 'files';
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  /* ---------- small helpers ---------- */

  const $ = (id) => document.getElementById(id);

  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'style') el.style.cssText = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const kid of kids.flat()) {
      if (kid == null || kid === false) continue;
      el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
    }
    return el;
  }

  const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const ymd = (d) => C.dateKey(d.getFullYear(), d.getMonth(), d.getDate());
  const fromKey = (k) => { const p = C.parseKey(k); return new Date(p.y, p.m, p.d); };

  let toastTimer;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 3800);
  }
  function showBanner(msg) {
    const b = $('banner');
    b.textContent = msg || '';
    b.hidden = !msg;
  }

  /* ---------- preferences (localStorage) ---------- */

  function defaultPrefs() {
    return {
      person: 'bernard c',
      showCo: false,
      coworkers: null, // null = everyone else
      view: 'month',
      location: '2263 Fuller Ct, Ann Arbor, MI',
      prefix: 'Work',
      sheetPattern: C.DEFAULT_SHEET_PATTERN,
      shifts: C.cloneShifts(),
      aliases: 'Bernard = Bernard K',
      order: 'Afternoon:\nBernard C\nJanis\nJohn\nKristin\nMidnight:\nBernard K\nAbe\nTechs:\nAlex\nSarah\nYaw',
      showOff: true,
      hide: 'Eliana\nZade',
      offFor: 'Bernard C\nJanis\nJohn\nAlex\nSarah\nYaw',
      time24: false,
      exEnd: true,
      exRange: 'future',
      exReminder: '0',
      exOff: false,
    };
  }

  function loadPrefs() {
    const d = defaultPrefs();
    let p = {};
    try { p = JSON.parse(localStorage.getItem(PREF_KEY) || '{}') || {}; } catch (e) { p = {}; }
    const out = { ...d, ...p };
    if (out.order === 'Bernard C\nJanis\nKristin\nBernard K\nAbe\nAlex\nSarah\nYaw' || out.order === 'Bernard C\nJanis\nJohn\nKristin\nBernard K\nAbe\nAlex\nSarah\nYaw') out.order = d.order; // earlier defaults, now grouped
    out.shifts = C.cloneShifts();
    for (const code of Object.keys(out.shifts)) {
      const s = p.shifts && p.shifts[code];
      if (s && /^\d{2}:\d{2}$/.test(s.start) && /^\d{2}:\d{2}$/.test(s.end)) {
        out.shifts[code] = { ...out.shifts[code], start: s.start, end: s.end };
      }
      const w = s && s.weekend;
      if (out.shifts[code].weekend && w && /^\d{2}:\d{2}$/.test(w.start) && /^\d{2}:\d{2}$/.test(w.end)) {
        out.shifts[code].weekend = { start: w.start, end: w.end };
      }
    }
    return out;
  }
  function savePrefs() {
    try { localStorage.setItem(PREF_KEY, JSON.stringify(prefs)); } catch (e) { /* storage blocked */ }
  }

  /* ---------- remembered file (IndexedDB) ---------- */

  function openDB() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) return reject(new Error('IndexedDB unavailable'));
      const rq = indexedDB.open(DB_NAME, 1);
      rq.onupgradeneeded = () => rq.result.createObjectStore(STORE);
      rq.onsuccess = () => resolve(rq.result);
      rq.onerror = () => reject(rq.error);
    });
  }
  async function dbRun(mode, fn) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const rq = fn(tx.objectStore(STORE));
      tx.oncomplete = () => { db.close(); resolve(rq && rq.result); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }
  const saveFile = (rec) => dbRun('readwrite', (s) => s.put(rec, 'last'));
  const loadFile = () => dbRun('readonly', (s) => s.get('last'));
  const clearFile = () => dbRun('readwrite', (s) => s.delete('last'));

  /* ---------- state ---------- */

  const fmtT = (min, short) => C.fmtTime(min, short, prefs.time24);
  const fmtR = (s, e, short) => C.fmtRange(s, e, short, prefs.time24);

  let prefs = loadPrefs();
  let book = null;      // parsed SheetJS workbook, kept so settings changes can re-read it
  let data = null;      // result of ShiftCore.parseWorkbook
  let fileMeta = null;  // { name, savedAt, persisted }
  let cursor = startOfDay(new Date());
  let hues = new Map();

  // The order box: a line ending in ":" starts a group ("Afternoon:"), other lines are names.
  function parseOrder() {
    const groups = [];
    let cur = { heading: null, names: [] };
    groups.push(cur);
    for (const raw of String(prefs.order || '').split(/\r?\n/)) {
      const line = raw.trim().replace(/\s+/g, ' ');
      if (!line) continue;
      if (line.endsWith(':')) { cur = { heading: line.slice(0, -1).trim(), names: [] }; groups.push(cur); continue; }
      cur.names.push(line);
    }
    return groups.filter((g) => g.names.length);
  }

  // Splits people (already in display order) into the order box's groups; anyone unlisted goes in "Others".
  function layoutGroups(people) {
    const groups = parseOrder();
    const idx = new Map();
    groups.forEach((g, i) => g.names.forEach((n) => { if (!idx.has(n.toLowerCase())) idx.set(n.toLowerCase(), i); }));
    const anyHeading = groups.some((g) => g.heading);
    const buckets = new Map();
    for (const p of people) {
      const gi = idx.has(p.key) ? idx.get(p.key) : groups.length;
      if (!buckets.has(gi)) buckets.set(gi, []);
      buckets.get(gi).push(p);
    }
    return [...buckets.keys()].sort((x, y) => x - y).map((gi) => ({
      heading: gi < groups.length ? groups[gi].heading : anyHeading ? 'Others' : null,
      people: buckets.get(gi),
    }));
  }

  // Names in the order box that aren't on the schedule yet (e.g. a new hire) still get an empty row.
  function ensurePlaceholders() {
    for (const [k, p] of [...data.people]) if (p.placeholder) data.people.delete(k);
    for (const g of parseOrder()) {
      for (const name of g.names) {
        const key = name.toLowerCase();
        if (!data.people.has(key)) data.people.set(key, { name, key, days: new Map(), months: new Set(), normalCode: null, placeholder: true });
      }
    }
    chooseDefaultPerson();
  }

  // Display order first; anyone not listed follows alphabetically.
  function sortedPeople() {
    const rank = new Map();
    parseOrder().forEach((g) => g.names.forEach((n) => { if (!rank.has(n.toLowerCase())) rank.set(n.toLowerCase(), rank.size); }));
    const r = (p) => (rank.has(p.key) ? rank.get(p.key) : 1e6);
    const hidden = hiddenKeys();
    return [...data.people.values()].filter((p) => !hidden.has(p.key)).sort((x, y) => r(x) - r(y) || x.name.localeCompare(y.name));
  }

  const currentPerson = () => data.people.get(prefs.person);

  function coworkerKeys() {
    if (!prefs.showCo || !data) return [];
    const others = sortedPeople().map((p) => p.key).filter((k) => k !== prefs.person);
    return prefs.coworkers == null ? others : others.filter((k) => prefs.coworkers.includes(k));
  }

  const hiddenKeys = () => new Set(String(prefs.hide || '').split(/\r?\n/).map((n) => n.trim().replace(/\s+/g, ' ').toLowerCase()).filter(Boolean));

  function chooseDefaultPerson() {
    if (data.people.has(prefs.person) && !hiddenKeys().has(prefs.person)) return;
    const keys = [...data.people.keys()];
    prefs.person = keys.includes('bernard c') ? 'bernard c' : keys.find((k) => k.startsWith('bernard c')) || sortedPeople()[0].key;
  }

  function jumpToDataMonth() {
    if (!data.months.length) return;
    const t = startOfDay(new Date());
    const idx = (m) => m.year * 12 + m.month;
    const want = t.getFullYear() * 12 + t.getMonth();
    if (data.months.some((m) => idx(m) === want)) { cursor = t; return; }
    const best = data.months.reduce((a, b) => (Math.abs(idx(b) - want) < Math.abs(idx(a) - want) ? b : a));
    cursor = new Date(best.year, best.month, 1);
  }

  /* ---------- loading a file ---------- */

  async function ingest(buf, name, opts) {
    let wb;
    try {
      wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    } catch (e) {
      showBanner(`Couldn't open ${name}. Choose an Excel file (.xlsx).`);
      return false;
    }
    const parsed = C.parseWorkbook(XLSX, wb, prefs.sheetPattern, C.parseAliases(prefs.aliases));
    if (!parsed.people.size) {
      showBanner(`No schedule found in ${name}. Looked for sheets whose names match "${prefs.sheetPattern}". You can change that pattern in Settings.`);
      return false;
    }
    showBanner('');
    book = wb;
    data = parsed;
    fileMeta = { name, savedAt: opts.savedAt || Date.now(), persisted: !!opts.fromStore };
    if (opts.persist) {
      try {
        await saveFile({ name, buffer: buf, savedAt: fileMeta.savedAt });
        fileMeta.persisted = true;
      } catch (e) {
        toast("Loaded, but this browser wouldn't save the file for next time.");
      }
    }
    afterDataChange(true);
    return true;
  }

  function reparse() {
    if (!book) return;
    const parsed = C.parseWorkbook(XLSX, book, prefs.sheetPattern, C.parseAliases(prefs.aliases));
    if (!parsed.people.size) {
      showBanner(`No sheets match "${prefs.sheetPattern}". Showing the previous result.`);
      return;
    }
    showBanner('');
    data = parsed;
    afterDataChange(true);
  }

  function afterDataChange(resetCursor) {
    chooseDefaultPerson();
    hues = new Map(sortedPeople().map((p, i) => [p.key, Math.round((i * 137.508) % 360)]));
    if (resetCursor) jumpToDataMonth();
    $('empty').hidden = true;
    $('app').hidden = false;
    renderAll();
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      showBanner(`${file.name} isn't an Excel file. Choose a .xlsx file.`);
      return;
    }
    const buf = await file.arrayBuffer();
    const ok = await ingest(buf, file.name, { persist: true });
    if (ok) toast(`Loaded ${file.name}`);
  }

  /* ---------- rendering ---------- */

  function renderAll() {
    ensurePlaceholders();
    hues = new Map(sortedPeople().map((p, i) => [p.key, Math.round((i * 137.508) % 360)]));
    renderFileStatus();
    renderPeople();
    renderCoworkers();
    renderNotes();
    renderLegend();
    renderCalendar();
    $('toggleCo').setAttribute('aria-pressed', String(prefs.showCo));
    $('coPanel').hidden = !prefs.showCo;
    $('showOff').checked = prefs.showOff;
    document.querySelectorAll('#hoursSeg button').forEach((b) => b.setAttribute('aria-pressed', String((b.dataset.hours === '24') === !!prefs.time24)));
    document.querySelectorAll('#viewSeg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.view === prefs.view)));
  }

  function renderFileStatus() {
    const el = $('fileStatus');
    if (!fileMeta) { el.textContent = ''; return; }
    const first = data.months[0], last = data.months[data.months.length - 1];
    const span = first === last ? `${MONTH_NAMES[first.month]} ${first.year}` : `${MONTH_NAMES[first.month].slice(0, 3)} ${first.year} to ${MONTH_NAMES[last.month].slice(0, 3)} ${last.year}`;
    const saved = fileMeta.persisted ? ' Remembered on this device.' : '';
    el.textContent = `${fileMeta.name}: ${data.months.length} month${data.months.length === 1 ? '' : 's'} (${span}).${saved}`;
    el.title = el.textContent;
  }

  function renderPeople() {
    const sel = $('person');
    const opt = (p) => h('option', { value: p.key }, p.placeholder ? `${p.name} (not scheduled yet)` : p.name);
    sel.replaceChildren(...layoutGroups(sortedPeople()).map((g) => (g.heading ? h('optgroup', { label: g.heading }, g.people.map(opt)) : g.people.map(opt))).flat());
    sel.value = prefs.person;
  }

  function renderCoworkers() {
    const list = $('coList');
    const others = sortedPeople().filter((p) => p.key !== prefs.person);
    const on = new Set(coworkerKeys());
    const box = (p) => {
      const cb = h('input', { type: 'checkbox', value: p.key });
      cb.checked = prefs.coworkers == null || on.has(p.key);
      cb.addEventListener('change', () => {
        const checked = [...list.querySelectorAll('input:checked')].map((i) => i.value);
        prefs.coworkers = checked;
        savePrefs();
        renderCalendar();
      });
      return h('label', {}, cb, h('i', { class: 'swatch', style: `--h:${hues.get(p.key)}` }), p.name, p.placeholder ? h('span', { class: 'muted' }, ' (not on the schedule yet)') : null);
    };
    list.replaceChildren(...layoutGroups(others).map((g) => h('div', { class: 'co-group' },
      g.heading ? h('div', { class: 'co-group-name' }, g.heading) : null,
      h('div', { class: 'co-items' }, g.people.map(box)))));
  }

  function renderNotes() {
    const items = [...data.warnings];
    const unknown = new Map();
    const hidden = hiddenKeys();
    for (const p of data.people.values()) {
      if (hidden.has(p.key)) continue;
      for (const [k, cell] of p.days) {
        const r = C.resolveCell(cell.raw, prefs.shifts, C.ctxFor(p, k));
        if (r && r.kind === 'unknown') {
          const e = unknown.get(r.code) || { n: 0, who: new Set() };
          e.n++; e.who.add(p.name);
          unknown.set(r.code, e);
        }
      }
    }
    for (const [code, e] of unknown) {
      items.push(`Code "${code}" isn't in the legend and isn't a time (${e.n} day${e.n === 1 ? '' : 's'}: ${[...e.who].join(', ')}). It's shown with a dashed outline and left out of the .ics.`);
    }
    $('notes').hidden = !items.length;
    $('notesSummary').textContent = `${items.length} thing${items.length === 1 ? '' : 's'} to check in the spreadsheet`;
    $('notesList').replaceChildren(...items.map((t) => h('li', {}, t)));
  }

  function rangeText(s) {
    const st = C.toMin(s.start), en = C.toMin(s.end);
    return fmtR(st, en <= st ? en + 1440 : en, false);
  }

  function renderLegend() {
    const L = $('legend');
    const items = Object.entries(prefs.shifts).map(([code, s]) =>
      h('span', { 'data-code': code }, h('i'), `${s.label}: ${rangeText(s)}`));
    for (const [code, s] of Object.entries(prefs.shifts)) {
      if (s.weekend) items.splice(Object.keys(prefs.shifts).indexOf(code) + 1, 0, h('span', { 'data-code': code }, h('i'), `${s.label} on Sat/Sun: ${rangeText(s.weekend)}`));
    }
    items.push(h('span', { 'data-code': 'S' }, h('i'), 'Special shift (no time listed)'));
    if (prefs.showOff && prefs.showCo && parseNames(prefs.offFor).size) items.push(h('span', { class: 'muted' }, 'Dashed = day off'));
    L.replaceChildren(...items);
  }

  function cellFor(person, date) {
    const key = ymd(date);
    const c = person.days.get(key);
    return c ? C.resolveCell(c.raw, prefs.shifts, C.ctxFor(person, key)) : null;
  }

  function chipTitle(person, r) {
    if (r.kind === 'shift') return `${person.name}: ${r.label} (${fmtR(r.startMin, r.endMin, false)})`;
    return `${person.name}: ${r.label}`;
  }

  // mode: 'you' (filled), 'co' (outlined, with name, for the month grid), 'row' (rota board cell)
  function makeChip(person, r, mode, showName) {
    if (!r) return null;
    const isYou = mode === 'you';
    if (!isYou && (r.kind === 'timeoff' || r.kind === 'sick')) return null;
    const cls = ['chip', mode];
    if (showName) cls.push('named');
    if (r.kind === 'timeoff' || r.kind === 'sick') cls.push('status');
    if (r.kind === 'unknown') cls.push('unknown');
    const attrs = { class: cls.join(' '), title: chipTitle(person, r) };
    if (r.kind === 'shift' || r.kind === 'allday') attrs['data-code'] = r.kind === 'allday' ? 'S' : r.code;
    if (mode === 'co') attrs.style = `--h:${hues.get(person.key)}`;

    let text;
    if (r.kind === 'shift') {
      const range = fmtR(r.startMin, r.endMin, true);
      text = range;
    } else if (r.kind === 'allday') text = 'Special shift';
    else if (r.kind === 'timeoff') text = 'Off (requested)';
    else if (r.kind === 'sick') text = 'Called in sick';
    else text = `? ${r.code}`;

    const el = h('div', attrs);
    if (r.kind === 'shift') {
      const range = fmtR(r.startMin, r.endMin, true);
      if (mode === 'co') el.append(h('span', { class: 'sc' }), h('span', { class: 'nm' }, person.name), h('span', { class: 'time' }, range));
      else {
        if (showName) el.append(h('span', { class: 'nm' }, person.name));
        el.append(h('span', { class: 't-full' }, range), h('span', { class: 't-start' }, fmtT(r.startMin, true)));
      }
    } else {
      el.textContent = text;
    }
    return el;
  }

  function makeOff(person, mode) {
    const el = h('div', { class: `chip ${mode} off`, title: `${person.name}: off`, style: mode === 'co' ? `--h:${hues.get(person.key)}` : '' });
    if (mode === 'co') el.append(h('span', { class: 'nm' }, person.name), h('span', { class: 'time' }, 'off'));
    else el.textContent = 'off';
    return el;
  }

  // One person on one day: their shift, or a dashed "off" chip when they're on this month's schedule but not working.
  let offKeys = new Set();
  const parseNames = (t) => new Set(String(t || '').split(/\r?\n/).map((n) => n.trim().replace(/\s+/g, ' ').toLowerCase()).filter(Boolean));

  function chipFor(person, date, mode, showName) {
    const r = cellFor(person, date);
    if (mode === 'you') return makeChip(person, r, 'you', showName);
    if (r && (r.kind === 'shift' || r.kind === 'allday' || r.kind === 'unknown')) return makeChip(person, r, mode);
    if (prefs.showOff && offKeys.has(person.key) && person.months && person.months.has(ymd(date).slice(0, 7))) return makeOff(person, mode);
    return null;
  }

  function renderCalendar() {
    if (!data) return;
    const cal = $('calendar');
    const you = currentPerson();
    const cos = coworkerKeys().map((k) => data.people.get(k)).filter(Boolean);
    offKeys = parseNames(prefs.offFor);
    $('btnExport').disabled = !you;
    if (prefs.view === 'week') renderWeek(cal, you, cos);
    else renderMonth(cal, you, cos);
  }

  function renderMonth(cal, you, cos) {
    const y = cursor.getFullYear(), m = cursor.getMonth();
    $('calTitle').textContent = `${MONTH_NAMES[m]} ${y}`;
    const first = new Date(y, m, 1);
    const weeks = Math.ceil((first.getDay() + C.daysInMonth(y, m)) / 7);
    const start = addDays(first, -first.getDay());
    const todayKey = ymd(new Date());
    const cgroups = layoutGroups(cos);

    const rows = [];
    for (let w = 0; w < weeks; w++) {
      const cells = [];
      for (let i = 0; i < 7; i++) {
        const d = addDays(start, w * 7 + i);
        const cls = ['cell'];
        if (d.getMonth() !== m) cls.push('outside');
        if (i === 0 || i === 6) cls.push('weekend');
        if (ymd(d) === todayKey) cls.push('today');
        const label = `${DOW[d.getDay()]} ${MONTH_NAMES[d.getMonth()]} ${d.getDate()}`;
        cells.push(h('div', { class: cls.join(' '), 'aria-label': label },
          h('span', { class: 'num' }, d.getDate()),
          chipFor(you, d, 'you', true),
          cgroups.map((g) => {
            const chips = g.people.map((p) => chipFor(p, d, 'co')).filter(Boolean);
            return chips.length ? h('div', { class: 'grp' }, g.heading ? h('div', { class: 'grp-name' }, g.heading) : null, chips) : null;
          })));
      }
      rows.push(h('div', { class: 'week-row' }, cells));
    }
    cal.replaceChildren(h('div', { class: 'month' },
      h('div', { class: 'dow' }, DOW.map((n) => h('div', {}, n))),
      rows));
  }

  // Rota board: one row per person, one column per day.
  function renderWeek(cal, you, cos) {
    const ws = addDays(cursor, -cursor.getDay());
    const we = addDays(ws, 6);
    const mon = (d) => MONTH_NAMES[d.getMonth()].slice(0, 3);
    $('calTitle').textContent = ws.getMonth() === we.getMonth()
      ? `${mon(ws)} ${ws.getDate()} to ${we.getDate()}, ${we.getFullYear()}`
      : `${mon(ws)} ${ws.getDate()} to ${mon(we)} ${we.getDate()}, ${we.getFullYear()}`;

    const people = [you, ...cos];
    const todayKey = ymd(new Date());
    const days = Array.from({ length: 7 }, (_, i) => addDays(ws, i));
    const dayClass = (d) => 'day' + (d.getDay() === 0 || d.getDay() === 6 ? ' weekend' : '') + (ymd(d) === todayKey ? ' today' : '');

    const head = h('div', { class: 'board-row board-head' }, h('div', { class: 'who' }),
      days.map((d) => h('div', { class: dayClass(d) }, DOW[d.getDay()], h('b', {}, d.getDate()))));
    const personRow = (p, idx) =>
      h('div', { class: 'board-row' + (idx === 0 ? ' you-row' : ''), style: idx ? `--h:${hues.get(p.key)}` : '' },
        h('div', { class: 'who', title: p.name }, idx ? h('i', { class: 'swatch' }) : null, p.name),
        days.map((d) => h('div', { class: dayClass(d) }, chipFor(p, d, idx === 0 ? 'you' : 'row'))));
    const rows = [personRow(you, 0)];
    for (const g of layoutGroups(cos)) {
      if (g.heading) rows.push(h('div', { class: 'board-group' }, g.heading));
      g.people.forEach((p) => rows.push(personRow(p, 1)));
    }
    cal.replaceChildren(h('div', { class: 'board' }, h('div', { class: 'board-inner' }, head, rows)));
  }

  /* ---------- navigation ---------- */

  function step(dir) {
    if (prefs.view === 'week') cursor = addDays(cursor, 7 * dir);
    else cursor = new Date(cursor.getFullYear(), cursor.getMonth() + dir, 1);
    renderCalendar();
  }

  /* ---------- export ---------- */

  function exportRange() {
    const range = $('exRange').value;
    if (range === 'future') return { from: ymd(startOfDay(new Date())) };
    if (range === 'month') {
      const y = cursor.getFullYear(), m = cursor.getMonth();
      return { from: C.dateKey(y, m, 1), to: C.dateKey(y, m, C.daysInMonth(y, m)) };
    }
    return {};
  }

  function currentEvents() {
    return C.buildPersonEvents(currentPerson(), prefs.shifts, {
      ...exportRange(), time24: prefs.time24, includeOff: $('exOff').checked, prefix: $('exPrefix').value.trim() || 'Work', midnightAs2359: $('exEnd').checked,
    });
  }

  function updateExportCount() {
    const n = currentEvents().length;
    $('exCount').textContent = n ? `${n} event${n === 1 ? '' : 's'} will be added.` : 'No shifts in that range.';
    $('exGo').disabled = n === 0;
  }

  function openExport() {
    $('exportWho').textContent = currentPerson().name;
    $('exRange').value = prefs.exRange;
    $('exLocation').value = prefs.location;
    $('exPrefix').value = prefs.prefix;
    $('exReminder').value = prefs.exReminder;
    $('exOff').checked = prefs.exOff;
    $('exEnd').checked = prefs.exEnd;
    updateExportCount();
    $('exportDlg').showModal();
  }

  function downloadICS() {
    const p = currentPerson();
    prefs.exRange = $('exRange').value;
    prefs.location = $('exLocation').value.trim();
    prefs.prefix = $('exPrefix').value.trim() || 'Work';
    prefs.exReminder = $('exReminder').value;
    prefs.exOff = $('exOff').checked;
    prefs.exEnd = $('exEnd').checked;
    savePrefs();

    const events = currentEvents();
    const ics = C.buildICS(events, {
      calName: `${p.name} work schedule`,
      location: prefs.location,
      reminderMin: Number(prefs.exReminder),
    });
    const url = URL.createObjectURL(new Blob([ics], { type: 'text/calendar;charset=utf-8' }));
    const a = h('a', { href: url, download: `${C.slug(p.name)}-schedule.ics` });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    toast(`Downloaded ${events.length} event${events.length === 1 ? '' : 's'} for ${p.name}`);
  }

  /* ---------- settings ---------- */

  function openSettings() {
    const rows = [h('div', { class: 'head' }, 'Code'), h('div', { class: 'head' }, 'Starts'), h('div', { class: 'head' }, 'Ends')];
    for (const [code, s] of Object.entries(prefs.shifts)) {
      rows.push(
        h('b', { 'data-code': code }, h('i'), code === 'D' ? 'D (days)' : code),
        h('input', { type: 'time', value: s.start, 'data-code': code, 'data-edge': 'start', 'aria-label': `${code} start` }),
        h('input', { type: 'time', value: s.end, 'data-code': code, 'data-edge': 'end', 'aria-label': `${code} end` }));
      if (s.weekend) {
        rows.push(
          h('b', { 'data-code': code }, h('i'), `${code} Sat/Sun`),
          h('input', { type: 'time', value: s.weekend.start, 'data-code': code, 'data-edge': 'start', 'data-weekend': '1', 'aria-label': `${code} weekend start` }),
          h('input', { type: 'time', value: s.weekend.end, 'data-code': code, 'data-edge': 'end', 'data-weekend': '1', 'aria-label': `${code} weekend end` }));
      }
    }
    $('shiftTable').replaceChildren(...rows);
    $('stPattern').value = prefs.sheetPattern;
    $('stAliases').value = prefs.aliases;
    $('stOrder').value = prefs.order;
    $('stOff').value = prefs.offFor;
    $('stHide').value = prefs.hide;
    $('settingsDlg').showModal();
  }

  function saveSettings() {
    $('shiftTable').querySelectorAll('input[type=time]').forEach((i) => {
      if (!i.value) return;
      const target = i.dataset.weekend ? prefs.shifts[i.dataset.code].weekend : prefs.shifts[i.dataset.code];
      target[i.dataset.edge] = i.value;
    });
    const pattern = $('stPattern').value.trim() || C.DEFAULT_SHEET_PATTERN;
    const aliases = $('stAliases').value;
    prefs.order = $('stOrder').value;
    prefs.offFor = $('stOff').value;
    prefs.hide = $('stHide').value;
    const patternChanged = pattern !== prefs.sheetPattern || aliases !== prefs.aliases;
    prefs.aliases = aliases;
    try { new RegExp(pattern, 'i'); prefs.sheetPattern = pattern; } catch (e) { toast('That sheet pattern isn\'t valid, so it was left as it was.'); }
    savePrefs();
    if (patternChanged && data) reparse(); else if (data) renderAll();
  }

  /* ---------- wiring ---------- */

  function wire() {
    const pick = () => $('fileInput').click();
    $('btnUpload').addEventListener('click', pick);
    $('btnUpload2').addEventListener('click', pick);
    $('fileInput').addEventListener('change', (e) => { handleFile(e.target.files[0]); e.target.value = ''; });

    // drag and drop anywhere on the page
    let depth = 0;
    window.addEventListener('dragenter', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { depth++; $('dropVeil').hidden = false; } });
    window.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) $('dropVeil').hidden = true; });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => { e.preventDefault(); depth = 0; $('dropVeil').hidden = true; handleFile(e.dataTransfer.files[0]); });

    $('person').addEventListener('change', (e) => { prefs.person = e.target.value; savePrefs(); renderCoworkers(); renderCalendar(); });
    $('toggleCo').addEventListener('click', () => {
      prefs.showCo = !prefs.showCo; savePrefs();
      $('toggleCo').setAttribute('aria-pressed', String(prefs.showCo));
      $('coPanel').hidden = !prefs.showCo;
      renderCoworkers(); renderCalendar();
    });
    $('showOff').addEventListener('change', (e) => { prefs.showOff = e.target.checked; savePrefs(); renderLegend(); renderCalendar(); });
    $('coAll').addEventListener('click', () => { prefs.coworkers = null; savePrefs(); renderCoworkers(); renderCalendar(); });
    $('coNone').addEventListener('click', () => { prefs.coworkers = []; savePrefs(); renderCoworkers(); renderCalendar(); });

    document.querySelectorAll('#viewSeg button').forEach((b) => b.addEventListener('click', () => {
      prefs.view = b.dataset.view; savePrefs(); renderAll();
    }));
    document.querySelectorAll('#hoursSeg button').forEach((b) => b.addEventListener('click', () => {
      prefs.time24 = b.dataset.hours === '24'; savePrefs(); renderAll();
    }));
    $('prev').addEventListener('click', () => step(-1));
    $('next').addEventListener('click', () => step(1));
    $('today').addEventListener('click', () => { cursor = startOfDay(new Date()); renderCalendar(); });

    $('btnExport').addEventListener('click', openExport);
    ['exRange', 'exOff', 'exPrefix', 'exEnd'].forEach((id) => $(id).addEventListener('input', updateExportCount));
    $('exportForm').addEventListener('submit', (e) => {
      if (e.submitter && e.submitter.id === 'exGo') downloadICS();
    });

    $('btnSettings').addEventListener('click', openSettings);
    $('settingsForm').addEventListener('submit', (e) => {
      if (e.submitter && e.submitter.value === 'save') saveSettings();
    });
    $('stReset').addEventListener('click', () => {
      prefs.shifts = C.cloneShifts();
      $('shiftTable').querySelectorAll('input[type=time]').forEach((i) => {
        const t = i.dataset.weekend ? prefs.shifts[i.dataset.code].weekend : prefs.shifts[i.dataset.code];
        i.value = t[i.dataset.edge];
      });
    });
    $('stForget').addEventListener('click', async () => {
      try { await clearFile(); } catch (e) { /* nothing stored */ }
      if (fileMeta) fileMeta.persisted = false;
      if (data) renderFileStatus();
      toast('Saved file forgotten. It will ask for the spreadsheet next time.');
    });
    document.querySelectorAll('dialog [data-close]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close()));
    document.querySelectorAll('dialog').forEach((d) => d.addEventListener('click', (e) => { if (e.target === d) d.close(); }));
  }

  async function init() {
    wire();
    try {
      const rec = await loadFile();
      if (rec && rec.buffer) await ingest(rec.buffer, rec.name, { fromStore: true, savedAt: rec.savedAt });
    } catch (e) { /* no stored file, or storage unavailable: show the upload prompt */ }
  }

  init();
})();
