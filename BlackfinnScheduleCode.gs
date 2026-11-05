/**
 * Turns my weekly Blackfinn schedule (an Excel sheet that comes in through Gmail)
 * into Google Calendar events.
 *
 * Every 8 hours this checks Gmail (READ-ONLY) for the weekly schedule email
 * from my managers, reads the attached Excel sheet, and creates a "Blackfinn"
 * event for each of my shifts with the location, an estimated end time,
 * everyone working that day, and a KEY HOURLY flag when I'm the key hourly
 * (essentially managing).
 *
 * SETUP (one time): paste appsscript.json into the manifest file, then choose
 * "setup" in the toolbar and click Run.
 *
 * Functions:
 *   setup            - run once. Starts the timer that checks my email every 8 hours,
 *                      then does a first check right away. Safe to re-run: it replaces
 *                      the old timer instead of adding a second one.
 *   runNow           - redoes the newest schedule email: deletes that week's Blackfinn
 *                      events and recreates them. Use after changing the code, or to test.
 *                      If a change only needs to show up on next week's schedule, saving is enough.
 *   checkForSchedule - what the timer runs. Only handles emails it hasn't processed yet.
 *   buildShifts      - reads the spreadsheet. The other functions use it; it doesn't
 *                      work when run alone.
 */

// SETTINGS
const CONFIG = {
  MY_NAME: 'Tomas',
  EVENT_TITLE: 'Blackfinn',
  LOCATION: 'Blackfinn, 1620 I St NW, Washington, DC 20006',
  MANAGER_EMAILS: ['manager1@gmail.com', 'manager2@gmail.com'],
  CALENDAR_ID: 'primary',          // 'primary' = this account's own calendar
  CHECK_EVERY_HOURS: 8,

  // Shift-length rules (24h clock)
  CLOSE_HOUR: 24,                  // closers stay until midnight
  FIRST_PM_CUT_HOUR: 20,           // earliest PM bartender (usually the 3pm) leaves ~8pm
  AM_CUT_HOUR: 16,                 // AM shifts leave ~4pm when the PM crew arrives
  DEFAULT_SHIFT_HOURS: 6,          // fallback if nothing else applies
  MANAGER_SHIFT_HOURS: 6,          // a manager who comes in early is on duty ~6 hours
  MANAGER_CLOSER_FROM_HOUR: 15,    // a manager who comes in at 3pm or later closes
  TOUR_MANAGER_HOURS: 1.5,         // on tour days the manager stays ~1.5h after the tour starts

  FLIP_NIGHT_DAY: 'Thursday',      // PM crew closes late on flip night
  FLIP_CLOSE_HOUR: 25,             // 25 = 1am the next morning
  EVENT_COLOR: '8',                // 8 = Graphite for every shift
  REMINDER_MINUTES: 60,            // popup reminder before each shift (0 = none)
};

const TAG = 'blackfinnSchedule';
// Sections in order on the sheet (split by the repeated Monday…Sunday header rows).
// The last section holds Hosts (orange) then, after a blank row, Barbacks (gray).
const SECTION_NAMES = ['Managers', 'Bar', 'Servers', 'Hosts'];
const DONE_KEY = 'processedMessageIds';

/** Run once: creates the timer that checks your email automatically. */
function setup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkForSchedule')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkForSchedule').timeBased().everyHours(CONFIG.CHECK_EVERY_HOURS).create();
  checkForSchedule();
}

/** Re-processes the newest schedule email even if it was already added. */
function runNow() {
  const ids = findScheduleEmails_(1);
  if (!ids.length) { Logger.log('No schedule email found.'); return; }
  processMessage_(ids[0]);
}

/** Called by the timer: handles any schedule emails not yet added. */
function checkForSchedule() {
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty(DONE_KEY) || '[]');
  const ids = findScheduleEmails_(10).filter(id => !done.includes(id));
  ids.reverse().forEach(id => {                   // oldest first so revisions win
    if (processMessage_(id)) done.push(id);
  });
  props.setProperty(DONE_KEY, JSON.stringify(done.slice(-100)));
}

/** Returns IDs of recent schedule emails, newest first (read-only Gmail access). */
function findScheduleEmails_(max) {
  const from = CONFIG.MANAGER_EMAILS.map(e => 'from:' + e).join(' OR ');
  const q = `(${from}) subject:schedule has:attachment newer_than:21d`;
  const res = Gmail.Users.Messages.list('me', { q, maxResults: max });
  return (res.messages || []).map(m => m.id);
}

function processMessage_(id) {
  const msg = Gmail.Users.Messages.get('me', id);
  const subject = ((msg.payload.headers || []).find(h => /^subject$/i.test(h.name)) || {}).value || '';
  const part = findXlsxPart_(msg.payload);
  if (!part) return false;
  const att = Gmail.Users.Messages.Attachments.get('me', id, part.body.attachmentId);
  const bytes = typeof att.data === 'string' ? Utilities.base64DecodeWebSafe(att.data) : att.data;
  const blob = Utilities.newBlob(bytes, part.mimeType, part.filename);

  const grid = readExcel_(blob);
  const weekStart = grid.weekStart || dateFromText_(part.filename) || dateFromText_(subject);
  if (!weekStart) throw new Error('Could not find the week start date in ' + part.filename);
  const shifts = buildShifts(grid.display, weekStart, CONFIG.MY_NAME, grid.backgrounds);
  writeToCalendar_(weekStart, shifts);
  Logger.log(`Added ${shifts.length} shift(s) for week of ${weekStart.toDateString()}`);
  return true;
}

function findXlsxPart_(part) {
  if (!part) return null;
  if (part.filename && /\.xlsx?$/i.test(part.filename) && part.body && part.body.attachmentId) return part;
  for (const p of (part.parts || [])) { const f = findXlsxPart_(p); if (f) return f; }
  return null;
}

/** Converts the Excel file to a temporary Google Sheet, reads it, then trashes the copy. */
function readExcel_(blob) {
  const file = Drive.Files.create(
    { name: 'tmp-schedule-' + Date.now(), mimeType: 'application/vnd.google-apps.spreadsheet' },
    blob
  );
  try {
    const sheet = SpreadsheetApp.openById(file.id).getSheets()[0];
    const range = sheet.getDataRange();
    const values = range.getValues();
    const display = range.getDisplayValues();
    const backgrounds = range.getBackgrounds();
    let weekStart = null;
    for (const v of (values[1] || [])) { if (v instanceof Date) { weekStart = v; break; } }
    return { display, weekStart, backgrounds };
  } finally {
    Drive.Files.update({ trashed: true }, file.id);
  }
}

function dateFromText_(text) {
  const m = String(text || '').match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if (!m) return null;
  let y = +m[3]; if (y < 100) y += 2000;
  return new Date(y, +m[1] - 1, +m[2]);
}

function writeToCalendar_(weekStart, shifts) {
  let cal = CONFIG.CALENDAR_ID === 'primary'
    ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) {
    try { cal = CalendarApp.subscribeToCalendar(CONFIG.CALENDAR_ID); } catch (e) {}
  }
  if (!cal) throw new Error('Cannot open calendar ' + CONFIG.CALENDAR_ID +
    '. Share it with this account with "Make changes to events" permission.');
  // Remove events this script made earlier for the same week (handles revised schedules).
  const from = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate());
  const to = new Date(from.getTime()); to.setDate(to.getDate() + 8);
  cal.getEvents(from, to).filter(e => e.getTag(TAG)).forEach(e => e.deleteEvent());

  shifts.forEach(s => {
    const ev = cal.createEvent(s.title, s.start, s.end, { location: CONFIG.LOCATION, description: s.description });
    ev.setTag(TAG, '1');
    ev.setColor(CONFIG.EVENT_COLOR);
    if (CONFIG.REMINDER_MINUTES > 0) { ev.removeAllReminders(); ev.addPopupReminder(CONFIG.REMINDER_MINUTES); }
  });
}

// SCHEDULE PARSING (pure logic) 

/**
 * display: 2D array of cell text. weekStart: Date of Monday.
 * Returns [{start, end, title, description, isKey}]
 */
function buildShifts(display, weekStart, myName, backgrounds) {
  const days = findDayColumns_(display);
  const people = readPeople_(display, days.length, backgrounds);
  const events = readEvents_(display, days);
  const me = people.find(p => p.name.toLowerCase() === myName.toLowerCase());
  if (!me) throw new Error(myName + ' not found on the schedule');

  const shifts = [];
  days.forEach((day, d) => {
    ['AM', 'PM'].forEach(slot => {
      const cell = me.cells[d][slot];
      if (!cell || cell.start == null) return;

      const date = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + d);
      const start = atTime_(date, cell.start);
      const flipNight = day.name === CONFIG.FLIP_NIGHT_DAY ||
                        (events[d] || []).some(e => /flip/i.test(e));
      const est = estimateEnd_(cell, slot, d, people, me, flipNight);
      const end = atTime_(date, est.hour);

      const closeHour = flipNight ? CONFIG.FLIP_CLOSE_HOUR : CONFIG.CLOSE_HOUR;
      const timeline = inChargeTimeline_(d, people, closeHour, events[d]);
      const mine = timeline.filter(t => t.people.includes(me) && t.to > cell.start && t.from < est.hour);
      const isKey = mine.length > 0;

      const title = CONFIG.EVENT_TITLE + (isKey ? ' ⭐ KEY HOURLY' : '');
      const description = describe_(day.name, slot, d, people, me, cell, est, timeline, mine, events[d]);
      shifts.push({ start, end, title, description, isKey, dayName: day.name, slot });
    });
  });
  return shifts;
}

/**
 * Who is in charge through the day:
 *  - while a manager (Manager1/Manager2 rows) is on, there is no key hourly
 *  - otherwise every key hourly (ky/key/clky/KyMgr) still in the building
 *    is in charge together (AM keys leave ~4pm, PM keys close)
 * Returns [{from, to, people, isManager}]
 */
function inChargeTimeline_(d, people, closeHour, dayEvents) {
  const C = CONFIG;
  const mgrs = [], keys = [];
  // Tours (e.g. "45@515 Tour") need the manager until ~1.5h after they start.
  let tourEnd = 0;
  (dayEvents || []).filter(e => /tour/i.test(e)).forEach(e => {
    const m = e.replace(/\s/g, '').match(/@(\d{1,4})(am|pm|a|p)?/i);
    if (m) tourEnd = Math.max(tourEnd, toHour_(m[1], m[2] && m[2].toLowerCase(), 'PM') + C.TOUR_MANAGER_HOURS);
  });
  people.forEach(p => {
    const cs = ['AM', 'PM'].map(s => ({ s, c: p.cells[d][s] })).filter(x => x.c);
    const starts = cs.filter(x => x.c.start != null).map(x => x.c.start);
    if (!starts.length) return;
    if (p.role === 'Managers') {
      cs.filter(x => x.c.start != null).forEach(x => mgrs.push({ person: p, from: x.c.start, cell: x.c }));
      return;
    }
    const keyCell = cs.find(x => x.c.key);
    if (!keyCell) return;
    const from = keyCell.c.start != null ? keyCell.c.start : Math.min(...starts);
    // Use their last cell with an actual time ("volume"/"KyMgr" alone don't extend to close).
    const timed = cs.filter(x => x.c.start != null);
    const lastCell = (keyCell.c.start == null && keyCell.s === 'PM') ? keyCell : timed[timed.length - 1];
    let to;
    if (lastCell.c.endExplicit != null) to = lastCell.c.endExplicit;
    else if (lastCell.s === 'PM') to = closeHour;
    else to = Math.max(C.AM_CUT_HOUR, from + 4);
    keys.push({ person: p, from, to, pmKey: keyCell.s === 'PM' && keyCell.c.start != null });
  });

  // Manager hours:
  //  1) comes in 3pm+ and no PM key hourly scheduled -> closes (tours don't matter)
  //  2) comes in 3pm+ with a PM key hourly -> leaves after the tour (~1.5h after it starts),
  //    or after the usual shift length if there's no tour
  //  3) comes in earlier -> ~6h, or until the tour is over if that's later
  const nightKey = keys.some(k => k.pmKey);
  mgrs.forEach(m => {
    let to;
    if (m.cell.endExplicit != null) to = m.cell.endExplicit;
    else if (m.from >= C.MANAGER_CLOSER_FROM_HOUR) to = !nightKey ? closeHour
      : (tourEnd > m.from ? tourEnd : m.from + C.MANAGER_SHIFT_HOURS);
    else to = Math.max(m.from + C.MANAGER_SHIFT_HOURS, tourEnd);
    m.to = Math.min(to, closeHour);
  });

  const points = [...new Set([...mgrs, ...keys].flatMap(x => [x.from, x.to]))].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < points.length - 1; i++) {
    const t = (points[i] + points[i + 1]) / 2;
    const mgr = mgrs.filter(m => m.from <= t && t < m.to).sort((a, b) => a.from - b.from)[0];
    const onKeys = keys.filter(k => k.from <= t && t < k.to).sort((a, b) => a.from - b.from);
    const who = mgr ? [mgr.person] : onKeys.map(k => k.person);   // several keys can share a shift
    if (!who.length) continue;
    const prev = out[out.length - 1];
    const same = prev && prev.to === points[i] && prev.people.length === who.length &&
                 prev.people.every((p, j) => p === who[j]);
    if (same) prev.to = points[i + 1];
    else out.push({ from: points[i], to: points[i + 1], people: who, isManager: !!mgr });
  }
  return out;
}

function findDayColumns_(display) {
  const names = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  const header = display[0].map(v => String(v).trim().toLowerCase());
  return names.map(n => {
    const col = header.indexOf(n);
    if (col < 0) throw new Error('Day header not found: ' + n);
    return { name: n[0].toUpperCase() + n.slice(1), am: col, pm: col + 1 };
  });
}

function readPeople_(display, nDays, backgrounds) {
  const days = findDayColumns_(display);
  const people = [];
  let section = -1, pastBlank = false;
  for (let r = 0; r < display.length; r++) {
    const row = display[r];
    const a = String(row[0] || '').trim();
    if (String(row[days[0].am] || '').trim().toLowerCase() === 'monday') { section++; pastBlank = false; continue; }
    if (/^events$/i.test(a)) break;
    if (r < 3) continue;
    if (!a) {
      if (section >= 3 && people.some(p => p.section === section) &&
          row.every(v => !String(v || '').trim())) pastBlank = true;
      continue;
    }
    const cells = days.map(day => ({
      AM: parseCell_(row[day.am], 'AM'),
      PM: parseCell_(row[day.pm], 'PM'),
    }));
    let role = SECTION_NAMES[Math.min(section, 3)] || 'Hosts';
    if (section >= 3 && (pastBlank || isGray_(backgrounds && backgrounds[r] && backgrounds[r][days[0].am]))) role = 'Barbacks';
    people.push({ name: a, section, role, cells });
  }
  return people;
}

function isGray_(hex) {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(hex || ''));
  if (!m) return false;
  const [r, g, b] = [m[1], m[2], m[3]].map(x => parseInt(x, 16));
  return Math.abs(r - g) < 8 && Math.abs(g - b) < 8 && r < 0xf0 && r > 0x80;
}

/**
 * What someone is doing on a given day: their section's role, unless that
 * day's cell says otherwise ("ser" = serving, "br" = bartending).
 */
function dayRole_(p, d) {
  if (p.role === 'Managers') return 'Managers';
  const cs = [p.cells[d].AM, p.cells[d].PM].filter(Boolean);
  if (cs.some(c => c.bar)) return 'Bar';
  if (p.role === 'Bar') return 'Bar';
  if (cs.some(c => c.server)) return 'Servers';
  return p.role;
}

function readEvents_(display, days) {
  const out = days.map(() => []);
  let start = display.findIndex(row => /^events$/i.test(String(row[0] || '').trim()));
  if (start < 0) return out;
  for (let r = start; r < display.length; r++) {
    days.forEach((day, d) => {
      const txt = [display[r][day.am], display[r][day.pm]].map(v => String(v || '').trim()).filter(Boolean).join(' ');
      if (txt) out[d].push(txt.replace(/\s+/g, ' '));
    });
  }
  return out;
}

/** Parses things like "4pmclky", "10amky", "530pm", "12-8pm", "2p", "volume", "o", "KyMgr". */
function parseCell_(raw, slot) {
  const text = String(raw == null ? '' : raw).trim();
  if (!text) return null;
  const t = text.toLowerCase().replace(/\s+/g, '');
  const res = { raw: text, start: null, endExplicit: null, code: '', key: false, closer: false,
                server: false, bar: false, training: false, volume: false, off: false, mgr: false };

  let m = t.match(/^(\d{1,4})(am|pm|a|p)?-(\d{1,4})(am|pm|a|p)?(.*)$/);
  if (m) {
    res.start = toHour_(m[1], m[2], slot);
    res.endExplicit = toHour_(m[3], m[4] || 'pm', 'PM');
    if (res.endExplicit <= res.start) res.endExplicit += 12;
    res.code = m[5];
  } else if ((m = t.match(/^(\d{1,4})(am|pm|a|p)?(.*)$/))) {
    res.start = toHour_(m[1], m[2], slot);
    res.code = m[3];
  } else {
    res.code = t;
  }
  const c = res.code;
  res.key = /k(e)?y/.test(c);
  res.closer = /cl/.test(c);
  res.server = /ser/.test(c);
  res.bar = /br/.test(c);
  res.training = /tr$|trn|train/.test(c);
  res.volume = /vol/.test(c);
  res.off = c === 'o' || c === 'off' || c === 'out';
  res.mgr = /mgr/.test(c);
  if (res.off) return null;
  return res;
}

function toHour_(digits, mer, slot) {
  let h, min = 0;
  if (digits.length <= 2) h = +digits;
  else if (digits.length === 3) { h = +digits[0]; min = +digits.slice(1); }
  else { h = +digits.slice(0, 2); min = +digits.slice(2); }
  let pm;
  if (mer) pm = mer[0] === 'p';
  else pm = !(slot === 'AM' && h >= 6 && h <= 11);
  if (pm && h < 12) h += 12;
  if (!pm && h === 12) h = 0;
  return h + min / 60;
}

/**
 * End-time rules:
 *  - explicit range (12-8pm) -> use it
 *  - AM shift -> leaves ~4pm (note if staying on volume)
 *  - closer (cl / clky) or PM key hourly -> midnight
 *  - otherwise PM bartenders leave in the order they came in: the first one
 *    out ~8pm, the last one in closes at midnight, others spread in between.
 */
function estimateEnd_(cell, slot, d, people, me, flipNight) {
  const C = Object.assign({}, CONFIG);
  if (flipNight) C.CLOSE_HOUR = CONFIG.FLIP_CLOSE_HOUR;
  if (cell.endExplicit != null) return { hour: cell.endExplicit, note: 'as scheduled' };
  if (flipNight && slot === 'PM') return { hour: C.CLOSE_HOUR, note: 'flip night – PM crew closes' };

  if (slot === 'AM') {
    const pmCell = me.cells[d].PM;
    if (pmCell && pmCell.volume) return { hour: C.AM_CUT_HOUR, note: 'may stay later depending on volume' };
    return { hour: Math.max(C.AM_CUT_HOUR, cell.start + 4), note: 'estimated' };
  }

  if (cell.closer) return { hour: C.CLOSE_HOUR, note: 'closing' };
  if (cell.key) return { hour: C.CLOSE_HOUR, note: 'key hourly, closing' };
  if (dayRole_(me, d) !== 'Bar') return { hour: Math.min(C.CLOSE_HOUR, cell.start + C.DEFAULT_SHIFT_HOURS), note: 'estimated' };

  const bar = people
    .filter(p => dayRole_(p, d) === 'Bar')
    .map(p => p.cells[d].PM)
    .filter(c => c && c.start != null && c.start >= 12);
  const closerStarts = bar.filter(c => c.closer || c.key).map(c => c.start);
  const latest = Math.max(...bar.map(c => c.start));
  // If nobody is marked closer, whoever comes in last closes.
  if (!closerStarts.length && cell.start === latest) return { hour: C.CLOSE_HOUR, note: 'last in, closing' };

  const nonCloserStarts = [...new Set(bar
    .filter(c => !(c.closer || c.key))
    .filter(c => closerStarts.length || c.start !== latest)
    .map(c => c.start))].sort((a, b) => a - b);
  const idx = nonCloserStarts.indexOf(cell.start);
  const n = nonCloserStarts.length;
  const span = C.CLOSE_HOUR - C.FIRST_PM_CUT_HOUR;
  let hour = C.FIRST_PM_CUT_HOUR + (idx < 0 ? 0 : idx * span / Math.max(n, 1));
  hour = Math.round(hour * 2) / 2;                                 // nearest half hour
  hour = Math.max(hour, cell.start + 3);
  return { hour: Math.min(hour, C.CLOSE_HOUR), note: `estimated – #${idx + 1} of ${n} out before closers` };
}

function atTime_(date, hour) {
  const d = new Date(date.getTime());
  d.setHours(Math.floor(hour), Math.round((hour % 1) * 60), 0, 0);
  return d;
}

function fmt_(hour) {
  if (hour >= 24) hour -= 24;
  const h = Math.floor(hour), m = Math.round((hour % 1) * 60);
  const h12 = ((h + 11) % 12) + 1;
  return h12 + (m ? ':' + String(m).padStart(2, '0') : '') + (h < 12 ? 'am' : 'pm');
}

function label_(c) {
  const tags = [];
  if (c.key && c.mgr) tags.push('key manager');
  else if (c.key && c.closer) tags.push('closer, KEY');
  else if (c.key) tags.push('KEY');
  else if (c.closer) tags.push('closer');
  if (c.bar && c.training) tags.push('bartender training');
  if (c.volume) tags.push('stays on volume');
  const time = c.start != null ? fmt_(c.start) + (c.endExplicit != null ? '–' + fmt_(c.endExplicit) : '') : '';
  if (!time) return tags.length ? tags.join(', ') : c.raw;
  return time + (tags.length ? ` (${tags.join(', ')})` : '');
}

function names_(people, me) {
  return people.map(p => p === me ? 'You' : p.name).join(' & ');
}

function describe_(dayName, slot, d, people, me, cell, est, timeline, mine, dayEvents) {
  const lines = [];
  if (mine.length) {
    const from = Math.max(mine[0].from, cell.start);
    lines.push(`⭐ YOU ARE THE KEY HOURLY from ${fmt_(from)}` + (from > cell.start ? ' (manager on duty before that)' : ''));
  } else {
    const during = timeline.filter(t => t.to > cell.start && t.from < est.hour);
    lines.push('In charge during your shift: ' + (during.length
      ? during.map(t => `${names_(t.people, me)} (${t.isManager ? 'manager' : 'key hourly'})`).join(' → ')
      : 'no manager or key marked'));
  }
  lines.push(`Your shift: ${fmt_(cell.start)} – ${fmt_(est.hour)} (${est.note})`);
  lines.push('');
  if (timeline.length) {
    lines.push("WHO'S IN CHARGE TODAY");
    timeline.forEach(t => lines.push(`• ${fmt_(t.from)}–${fmt_(t.to)}: ${names_(t.people, me)} (${t.isManager ? 'manager' : 'key hourly'})`));
    lines.push('');
  }

  const groups = ['Managers', 'Bar', 'Servers', 'Hosts', 'Barbacks']
    .map(g => [g, p => dayRole_(p, d) === g]);
  groups.forEach(([title, test]) => {
    const rows = [];
    people.filter(test).forEach(p => {
      const cs = ['AM', 'PM'].map(s => p.cells[d][s]).filter(Boolean);
      if (!cs.length) return;
      const starts = cs.filter(c => c.start != null).map(c => c.start);
      rows.push({ sort: starts.length ? Math.min(...starts) : 23.9,
                  text: `• ${p.name}${p === me ? ' (you)' : ''} – ${cs.map(label_).join(', then ')}` });
    });
    if (!rows.length) return;
    rows.sort((a, b) => a.sort - b.sort);
    lines.push(title.toUpperCase());
    rows.forEach(r => lines.push(r.text));
    lines.push('');
  });

  // HTML for google cal events
  const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  while (lines.length && !lines[lines.length - 1]) lines.pop();
  const html = lines.map(esc);
  if (dayEvents && dayEvents.length) {
    html.push('', '<b>EVENTS / PARTIES</b>');
    dayEvents.forEach(e => html.push('<b>• ' + esc(e) + '</b>'));
  }
  return html.join('<br>');
}

// Allow the parsing logic to be tested outside Apps Script.
if (typeof module !== 'undefined') module.exports = { buildShifts, parseCell_, CONFIG };