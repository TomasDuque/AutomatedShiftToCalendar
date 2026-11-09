/**
 * Pulls my weekly work schedule (.xlsx) out of Gmail and adds my shifts to Google Calendar,
 * with who I'm working with, estimated end times, and whether I'm the key hourly.
 * Gmail access is read-only.
 *
 * setup()  - run once; starts the timer (every CHECK_EVERY_HOURS) and checks right away.
 * runNow() - rebuilds the newest week, e.g. after changing a setting.
 */

// ---- Settings ----
const CONFIG = {
  MY_NAME: 'Your Name',            // as written on the schedule
  EVENT_TITLE: 'Blackfinn',
  LOCATION: 'Restaurant Name, 123 Main St, City, ST 00000',
  MANAGER_EMAILS: ['manager1@example.com', 'manager2@example.com'],
  CALENDAR_ID: 'primary',
  CHECK_EVERY_HOURS: 8,

  // Hours use a 24h clock (25 = 1am)
  CLOSE_HOUR: 24,
  FIRST_PM_CUT_HOUR: 20,           // first PM bartender cut
  AM_CUT_HOUR: 16,                 // AM shift cut
  DEFAULT_SHIFT_HOURS: 6,
  MANAGER_SHIFT_HOURS: 6,          // early manager shift length
  MANAGER_CLOSER_FROM_HOUR: 15,    // manager in at 3pm+ closes
  TOUR_MANAGER_HOURS: 1.5,         // manager stays this long after a tour starts

  FLIP_NIGHT_DAY: 'Thursday',      // PM crew closes late
  FLIP_CLOSE_HOUR: 25,
  EVENT_COLOR: '8',                // Graphite
  REMINDER_MINUTES: 60,            // 0 = no reminder
};

const TAG = 'workSchedule';
const DONE_KEY = 'processedMessageIds';

// Run once: sets up the trigger and checks right away.
function setup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkForSchedule')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkForSchedule').timeBased().everyHours(CONFIG.CHECK_EVERY_HOURS).create();
  checkForSchedule();
}

// Rebuilds the newest week, even if it was already added.
function runNow() {
  const ids = findScheduleEmails_(5);
  for (const id of ids) { if (processMessage_(id)) return; }   // newest real schedule
  Logger.log('No schedule email found.');
}

// Trigger: handle any emails I haven't checked yet.
function checkForSchedule() {
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty(DONE_KEY) || '[]');
  const ids = findScheduleEmails_(10).filter(id => !done.includes(id));
  ids.reverse().forEach(id => {                   // oldest first so a revised schedule wins
    processMessage_(id);
    done.push(id);                                // don't re-check other spreadsheets every run
  });
  props.setProperty(DONE_KEY, JSON.stringify(done.slice(-100)));
}

// Recent emails from my managers with an Excel file, newest first.
// I don't search by subject anymore since it's sometimes misspelled.
function findScheduleEmails_(max) {
  const from = CONFIG.MANAGER_EMAILS.map(e => 'from:' + e).join(' OR ');
  const q = `(${from}) has:attachment (filename:xlsx OR filename:xls) newer_than:21d`;
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
  let shifts;
  try {
    if (!weekStart) throw new Error('no week start date found');
    shifts = buildShifts(grid.display, weekStart, CONFIG.MY_NAME, grid.backgrounds);
  } catch (e) {
    // Not a schedule (or I'm not on it), so skip it.
    Logger.log(`Skipped "${subject}" (${part.filename}): ${e.message}`);
    return false;
  }
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

// Convert the .xlsx to a temp Google Sheet, read it, then trash it.
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
  // Clear my earlier events for this week so revisions don't duplicate.
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

// ---- Parsing / rules (no Google APIs, so I can test it in Node) ----

// Turns the sheet into my shifts: [{start, end, title, description, isKey}]
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

// Who's in charge through the day. A manager on duty overrides any key hourly;
// otherwise every key hourly still in the building shares it.
// Returns [{from, to, people, isManager}]
function inChargeTimeline_(d, people, closeHour, dayEvents) {
  const C = CONFIG;
  const mgrs = [], keys = [];
  // Tours (e.g. "45@515 Tour") keep the manager around.
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
    // "volume"/"KyMgr" have no time, so they don't keep someone until close.
    const timed = cs.filter(x => x.c.start != null);
    const lastCell = (keyCell.c.start == null && keyCell.s === 'PM') ? keyCell : timed[timed.length - 1];
    let to;
    if (lastCell.c.endExplicit != null) to = lastCell.c.endExplicit;
    else if (lastCell.s === 'PM') to = closeHour;
    else to = Math.max(C.AM_CUT_HOUR, from + 4);
    keys.push({ person: p, from, to, pmKey: keyCell.s === 'PM' && keyCell.c.start != null });
  });

  // Manager in at 3pm+: closes, unless a night key is on (then leaves after the tour).
  // Earlier manager: ~6h, or until the tour is over.
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
    const who = mgr ? [mgr.person] : onKeys.map(k => k.person);
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
  // Groups go top to bottom. A day-header row or a blank row starts the next one
  // (the schedule uses both, depending on the week).
  const GROUPS = ['Managers', 'Bar', 'Servers', 'Hosts', 'Barbacks'];
  const days = findDayColumns_(display);
  const people = [];
  let group = 0, inGroup = 0, pendingBreak = false;
  for (let r = 0; r < display.length; r++) {
    const row = display[r];
    const a = String(row[0] || '').trim();
    if (/^events$/i.test(a)) break;
    if (r < 3) continue;
    const isHeader = String(row[days[0].am] || '').trim().toLowerCase() === 'monday';
    const isBlank = row.every(v => !String(v || '').trim());
    if (isHeader || isBlank) { if (inGroup) pendingBreak = true; continue; }
    if (!a) continue;                                   // times with no name
    if (pendingBreak) { group = Math.min(group + 1, GROUPS.length - 1); inGroup = 0; pendingBreak = false; }
    inGroup++;
    const cells = days.map(day => {
      const am = parseCell_(row[day.am], 'AM'), pm = parseCell_(row[day.pm], 'PM');
      // "3pm | server" -> the role in the next cell goes with that shift
      if (am && am.start != null && pm && pm.start == null && (pm.server || pm.bar)) {
        am.server = am.server || pm.server; am.bar = am.bar || pm.bar;
        return { AM: am, PM: null };
      }
      return { AM: am, PM: pm };
    });
    let role = GROUPS[group];
    if (role === 'Hosts' && isGray_(backgrounds && backgrounds[r] && backgrounds[r][days[0].am])) role = 'Barbacks';
    people.push({ name: a, section: group, role, cells });
  }
  return people;
}

function isGray_(hex) {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(hex || ''));
  if (!m) return false;
  const [r, g, b] = [m[1], m[2], m[3]].map(x => parseInt(x, 16));
  return Math.abs(r - g) < 8 && Math.abs(g - b) < 8 && r < 0xf0 && r > 0x80;
}

// Role for one shift: "br" = bar, "ser" = serving, otherwise their group.
function cellRole_(p, c) {
  if (p.role === 'Managers') return 'Managers';
  if (c.bar) return 'Bar';
  if (c.server) return 'Servers';
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

// Cells look like "4pmclky", "10amky", "530pm", "12-8pm", "volume", "o", "KyMgr".
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

// End times: AM out ~4pm, closers/keys stay to close, and the rest of the PM bar
// gets cut in the order they came in (first in, first out).
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
  if (cellRole_(me, cell) !== 'Bar') return { hour: Math.min(C.CLOSE_HOUR, cell.start + C.DEFAULT_SHIFT_HOURS), note: 'estimated' };

  const bar = people
    .filter(p => { const c = p.cells[d].PM; return c && c.start != null && c.start >= 12 && cellRole_(p, c) === 'Bar'; })
    .map(p => p.cells[d].PM);
  const closerStarts = bar.filter(c => c.closer || c.key).map(c => c.start);
  const latest = Math.max(...bar.map(c => c.start));
  // No closer marked -> last one in closes.
  if (!closerStarts.length && cell.start === latest) return { hour: C.CLOSE_HOUR, note: 'last in, closing' };

  const nonCloserStarts = [...new Set(bar
    .filter(c => !(c.closer || c.key))
    .filter(c => closerStarts.length || c.start !== latest)
    .map(c => c.start))].sort((a, b) => a - b);
  const idx = nonCloserStarts.indexOf(cell.start);
  const n = nonCloserStarts.length;
  const span = C.CLOSE_HOUR - C.FIRST_PM_CUT_HOUR;
  let hour = C.FIRST_PM_CUT_HOUR + (idx < 0 ? 0 : idx * span / Math.max(n, 1));
  hour = Math.round(hour * 2) / 2;
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

  ['Managers', 'Bar', 'Servers', 'Hosts', 'Barbacks'].forEach(title => {
    const rows = [];
    people.forEach(p => {
      const all = ['AM', 'PM'].map(s => p.cells[d][s]).filter(Boolean);
      const timedRole = (all.find(c => c.start != null) && cellRole_(p, all.find(c => c.start != null))) || p.role;
      // Notes like "volume" stay with their first shift.
      const cs = all.filter(c => (c.start != null ? cellRole_(p, c) : timedRole) === title);
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

  // Calendar descriptions take basic HTML, so the events section can be bold.
  const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  while (lines.length && !lines[lines.length - 1]) lines.pop();
  const html = lines.map(esc);
  if (dayEvents && dayEvents.length) {
    html.push('', '<b>EVENTS / PARTIES</b>');
    dayEvents.forEach(e => html.push('<b>• ' + esc(e) + '</b>'));
  }
  return html.join('<br>');
}

// For local testing in Node
if (typeof module !== 'undefined') module.exports = { buildShifts, parseCell_, CONFIG };