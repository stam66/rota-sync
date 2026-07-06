/**
 * rota-sync — serves per-person iCalendar feeds (and JSON) from the
 * Echo and HF rota Google Sheets.
 *
 * Deploy as a Web App (execute as: Me, access: Anyone) and hand each
 * person their personal feed URL, which they subscribe to in Outlook
 * on the web (nhs.net): Calendar -> Add calendar -> Subscribe from web.
 *
 * Endpoints (all via the web app URL):
 *   ?person=SK&token=xxxx              -> ICS feed for SK (both rotas)
 *   ?person=ALL&token=xxxx             -> ICS feed of every assignment
 *   ?person=SK&token=xxxx&format=json  -> same data as JSON (for the PWA)
 *
 * One-time setup (run from the editor):
 *   1. setup()        — generates the secret used to mint tokens
 *   2. deploy as web app, then paste the web app URL into the
 *      WEB_APP_URL script property (Project settings -> Script properties)
 *   3. listFeedUrls() — logs every person's personal URL to hand out
 */

const CONFIG = {
  timezone: 'Europe/London',

  // How far the feed looks back/ahead, in days.
  daysBack: 30,
  daysAhead: 400,

  rotas: [
    {
      name: 'Echo',
      spreadsheetId: '1zaCKRSNzH82PZSl9zs1-ihcFiOQDDg6yRwNl7MH6R3M',
      // Role columns, matched against the header row (whitespace-insensitive,
      // matched by "starts with"). start/end omitted -> all-day event.
      // TODO: adjust labels and times to what the codes really mean.
      roles: [
        { match: 'COTWE',          label: 'Consultant on call (weekend)' },
        { match: 'COTW',           label: 'Consultant of the week' },
        { match: 'Valve',          label: 'Valve' },
        { match: 'Endocarditis',   label: 'Endocarditis' }, // pre-2026 tabs
        { match: 'IE MDT',         label: 'IE MDT' },
        { match: 'IE',             label: 'Endocarditis' },
        { match: 'HF',             label: 'Heart failure' },
        { match: 'ECHO cover',     label: 'ECHO cover', start: '08:00', end: '17:00' },
        // Weeknights 17:00-08:00; weekends & bank holidays are 24h (08:00-08:00).
        { match: 'ECHO on call',   label: 'ECHO on call', start: '17:00', end: '08:00', weekend24h: true },
        { match: 'CIU TOE',        label: 'CIU TOE' },
        { match: 'TOE Fellows',    label: 'TOE Fellows' }, // pre-2026 tabs
        { match: 'TOEF',           label: 'TOE Fellows' }, // TOEF1 / TOEF2
        { match: 'Structural MDT', label: 'Structural MDT' },
        { match: 'ECHO MDT',       label: 'ECHO MDT' },
      ],
      // Per-person status columns (A/L, OOO, exams...). Header must equal
      // the initials exactly.
      people: ['CD', 'RR', 'SK', 'KG', 'BR', 'NL', 'JO', 'LB', 'CA'],
    },
    {
      name: 'HF',
      spreadsheetId: '10dl-fOFzD17gjbG8wLK9ZtDTSbWQdrFH4R6B6P4ybJU',
      roles: [
        { match: 'COTWE',      label: 'Consultant on call (weekend)' },
        { match: 'COTW',       label: 'Consultant of the week' },
        { match: 'HF ROTA',    label: 'HF of the week' },
        { match: 'Vetting HF', label: 'Vetting HF' },
      ],
      people: ['JW', 'SK', 'GCW', 'LAM', 'ADS', 'PS'],
    },
  ],

  // Values in person columns that should NOT become calendar events.
  ignoredStatuses: ['BH'],

  // Roles that may legitimately coincide with leave (triaging flexes around
  // it), matched by label. Kept in step with the sheets' red-cell rules.
  conflictExemptRoles: ['Vetting HF'],

  // Weekly MediRota change report: diffs the rotas against the snapshot
  // taken at the last report and emails the admin what changed, plus any
  // duty-while-on-leave conflicts. Run setupChangeReport() once to grant
  // permissions, take the baseline snapshot and install the weekly trigger.
  report: {
    recipients: [], // TODO: e.g. ['rota.admin@nhs.net'] — empty = report only logs
    subjectPrefix: '[Rota] ',
    weekday: 'MONDAY',
    hour: 7,          // trigger fires between 7 and 8 am
    lookBackDays: 14, // include retroactive edits this far back
    conflictDays: 90, // how far ahead conflicts are listed
    sendIfEmpty: false,
  },

  // Invite mode: shifts are mirrored into a Google Calendar with the
  // person's nhs.net address as guest, so they arrive in the person's REAL
  // work calendar as meeting invitations (unlike the subscribed ICS feed,
  // which lives in a separate overlay calendar). Opt-in per person: only
  // initials listed in `emails` get invites. Run setupInviteCalendar() once,
  // then syncInvites() runs daily via the trigger it installs.
  invites: {
    calendarName: 'Rota sync',
    // How far ahead invites are sent. Start small: the first sync emails one
    // invitation per shift in the window. Raise once you're happy.
    windowDays: 28,
    emails: {
      // SK: 'someone@nhs.net',
    },
  },
};

// ---------------------------------------------------------------- web app

function doGet(e) {
  const params = (e && e.parameter) || {};
  const person = (params.person || '').trim().toUpperCase();
  const token = (params.token || '').trim();
  const format = (params.format || 'ics').toLowerCase();

  if (!person) return textResponse('Missing ?person= parameter', 400);
  if (token !== personalToken(person)) return textResponse('Invalid token', 403);

  let events = collectEvents().filter(function (ev) {
    return person === 'ALL' || ev.person === person;
  });

  if (format === 'json') {
    // Optional yyyy-MM-dd window params keep mobile payloads small.
    if (params.from) events = events.filter(function (e) { return (e.endDate || e.date) >= params.from; });
    if (params.to) events = events.filter(function (e) { return e.date <= params.to; });
    return ContentService.createTextOutput(
      JSON.stringify({ person: person, generated: new Date().toISOString(), events: events })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(buildIcs(person, events))
    .setMimeType(ContentService.MimeType.ICAL);
}

function textResponse(message) {
  return ContentService.createTextOutput(message).setMimeType(ContentService.MimeType.TEXT);
}

// ---------------------------------------------------------------- parsing

/**
 * Reads every year tab of every configured rota and returns a flat,
 * normalized list of events:
 *   { rota, kind: 'shift'|'status', date: 'yyyy-MM-dd', endDate, role,
 *     person, summary, note, start, end }
 * Results are cached for 10 minutes.
 */
function collectEvents() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('events-v1');
  if (cached) return JSON.parse(cached);

  const now = new Date();
  const from = new Date(now.getTime() - CONFIG.daysBack * 86400000);
  const to = new Date(now.getTime() + CONFIG.daysAhead * 86400000);

  let events = [];
  CONFIG.rotas.forEach(function (rota) {
    const ss = SpreadsheetApp.openById(rota.spreadsheetId);
    ss.getSheets().forEach(function (sheet) {
      const name = sheet.getName().trim();
      if (!/^\d{4}$/.test(name)) return; // only year tabs
      const year = parseInt(name, 10);
      if (year < from.getFullYear() || year > to.getFullYear()) return;
      events = events.concat(parseSheet(rota, sheet, from, to));
    });
  });

  events = mergeConsecutiveStatuses(events);
  events.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  try { cache.put('events-v1', JSON.stringify(events), 600); } catch (err) { /* too big for cache: fine */ }
  return events;
}

function parseSheet(rota, sheet, from, to) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function (h) {
    return String(h).replace(/\s+/g, ' ').trim();
  });

  // Map columns to roles (longest match first so COTWE wins over COTW)
  // and to person-status columns.
  const roleCols = {}, personCols = {};
  const rolesByLength = rota.roles.slice().sort(function (a, b) { return b.match.length - a.match.length; });
  headers.forEach(function (header, col) {
    if (!header) return;
    if (rota.people.indexOf(header.toUpperCase()) !== -1) {
      personCols[col] = header.toUpperCase();
      return;
    }
    for (let i = 0; i < rolesByLength.length; i++) {
      if (header.toLowerCase().indexOf(rolesByLength[i].match.toLowerCase()) === 0) {
        roleCols[col] = rolesByLength[i];
        return;
      }
    }
  });

  const dateCol = findDateColumn(values);
  if (dateCol === -1) return [];

  const eventsCol = headers.map(function (h) { return h.toLowerCase(); }).indexOf('events');

  const events = [];
  for (let r = 1; r < values.length; r++) {
    const d = values[r][dateCol];
    if (!(d instanceof Date) || isNaN(d) || d < from || d > to) continue;
    const dateStr = Utilities.formatDate(d, CONFIG.timezone, 'yyyy-MM-dd');
    const isWeekend = d.getDay() === 0 || d.getDay() === 6;
    const isBH = eventsCol !== -1 && /\bBH\b/i.test(String(values[r][eventsCol]));

    Object.keys(roleCols).forEach(function (col) {
      const cell = String(values[r][col]).trim();
      if (!cell) return;
      const role = roleCols[col];
      const allDayShift = role.weekend24h && (isWeekend || isBH);
      const parsed = parseAssignmentCell(cell);
      parsed.people.forEach(function (initials) {
        events.push({
          rota: rota.name, kind: 'shift', date: dateStr, endDate: dateStr,
          role: role.label, person: initials,
          summary: role.label + ' (' + rota.name + ')',
          note: parsed.note,
          start: allDayShift ? '08:00' : (role.start || null),
          end: allDayShift ? '08:00' : (role.end || null),
        });
      });
    });

    Object.keys(personCols).forEach(function (col) {
      const cell = String(values[r][col]).trim();
      if (!cell || CONFIG.ignoredStatuses.indexOf(cell.toUpperCase()) !== -1) return;
      events.push({
        rota: rota.name, kind: 'status', date: dateStr, endDate: dateStr,
        role: null, person: personCols[col],
        summary: cell + ' - ' + personCols[col],
        note: '', start: null, end: null,
      });
    });
  }
  return events;
}

/** The date lives in column A (Echo) or B (HF); find it by content. */
function findDateColumn(values) {
  for (let col = 0; col < Math.min(3, values[0].length); col++) {
    let hits = 0;
    for (let r = 1; r < Math.min(15, values.length); r++) {
      if (values[r][col] instanceof Date && !isNaN(values[r][col])) hits++;
    }
    if (hits >= 5) return col;
  }
  return -1;
}

/**
 * "NL (swap with KG)" -> people [NL], note "swap with KG"
 * "TP/CA"             -> people [TP, CA]
 * "Filiz, Fatma"      -> people [FILIZ, FATMA]  (fellows go by first name)
 * "Lily off"          -> people []              (free-text notes are dropped)
 */
function parseAssignmentCell(cell) {
  let note = '';
  const withNote = cell.match(/^(.*?)\s*\((.+)\)\s*$/);
  let body = cell;
  if (withNote) { body = withNote[1].trim(); note = withNote[2].trim(); }
  // Initials: 1-4 letters starting uppercase (SK, JWh). Names: Filiz, Fatma.
  const people = body.split(/[\/+,&?]/).map(function (t) { return t.trim(); })
    .filter(function (t) { return /^[A-Z][A-Za-z]{0,3}$/.test(t) || /^[A-Z][a-z]+$/.test(t); })
    .map(function (t) { return t.toUpperCase(); });
  return { people: people, note: note };
}

/** Collapse runs of identical consecutive all-day status events (e.g. two
 *  weeks of A/L, or JO's long-term OOO) into single multi-day events. */
function mergeConsecutiveStatuses(events) {
  const statuses = events.filter(function (e) { return e.kind === 'status'; });
  const rest = events.filter(function (e) { return e.kind !== 'status'; });

  statuses.sort(function (a, b) {
    const ka = a.person + '|' + a.summary + '|' + a.date;
    const kb = b.person + '|' + b.summary + '|' + b.date;
    return ka < kb ? -1 : ka > kb ? 1 : 0;
  });

  const merged = [];
  statuses.forEach(function (e) {
    const last = merged[merged.length - 1];
    if (last && last.person === e.person && last.summary === e.summary &&
        nextDay(last.endDate) === e.date) {
      last.endDate = e.date;
    } else {
      merged.push(e);
    }
  });
  return rest.concat(merged);
}

function nextDay(dateStr) {
  const parts = dateStr.split('-').map(Number);
  const d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + 1));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// ---------------------------------------------------------------- ICS

function buildIcs(person, events) {
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//rota-sync//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'X-WR-CALNAME:' + icsEscape('Rota - ' + person),
    'X-WR-TIMEZONE:' + CONFIG.timezone,
    'BEGIN:VTIMEZONE',
    'TZID:Europe/London',
    'BEGIN:DAYLIGHT',
    'TZOFFSETFROM:+0000', 'TZOFFSETTO:+0100', 'TZNAME:BST',
    'DTSTART:19700329T010000', 'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
    'END:DAYLIGHT',
    'BEGIN:STANDARD',
    'TZOFFSETFROM:+0100', 'TZOFFSETTO:+0000', 'TZNAME:GMT',
    'DTSTART:19701025T020000', 'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
    'END:STANDARD',
    'END:VTIMEZONE',
  ];

  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");

  events.forEach(function (ev) {
    const day = ev.date.replace(/-/g, '');
    lines.push('BEGIN:VEVENT');
    lines.push(foldLine('UID:' + eventUid(ev)));
    lines.push('DTSTAMP:' + stamp);
    if (ev.start && ev.end) {
      // An end time at or before the start means the shift runs overnight.
      const endDay = ev.end <= ev.start ? nextDay(ev.date).replace(/-/g, '') : day;
      lines.push('DTSTART;TZID=Europe/London:' + day + 'T' + ev.start.replace(':', '') + '00');
      lines.push('DTEND;TZID=Europe/London:' + endDay + 'T' + ev.end.replace(':', '') + '00');
    } else {
      lines.push('DTSTART;VALUE=DATE:' + day);
      lines.push('DTEND;VALUE=DATE:' + nextDay(ev.endDate || ev.date).replace(/-/g, ''));
    }
    const summary = (person === 'ALL' && ev.kind === 'shift')
      ? ev.person + ': ' + ev.summary
      : ev.summary;
    lines.push(foldLine('SUMMARY:' + icsEscape(summary)));
    if (ev.note) lines.push(foldLine('DESCRIPTION:' + icsEscape(ev.note)));
    lines.push('TRANSP:TRANSPARENT'); // don't block free/busy
    lines.push('END:VEVENT');
  });

  lines.push('END:VCALENDAR');
  return lines.join('\r\n') + '\r\n';
}

function icsEscape(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/;/g, '\\;')
    .replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

/** RFC 5545: lines longer than 75 octets are folded with CRLF + space. */
function foldLine(line) {
  const out = [];
  while (line.length > 73) {
    out.push(line.substring(0, 73));
    line = ' ' + line.substring(73);
  }
  out.push(line);
  return out.join('\r\n');
}

// ------------------------------------------------------- change report

const SNAPSHOT_FILENAME = 'rota-sync-snapshot.json';

/** Run once from the editor: grants Mail/Drive permission, takes the
 *  baseline snapshot and installs the weekly trigger. */
function setupChangeReport() {
  const hasTrigger = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendChangeReport';
  });
  if (!hasTrigger) {
    ScriptApp.newTrigger('sendChangeReport').timeBased()
      .onWeekDay(ScriptApp.WeekDay[CONFIG.report.weekday])
      .atHour(CONFIG.report.hour).create();
  }
  if (!loadSnapshot()) {
    saveSnapshot(currentState());
    Logger.log('Baseline snapshot taken.');
  }
  Logger.log('Weekly report trigger installed (%s %s:00). Recipients: %s',
    CONFIG.report.weekday, CONFIG.report.hour,
    CONFIG.report.recipients.join(', ') || '(none — logs only)');
}

/** Preview the report in the log without emailing or moving the snapshot. */
function previewChangeReport() {
  const old = loadSnapshot();
  if (!old) { Logger.log('No baseline yet — run setupChangeReport() first.'); return; }
  const diff = diffStates(old, currentState());
  Logger.log(JSON.stringify(diff, null, 2));
}

/** Weekly trigger target. Emails the diff since the last report, then
 *  advances the snapshot (only after a successful send). */
function sendChangeReport() {
  const cur = currentState();
  const old = loadSnapshot();
  if (!old) { saveSnapshot(cur); Logger.log('No baseline; snapshot taken, no report.'); return; }

  const diff = diffStates(old, cur);
  const today = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  const horizon = Utilities.formatDate(
    new Date(Date.now() + CONFIG.report.conflictDays * 86400000), CONFIG.timezone, 'yyyy-MM-dd');
  const conflicts = computeConflicts().filter(function (c) {
    return c.date >= today && c.date <= horizon;
  });

  const total = diff.duties.length + diff.statuses.length;
  if (!total && !conflicts.length && !CONFIG.report.sendIfEmpty) {
    saveSnapshot(cur);
    Logger.log('No changes and no conflicts — report skipped.');
    return;
  }

  const subject = CONFIG.report.subjectPrefix +
    (total ? total + ' change' + (total === 1 ? '' : 's') : 'no changes') +
    (conflicts.length ? ', ' + conflicts.length + ' conflict' + (conflicts.length === 1 ? '' : 's') : '');
  const html = reportHtml(diff, conflicts);

  if (CONFIG.report.recipients.length) {
    MailApp.sendEmail({
      to: CONFIG.report.recipients.join(','),
      subject: subject,
      htmlBody: html,
    });
    Logger.log('Report sent to %s.', CONFIG.report.recipients.join(', '));
  } else {
    Logger.log('No recipients configured; report below:\n%s', html);
  }
  saveSnapshot(cur);
}

/** Compact picture of the rotas: duty assignments and per-day statuses. */
function currentState() {
  const from = Utilities.formatDate(
    new Date(Date.now() - CONFIG.report.lookBackDays * 86400000), CONFIG.timezone, 'yyyy-MM-dd');
  const duties = {}, statuses = {};
  collectEvents().forEach(function (e) {
    if (e.date < from && (e.endDate || e.date) < from) return;
    if (e.kind === 'shift') {
      const k = e.date + '|' + e.rota + '|' + e.role;
      duties[k] = duties[k] ? duties[k] + '+' + e.person : e.person;
    } else {
      const label = e.summary.split(' - ')[0];
      for (let d = e.date; d <= (e.endDate || e.date); d = nextDay(d)) {
        if (d >= from) statuses[e.person + '|' + d] = label;
      }
    }
  });
  return { from: from, duties: duties, statuses: statuses };
}

function diffStates(oldState, cur) {
  // Only compare dates both snapshots could see.
  const from = oldState.from > cur.from ? oldState.from : cur.from;

  const duties = [];
  const dutyKeys = {};
  Object.keys(oldState.duties).concat(Object.keys(cur.duties)).forEach(function (k) { dutyKeys[k] = true; });
  Object.keys(dutyKeys).forEach(function (k) {
    const date = k.split('|')[0];
    if (date < from) return;
    const was = oldState.duties[k] || null, now = cur.duties[k] || null;
    if (was !== now) {
      duties.push({ date: date, rota: k.split('|')[1], role: k.split('|')[2], was: was, now: now });
    }
  });
  duties.sort(function (a, b) { return a.date < b.date ? -1 : 1; });

  // Status diffs, compressed into runs of consecutive days.
  const perPerson = {}; // 'SK|A/L (added)' -> [dates]
  const statusKeys = {};
  Object.keys(oldState.statuses).concat(Object.keys(cur.statuses)).forEach(function (k) { statusKeys[k] = true; });
  Object.keys(statusKeys).forEach(function (k) {
    const date = k.split('|')[1];
    if (date < from) return;
    const was = oldState.statuses[k] || null, now = cur.statuses[k] || null;
    if (was === now) return;
    const desc = now === null ? was + ' removed'
      : was === null ? now + ' added'
      : was + ' → ' + now;
    const key = k.split('|')[0] + '|' + desc;
    (perPerson[key] = perPerson[key] || []).push(date);
  });
  const statuses = Object.keys(perPerson).sort().map(function (key) {
    return {
      person: key.split('|')[0],
      change: key.split('|')[1],
      ranges: compressDates(perPerson[key]),
    };
  });

  return { duties: duties, statuses: statuses };
}

/** ['2026-07-01','2026-07-02','2026-07-04'] -> '1–2 Jul, 4 Jul' */
function compressDates(dates) {
  dates.sort();
  const runs = [];
  dates.forEach(function (d) {
    const last = runs[runs.length - 1];
    if (last && nextDay(last.end) === d) last.end = d;
    else runs.push({ start: d, end: d });
  });
  return runs.map(function (r) {
    return r.start === r.end ? prettyDate(r.start) : prettyDate(r.start) + ' – ' + prettyDate(r.end);
  }).join(', ');
}

function prettyDate(iso) {
  const p = iso.split('-').map(Number);
  return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2]), CONFIG.timezone, 'EEE d MMM');
}

function reportHtml(diff, conflicts) {
  const esc = function (s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  };
  const th = 'style="text-align:left;padding:4px 10px;border-bottom:2px solid #ccc"';
  const td = 'style="padding:4px 10px;border-bottom:1px solid #eee"';
  let html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">';
  html += '<p>Rota changes since the last report, for updating MediRota.</p>';

  if (diff.duties.length) {
    html += '<h3>Duty changes (' + diff.duties.length + ')</h3><table style="border-collapse:collapse">' +
      '<tr><th ' + th + '>Date</th><th ' + th + '>Rota</th><th ' + th + '>Duty</th><th ' + th + '>Was</th><th ' + th + '>Now</th></tr>';
    diff.duties.forEach(function (c) {
      html += '<tr><td ' + td + '>' + prettyDate(c.date) + '</td><td ' + td + '>' + esc(c.rota) +
        '</td><td ' + td + '>' + esc(c.role) + '</td><td ' + td + '>' + esc(c.was || '—') +
        '</td><td ' + td + '><b>' + esc(c.now || '—') + '</b></td></tr>';
    });
    html += '</table>';
  }

  if (diff.statuses.length) {
    html += '<h3>Leave / absence changes</h3><ul>';
    diff.statuses.forEach(function (s) {
      html += '<li><b>' + esc(s.person) + '</b>: ' + esc(s.change) + ' — ' + esc(s.ranges) + '</li>';
    });
    html += '</ul>';
  }

  if (!diff.duties.length && !diff.statuses.length) {
    html += '<p>No rota changes since the last report.</p>';
  }

  if (conflicts.length) {
    html += '<h3 style="color:#b00020">Conflicts — rostered while marked away (next ' +
      CONFIG.report.conflictDays + ' days)</h3><table style="border-collapse:collapse">' +
      '<tr><th ' + th + '>Date</th><th ' + th + '>Who</th><th ' + th + '>Duty</th><th ' + th + '>Marked</th></tr>';
    conflicts.forEach(function (c) {
      html += '<tr><td ' + td + '>' + prettyDate(c.date) + '</td><td ' + td + '>' + esc(c.person) +
        '</td><td ' + td + '>' + esc(c.duty) + '</td><td ' + td + '>' + esc(c.leave) + '</td></tr>';
    });
    html += '</table>';
  }

  html += '<p style="color:#777;font-size:12px">Generated ' +
    Utilities.formatDate(new Date(), CONFIG.timezone, 'EEE d MMM yyyy, HH:mm') +
    ' · <a href="https://docs.google.com/spreadsheets/d/' + CONFIG.rotas[0].spreadsheetId +
    '">Echo rota</a> · <a href="https://docs.google.com/spreadsheets/d/' + CONFIG.rotas[1].spreadsheetId +
    '">HF rota</a></p></div>';
  return html;
}

function loadSnapshot() {
  const files = DriveApp.getFilesByName(SNAPSHOT_FILENAME);
  if (!files.hasNext()) return null;
  try { return JSON.parse(files.next().getBlob().getDataAsString()); }
  catch (err) { return null; }
}

function saveSnapshot(state) {
  const files = DriveApp.getFilesByName(SNAPSHOT_FILENAME);
  const json = JSON.stringify(state);
  if (files.hasNext()) files.next().setContent(json);
  else DriveApp.createFile(SNAPSHOT_FILENAME, json, 'application/json');
}

// ---------------------------------------------------------------- invites

/** Run once: creates the sync calendar and a daily 6am trigger for
 *  syncInvites(). Requires the script project timezone to be Europe/London
 *  (Project settings), so shift times land correctly. */
function setupInviteCalendar() {
  let cal = getInviteCalendar();
  if (!cal) cal = CalendarApp.createCalendar(CONFIG.invites.calendarName);

  const hasTrigger = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'syncInvites';
  });
  if (!hasTrigger) {
    ScriptApp.newTrigger('syncInvites').timeBased().everyDays(1).atHour(6).create();
  }
  Logger.log('Calendar "%s" ready; daily syncInvites trigger installed.', cal.getName());
}

function getInviteCalendar() {
  const cals = CalendarApp.getCalendarsByName(CONFIG.invites.calendarName);
  return cals.length ? cals[0] : null;
}

/**
 * Mirrors upcoming shifts into the sync calendar with the person's nhs.net
 * address as guest. Google emails the invitation; Outlook auto-adds it to
 * the person's primary calendar as tentative. When the sheet changes, the
 * old event is deleted (guest gets a cancellation) and a new one created.
 * Only shift events for people opted in via CONFIG.invites.emails are sent;
 * leave/status entries are not.
 */
function syncInvites() {
  const cal = getInviteCalendar();
  if (!cal) throw new Error('Run setupInviteCalendar() first');
  const emails = CONFIG.invites.emails;
  const now = new Date();
  const horizon = new Date(now.getTime() + CONFIG.invites.windowDays * 86400000);

  const wanted = {};
  collectEvents().forEach(function (ev) {
    if (ev.kind !== 'shift' || !emails[ev.person]) return;
    const t = eventTimes(ev);
    if (t.start < now || t.start > horizon) return;
    wanted[eventUid(ev)] = { ev: ev, t: t };
  });

  // Reconcile against what's already in the calendar. The uid encodes
  // rota/role/date/person, so any change shows up as delete + create.
  cal.getEvents(now, horizon).forEach(function (ge) {
    const uid = ge.getTag('rotaUid');
    if (!uid) return;
    if (wanted[uid]) {
      delete wanted[uid]; // already synced
    } else {
      ge.deleteEvent(); // guest receives a cancellation
    }
  });

  let created = 0;
  Object.keys(wanted).forEach(function (uid) {
    const w = wanted[uid];
    const opts = {
      guests: emails[w.ev.person],
      sendInvites: true,
      description: w.ev.note || '',
    };
    const ge = w.t.allDay
      ? cal.createAllDayEvent(w.ev.summary, w.t.start, opts)
      : cal.createEvent(w.ev.summary, w.t.start, w.t.end, opts);
    ge.setTag('rotaUid', uid);
    created++;
  });
  Logger.log('syncInvites: %s invites sent/updated.', created);
}

/** Concrete start/end Dates for an event, honouring overnight shifts.
 *  Uses the script project timezone (must be Europe/London). */
function eventTimes(ev) {
  const p = ev.date.split('-').map(Number);
  if (!ev.start || !ev.end) {
    return { allDay: true, start: new Date(p[0], p[1] - 1, p[2]), end: null };
  }
  const s = ev.start.split(':').map(Number);
  const e = ev.end.split(':').map(Number);
  const overnight = ev.end <= ev.start ? 1 : 0;
  return {
    allDay: false,
    start: new Date(p[0], p[1] - 1, p[2], s[0], s[1]),
    end: new Date(p[0], p[1] - 1, p[2] + overnight, e[0], e[1]),
  };
}

function eventUid(ev) {
  return [ev.rota, ev.kind, ev.role || ev.summary, ev.date, ev.person]
    .join('-').replace(/[^A-Za-z0-9-]/g, '_') + '@rota-sync';
}

// ---------------------------------------------------------------- admin

/** Run once from the editor: creates the token secret and warms the parser
 *  (which also triggers the permission grant for both spreadsheets). */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) {
    let secret = '';
    for (let i = 0; i < 32; i++) secret += 'abcdefghijklmnopqrstuvwxyz0123456789'.charAt(Math.floor(Math.random() * 36));
    props.setProperty('SECRET', secret);
  }
  const events = collectEvents();
  Logger.log('Secret ready. Parsed %s events across both rotas.', events.length);
}

/** Logs every person's personal feed URL. Set the WEB_APP_URL script
 *  property to your deployment URL first (Project settings -> Script properties). */
function listFeedUrls() {
  const base = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL')
    || ScriptApp.getService().getUrl();
  if (!base) { Logger.log('Deploy the web app first, then set WEB_APP_URL.'); return; }

  const all = {};
  CONFIG.rotas.forEach(function (r) { r.people.forEach(function (p) { all[p] = true; }); });
  Object.keys(all).sort().concat(['ALL']).forEach(function (p) {
    Logger.log('%s: %s?person=%s&token=%s', p, base, p, personalToken(p));
  });
}

function personalToken(person) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, secret + ':' + person.toUpperCase());
  return digest.slice(0, 8).map(function (b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

/**
 * Mirrors the sheets' red-cell rule: flags any day where someone is both
 * on leave (their P–V / H–M column) and assigned a duty. Run from the
 * editor to audit; the same list will feed the MediRota change report.
 */
function listConflicts() {
  const conflicts = computeConflicts();
  Logger.log('%s conflicts found:', conflicts.length);
  conflicts.forEach(function (c) {
    Logger.log('%s  %s assigned "%s" but marked "%s"', c.date, c.person, c.duty, c.leave);
  });
  return conflicts;
}

function computeConflicts() {
  const events = collectEvents();

  const leave = {}; // 'SK|2026-03-05' -> 'A/L'
  events.filter(function (e) { return e.kind === 'status'; }).forEach(function (e) {
    for (let d = e.date; d <= (e.endDate || e.date); d = nextDay(d)) {
      leave[e.person + '|' + d] = e.summary;
    }
  });

  return events.filter(function (e) {
    return e.kind === 'shift' &&
      CONFIG.conflictExemptRoles.indexOf(e.role) === -1 &&
      leave[e.person + '|' + e.date];
  }).map(function (e) {
    return { date: e.date, person: e.person, duty: e.summary, leave: leave[e.person + '|' + e.date] };
  });
}

/** Quick sanity check from the editor: logs SK's next 20 events. */
function testParse() {
  const events = collectEvents().filter(function (e) { return e.person === 'SK'; });
  Logger.log('%s events for SK; first 20:', events.length);
  events.slice(0, 20).forEach(function (e) {
    Logger.log('%s  %s%s', e.date, e.summary, e.note ? ' [' + e.note + ']' : '');
  });
}
