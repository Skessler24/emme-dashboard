/**
 * Reads a provider schedule out of whatever was uploaded, without needing one fixed layout.
 *
 * It looks at what each piece of text IS, not where it sits: a weekday name, a time of day, a status
 * (Full, Open 3x, Lunch...), a date, or anything else (a client's name, a note). Then it puts the pieces
 * together using whichever shape the file has:
 *   - a weekday heading with rows underneath (one table per day, or lines in a document),
 *   - one row per appointment with the weekday written in the row,
 *   - a grid with weekdays across the top and times down the side,
 *   - a single line of text such as "Mon 9:00 AM Full Jamie P".
 *
 * Input is a matrix of cells (spreadsheet rows, table rows, or text lines split into cells).
 * Loaded by app5.html in the browser and by the API tests in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.EmmeScheduleImport = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DAY_KEYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  var DAY_LABELS = { monday: 'Monday', tuesday: 'Tuesday', wednesday: 'Wednesday', thursday: 'Thursday', friday: 'Friday', saturday: 'Saturday', sunday: 'Sunday' };

  var DAY_WORD = '(mon(?:day)?s?|tue(?:s|sday)?s?|wed(?:nesday)?s?|thu(?:r|rs|rsday)?s?|fri(?:day)?s?|sat(?:urday)?s?|sun(?:day)?s?)';
  var DAY_AT_START = new RegExp('^\\W*' + DAY_WORD + '(?![a-z])', 'i');
  var DAY_ANYWHERE = new RegExp('(^|[^a-z])' + DAY_WORD + '(?![a-z])', 'i');

  var TIME_PART = '(1[0-2]|0?[1-9])(?::([0-5]\\d))?(?::[0-5]\\d)?\\s*([ap])\\.?\\s*m\\.?';
  var TIME_24 = '([01]?\\d|2[0-3]):([0-5]\\d)(?::[0-5]\\d)?';
  var SEP = '\\s*(?:-|–|—|to|until)\\s*';
  var ANY_TIME = '(?:' + TIME_PART + '|' + TIME_24 + ')';
  var TIME_RANGE = new RegExp('(?<![\\d:./-])(' + ANY_TIME + ')' + SEP + '(' + ANY_TIME + ')(?![\\d:/])', 'i');
  var TIME_BARE_RANGE = new RegExp('(?<![\\d:./-])(1[0-2]|0?[1-9])' + SEP + '(' + TIME_PART + ')(?![a-z\\d])', 'i');
  var TIME_AM_PM = new RegExp('(?<![\\d:./-])' + TIME_PART + '(?![a-z])', 'i');
  var TIME_COLON = new RegExp('(?<![\\d:./-])' + TIME_24 + '(?![\\d:/.-])');
  var TIME_BARE_HOUR = /^\\s*(1[0-2]|[1-9])\\s*$/;

  var DATE_RE = /(?:^|[^\d])((?:0?[1-9]|1[0-2])[\/.-](?:0?[1-9]|[12]\d|3[01])[\/.-](?:\d{4}|\d{2})|\d{4}-\d{2}-\d{2})(?![\d])/;

  var HEADER_WORDS = {
    time: 1, times: 1, appointment: 1, appointments: 1, status: 1, client: 1, clients: 1, name: 1, names: 1, notes: 1, note: 1,
    date: 1, last: 1, visit: 1, day: 1, slot: 1, provider: 1, staff: 1, schedule: 1, current: 1, weekly: 1, availability: 1,
    openings: 1, caseload: 1, of: 1, the: 1, and: 1, for: 1, a: 1, an: 1, to: 1, with: 1
  };

  // ---------------------------------------------------------------------------
  // Status
  // ---------------------------------------------------------------------------

  function weeksList(raw) {
    return String(raw || '').match(/[1-4]/g) || [];
  }

  /** The status wording the dashboard uses, for any reasonable way of writing it. Null when the text is not a status. */
  function canonicalStatus(text) {
    var t = String(text || '').toLowerCase().replace(/[‘’`]/g, "'").replace(/\s+/g, ' ').trim();
    if (!t) return null;
    if (/\bnot\s+open\b|\bclosed\b|\bunavailable\b|\bnot\s+available\b|\bno\s+sessions?\b|\bblocked\b/.test(t)) return 'Not Open';
    var open = t.match(/\bopen\b[^0-9a-z]*(?:(\d|one|two|three|four)\s*(?:x|times?|\/\s*(?:mo|month|week)))?\s*(?:\(?\s*(?:wks?\.?|weeks?)?\s*([1-4](?:\s*(?:,|&|and|\/|-)\s*[1-4])*)\s*\)?)?/);
    if (open) {
      var words = { one: 1, two: 2, three: 3, four: 4 };
      var count = open[1] ? (words[open[1]] || Number(open[1])) : 0;
      var weeks = weeksList(open[2]);
      if (!count && weeks.length) count = weeks.length;
      if (count === 4 || (!count && /\b(all|every)\b/.test(t))) return 'Open 4x';
      if (count === 3) return weeks.length === 3 ? 'Open 3x (' + weeks.join(',') + ')' : 'Open 3x (1,2,3)';
      if (count === 2) return weeks.length === 2 ? 'Open 2x (' + weeks.join('&') + ')' : 'Open 2x (1&3)';
      if (count === 1) return weeks.length === 1 ? 'Open 1x (' + weeks[0] + ')' : 'Open 1x (1)';
      if (!count) return 'Open 4x';
    }
    if (/\bscheduling\s+in\s+progress\b|\bin\s+progress\b|\bpending\b|\bhold(ing)?\b|\btentative\b/.test(t)) return 'Scheduling In Progress';
    if (/\blunch\b|\bpaperwork\b|\bbreak\b|\badmin\b/.test(t)) return 'Lunch';
    if (/\bmake[\s-]?ups?\b/.test(t)) return 'Makeup';
    if (/\bfull\b|\bbooked\b|\bfilled\b|\btaken\b/.test(t)) return 'Full';
    if (/\bavailable\b|\bavail\b|\bopening\b/.test(t)) return 'Open 4x';
    return null;
  }

  var STATUS_FIND = new RegExp(
    '\\bnot\\s+open\\b|\\bscheduling\\s+in\\s+progress\\b|\\bopen\\b(?:[^0-9a-z]*(?:\\d|one|two|three|four)\\s*(?:x|times?))?\\s*(?:\\(\\s*(?:weeks?\\s*)?[1-4](?:\\s*(?:,|&|and|\\/|-)\\s*[1-4])*\\s*\\))?' +
    '|\\bfull\\b|\\bbooked\\b|\\blunch(?:\\s*\\/\\s*paperwork)?\\b|\\bmake[\\s-]?ups?\\b|\\bclosed\\b|\\bunavailable\\b|\\bavailable\\b', 'i');

  // ---------------------------------------------------------------------------
  // Time
  // ---------------------------------------------------------------------------

  function clock(hour, minute, suffix) {
    var h = hour % 12 || 12;
    return h + ':' + (minute < 10 ? '0' : '') + minute + ' ' + suffix;
  }

  function fromAmPm(h, m, ap) {
    var hour = Number(h);
    var suffix = /^p/i.test(ap) ? 'PM' : 'AM';
    return { hour: hour, minute: Number(m || 0), suffix: suffix };
  }

  /** A time with no AM or PM: working hours are 7 to 5, so 1 to 6 means afternoon. */
  function fromBare(hour, minute) {
    if (hour >= 13) return { hour: hour, minute: minute, suffix: 'PM' };
    if (hour === 0) return { hour: 12, minute: minute, suffix: 'AM' };
    if (hour === 12) return { hour: 12, minute: minute, suffix: 'PM' };
    return { hour: hour, minute: minute, suffix: hour <= 6 ? 'PM' : 'AM' };
  }

  function parseOneTime(piece) {
    var s = String(piece || '').trim();
    var m = s.match(new RegExp('^' + TIME_PART + '$', 'i'));
    if (m) return fromAmPm(m[1], m[2], m[3]);
    m = s.match(new RegExp('^' + TIME_24 + '$'));
    if (m) return fromBare(Number(m[1]), Number(m[2]));
    m = s.match(/^(1[0-2]|0?[1-9])$/);
    if (m) return fromBare(Number(m[1]), 0);
    return null;
  }

  function sameText(t) { return clock(t.hour, t.minute, t.suffix); }

  function hasMeridiem(text) { return /[ap]\.?\s*m/i.test(text); }

  function joinRange(first, second) {
    var a = parseOneTime(first);
    var b = parseOneTime(second);
    if (!a || !b) return '';
    if (!hasMeridiem(first) && hasMeridiem(second)) {
      a.suffix = b.suffix;
      if (a.hour % 12 > b.hour % 12 && b.suffix === 'PM') a.suffix = 'AM';
    }
    return sameText(a) + ' - ' + sameText(b);
  }

  /** Finds a time (or a range such as 9:00-10:00 AM) in text. Returns { text, start, end } with character positions. */
  function findTime(text) {
    var s = String(text || '');
    var m = s.match(TIME_RANGE);
    if (m) {
      var parts = m[0].split(new RegExp(SEP, 'i'));
      var joined = parts.length === 2 ? joinRange(parts[0], parts[1]) : '';
      if (joined) return { text: joined, start: m.index, end: m.index + m[0].length };
    }
    m = s.match(TIME_BARE_RANGE);
    if (m) {
      var bits = m[0].split(new RegExp(SEP, 'i'));
      var joinedBare = bits.length === 2 ? joinRange(bits[0], bits[1]) : '';
      if (joinedBare) return { text: joinedBare, start: m.index, end: m.index + m[0].length };
    }
    m = s.match(TIME_AM_PM);
    if (m) return { text: sameText(fromAmPm(m[1], m[2], m[3])), start: m.index, end: m.index + m[0].length };
    m = s.match(TIME_COLON);
    if (m) return { text: sameText(fromBare(Number(m[1]), Number(m[2]))), start: m.index, end: m.index + m[0].length };
    return null;
  }

  /** A whole cell that is just a time: "9", "0900", "0.375" (Excel's raw number), or anything findTime reads with nothing else around it. */
  function wholeCellTime(text) {
    var s = String(text || '').trim();
    if (!s) return null;
    var found = findTime(s);
    if (found && s.slice(0, found.start).replace(/\W/g, '') === '' && s.slice(found.end).replace(/\W/g, '') === '') return found.text;
    var frac = s.match(/^0?\.(\d{2,})$/);
    if (frac) {
      var minutes = Math.round(Number('0.' + frac[1]) * 1440);
      if (minutes > 0 && minutes < 1440) return sameText(fromBare(Math.floor(minutes / 60), minutes % 60));
    }
    var hours = s.match(/^(1[0-2]|0?[1-9])\s*(?:-|–|—|to)\s*(1[0-2]|0?[1-9])$/i);
    if (hours) return joinRange(hours[1], hours[2]);
    var compact = s.match(/^([01]\d|2[0-3])([0-5]\d)$/);
    if (compact) return sameText(fromBare(Number(compact[1]), Number(compact[2])));
    if (TIME_BARE_HOUR.test(s)) return sameText(fromBare(Number(s), 0));
    return null;
  }

  // ---------------------------------------------------------------------------
  // One piece of text, taken apart
  // ---------------------------------------------------------------------------

  function dayKeyOf(word) {
    var w = String(word || '').toLowerCase();
    for (var i = 0; i < DAY_KEYS.length; i++) {
      if (w.indexOf(DAY_KEYS[i].slice(0, 3)) === 0) return DAY_KEYS[i];
    }
    return '';
  }

  function cleanRest(text) {
    return String(text || '')
      .replace(/\b(appointment|appt)s?\s*times?\b/gi, ' ')
      .replace(/^[\s\-–—:;,|•*#>()[\]]+|[\s\-–—:;,|•*#>()[\]]+$/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  /** What one cell holds: { day, dayOnly, time, status, date, rest }. */
  function facets(cell) {
    var text = String(cell == null ? '' : cell).replace(/\u00a0/g, ' ').trim();
    var out = { day: '', dayOnly: false, time: '', status: '', date: '', rest: '', whole: text };
    if (!text) return out;

    var hold = text;

    var dateMatch = hold.match(DATE_RE);
    if (dateMatch) {
      out.date = dateMatch[1];
      hold = hold.replace(dateMatch[1], ' ');
    }

    var timeFound = findTime(hold);
    if (timeFound) {
      out.time = timeFound.text;
      hold = hold.slice(0, Math.max(0, timeFound.start)) + ' ' + hold.slice(timeFound.end);
    }

    var statusMatch = hold.match(STATUS_FIND);
    if (statusMatch) {
      var canon = canonicalStatus(statusMatch[0]);
      if (canon) {
        out.status = canon;
        hold = hold.replace(statusMatch[0], ' ');
      }
    }

    var dayMatch = hold.match(DAY_ANYWHERE);
    if (dayMatch) {
      var without = hold.replace(dayMatch[0].replace(/^[^a-z]/i, ''), ' ');
      var leftover = cleanRest(without).replace(/\bschedule\b/gi, '').replace(/[^a-z0-9]/gi, '');
      var atStart = DAY_AT_START.test(hold);
      var fullName = /^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)s?$/i.test(dayMatch[2]);
      if (!leftover) {
        out.day = dayKeyOf(dayMatch[2]);
        out.dayOnly = !out.time && !out.status;
        hold = without;
      } else if ((out.time || out.status) && (atStart || fullName)) {
        out.day = dayKeyOf(dayMatch[2]);
        hold = without;
      }
    }

    out.rest = cleanRest(hold);
    return out;
  }

  function isHeaderRest(text) {
    var words = String(text || '').toLowerCase().replace(/[^a-z\s]/g, ' ').split(/\s+/).filter(Boolean);
    if (!words.length) return false;
    return words.every(function (word) { return !!HEADER_WORDS[word]; });
  }

  // ---------------------------------------------------------------------------
  // Putting a whole file together
  // ---------------------------------------------------------------------------

  function addSlot(days, key, slot) {
    if (!days[key]) days[key] = [];
    days[key].push(slot);
  }

  function finishSlot(facet, time, restParts, scheduleNotes) {
    var notes = restParts.filter(Boolean).join(' ').replace(/\s{2,}/g, ' ').trim();
    var status = facet.status;
    if (!status && notes) status = 'Full';
    return { time: time || '', status: status || '', notes: notes, lastDay: facet.date || '', scheduleNotes: scheduleNotes || '' };
  }

  /**
   * rows: array of arrays of cell text. Returns { days: { monday: [{time,status,notes,lastDay}] }, slots, dayCount }.
   */
  function scheduleFromMatrix(rows) {
    var days = {};
    var currentDay = '';
    var grid = null;
    var slots = 0;
    var scheduleNoteCol = -1;

    (rows || []).forEach(function (row) {
      var cells = [];
      (row || []).forEach(function (value, col) {
        var text = String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim();
        if (text) cells.push({ col: col, text: text, f: facets(text), whole: wholeCellTime(text) });
      });
      if (!cells.length) { grid = null; return; }

      var dayOnly = cells.filter(function (c) { return c.f.dayOnly; });
      var anyTime = cells.some(function (c) { return c.f.time || c.whole; });
      var anyStatus = cells.some(function (c) { return c.f.status; });

      if (dayOnly.length >= 2) {
        grid = {};
        dayOnly.forEach(function (c) { grid[c.col] = c.f.day; });
        currentDay = '';
        return;
      }
      if (dayOnly.length === 1 && !anyTime && !anyStatus) {
        currentDay = dayOnly[0].f.day;
        grid = null;
        return;
      }

      var allRests = cells.map(function (c) { return c.f.rest; }).filter(Boolean);
      var headerLike = !anyTime && !anyStatus && allRests.length && allRests.every(isHeaderRest);
      if (headerLike) {
        cells.forEach(function (c) {
          if (/schedule\s+notes/i.test(c.text)) scheduleNoteCol = c.col;
        });
        return;
      }

      var rowDay = '';
      cells.forEach(function (c) { if (!rowDay && c.f.day) rowDay = c.f.day; });
      if (rowDay && !grid) currentDay = rowDay;

      if (grid) {
        var timeCell = null;
        cells.forEach(function (c) {
          if (timeCell || grid[c.col]) return;
          if (c.whole || c.f.time) timeCell = c;
        });
        var rowTime = timeCell ? (timeCell.whole || timeCell.f.time) : '';
        var inGrid = cells.filter(function (c) { return grid[c.col] && c !== timeCell; });
        if (!rowTime) {
          if (!inGrid.length) return;
          var timeInside = inGrid.filter(function (c) { return c.f.time; });
          if (!timeInside.length) return;
        }
        inGrid.forEach(function (c) {
          var parts = c.f;
          var slotTime = parts.time || rowTime;
          var rest = parts.time && !rowTime ? [parts.rest] : [parts.rest];
          if (!slotTime && !parts.status && !parts.rest) return;
          addSlot(days, grid[c.col], finishSlot(parts, slotTime, rest));
          slots++;
        });
        return;
      }

      var day = rowDay || currentDay;
      if (!day) return;

      var time = '';
      var status = '';
      var date = '';
      var scheduleNotes = '';
      var rest = [];
      cells.forEach(function (c) {
        if (c.col === scheduleNoteCol) {
          scheduleNotes = c.text;
          return;
        }
        if (!time) time = c.whole || c.f.time || '';
        if (!status && c.f.status) status = c.f.status;
        if (!date && c.f.date) date = c.f.date;
      });
      var timeTaken = false;
      cells.forEach(function (c) {
        if (c.col === scheduleNoteCol) return;
        var isTime = !timeTaken && time && (c.whole === time || c.f.time === time);
        if (isTime) timeTaken = true;
        if (c.whole && c.whole === time && isTime) return;
        if (c.f.rest) rest.push(c.f.rest);
      });
      if (!time && !status) return;
      if (rest.length && rest.every(isHeaderRest)) rest = [];
      // Without a time, a line must be short ("Full - Jamie P"). A sentence that mentions "closed" is not a slot.
      if (!time && rest.join(' ').split(/\s+/).filter(Boolean).length > 3) return;
      addSlot(days, day, finishSlot({ status: status, date: date }, time, rest, scheduleNotes));
      slots++;
    });

    return { days: days, slots: slots, dayCount: Object.keys(days).length };
  }

  /** Text (a PDF, a Word document, a pasted schedule) as rows of cells. Tabs, pipes and wide gaps separate cells. */
  function matrixFromText(text) {
    return String(text || '').replace(/\r\n?/g, '\n').split('\n').map(function (line) {
      if (/\t/.test(line)) return line.split('\t');
      if (/\s\|\s|^\s*\|/.test(line)) return line.split('|');
      if (/ {3,}/.test(line)) return line.split(/ {3,}/);
      return [line];
    });
  }

  function scheduleFromText(text) {
    return scheduleFromMatrix(matrixFromText(text));
  }

  /** The best schedule among several sheets: a sheet called "current" wins when it has anything, else the fullest one. */
  function bestSheet(sheets) {
    var results = (sheets || []).map(function (sheet) {
      var parsed = scheduleFromMatrix(sheet.rows);
      parsed.name = sheet.name || '';
      return parsed;
    }).filter(function (r) { return r.slots > 0; });
    if (!results.length) return null;
    var current = results.filter(function (r) { return /current/i.test(r.name); })[0];
    if (current) return current;
    results.sort(function (a, b) { return b.slots - a.slots; });
    return results[0];
  }

  /** True when the result is clearly a schedule and not a stray time mentioned in a client list. */
  function looksLikeSchedule(result) {
    return !!result && result.slots >= 3 && result.dayCount >= 1;
  }

  return {
    DAY_KEYS: DAY_KEYS,
    DAY_LABELS: DAY_LABELS,
    canonicalStatus: canonicalStatus,
    findTime: findTime,
    facets: facets,
    scheduleFromMatrix: scheduleFromMatrix,
    scheduleFromText: scheduleFromText,
    matrixFromText: matrixFromText,
    bestSheet: bestSheet,
    looksLikeSchedule: looksLikeSchedule
  };
}));
