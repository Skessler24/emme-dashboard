/**
 * Reads client contact details out of loose text: a PDF printed from a web page,
 * a pasted email, a Word document with paragraphs, a list with no rows and columns.
 *
 * It does not need headings or a table. It finds email addresses, phone numbers,
 * and birth dates wherever they are (repairing the way PDFs split them apart),
 * and takes the name from the line just above or just before them.
 *
 * Loaded by app5.html in the browser and by the API tests in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.EmmeClientImport = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Words that appear on the schedule page and in forms. A line that holds one of them is not a person's name.
  var STOP = {};
  ('monday tuesday wednesday thursday friday saturday sunday mondays tuesdays wednesdays thursdays fridays saturdays sundays ' +
   'slot slots rename color colors delete time status notes note last visit board schedule caseload dashboard spreadsheet pdf share ' +
   'import saved sharepoint retry current regular spot clients client open full email phone dob page print export emme add day days ' +
   'none yes no sessions session tasks name names total signature date birth parent parents guardian coordinator').split(' ').forEach(function (word) { STOP[word] = true; });

  var PARTICLES = { de: 1, del: 1, la: 1, le: 1, van: 1, von: 1, der: 1, da: 1, di: 1, st: 1, bin: 1, al: 1, el: 1 };

  var STAFF_DOMAIN = /@emmeco\.org$/i;

  // ---------------------------------------------------------------------------
  // Putting back together what a PDF pulled apart
  // ---------------------------------------------------------------------------

  function twoDigits(value) {
    value = String(value);
    return value.length < 2 ? '0' + value : value;
  }

  /** Rebuilds a page of text from PDF text pieces: lines by height, pieces left to right, spaces only where there is a gap. */
  function pdfItemsToText(items) {
    var pieces = (items || []).filter(function (item) {
      return item && typeof item.str === 'string' && item.str.length && item.transform;
    }).map(function (item) {
      return {
        str: item.str,
        x: Number(item.transform[4]) || 0,
        y: Number(item.transform[5]) || 0,
        width: Number(item.width) || 0,
        height: Math.abs(Number(item.height) || Number(item.transform[3]) || 10),
      };
    });
    var lines = [];
    pieces.forEach(function (piece) {
      var line = null;
      for (var i = lines.length - 1; i >= 0 && i >= lines.length - 6; i--) {
        if (Math.abs(lines[i].y - piece.y) <= Math.max(2, piece.height * 0.4)) { line = lines[i]; break; }
      }
      if (!line) { line = { y: piece.y, pieces: [] }; lines.push(line); }
      line.pieces.push(piece);
    });
    lines.sort(function (a, b) { return b.y - a.y; });
    return lines.map(function (line) {
      line.pieces.sort(function (a, b) { return a.x - b.x; });
      var out = '';
      var end = null;
      var lastHeight = 10;
      line.pieces.forEach(function (piece) {
        if (out && end !== null) {
          var gap = piece.x - end;
          var space = /\s$/.test(out) || /^\s/.test(piece.str);
          // A wide gap is a table column. Keep it as a tab so cells stay apart.
          if (gap > Math.max(12, lastHeight * 2.2)) out = out.replace(/ +$/, '') + '\t';
          else if (!space && gap > Math.max(0.5, lastHeight * 0.12)) out += ' ';
        }
        out += piece.str;
        end = piece.x + piece.width;
        lastHeight = piece.height || lastHeight;
      });
      return out;
    }).join('\n');
  }

  function normalizeDate(raw) {
    var m = String(raw || '').match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (!m) return '';
    var year = m[3];
    if (year.length === 2) year = (Number(year) <= 30 ? '20' : '19') + year;
    if (year.length === 3) return '';
    return twoDigits(m[1]) + '/' + twoDigits(m[2]) + '/' + year;
  }

  function repairEmails(text) {
    var re = /((?:[A-Za-z0-9_%+\-]+[ \t]*\.[ \t]*)*[A-Za-z0-9_%+\-]+(?:[ \t]+\d{1,6})?)[ \t]*@[ \t]*([A-Za-z0-9\-]+(?:[ \t\n]*\.[ \t\n]*[A-Za-z0-9\-]+)+)/g;
    return text.replace(re, function (all, local, domain) {
      var labels = domain.split(/[ \t\n]*\.[ \t\n]*/);
      var spill = '';
      // "name@site.com. Next sentence" must not pull the next word into the address.
      while (labels.length > 2 && /^[A-Z]/.test(labels[labels.length - 1])) spill = '. ' + labels.pop() + spill;
      return local.replace(/[ \t]+/g, '') + '@' + labels.join('.') + spill;
    });
  }

  function repairPhones(text) {
    var one = /(?:\+?1[\s.\-]*)?(?:\((\d{3})\)[\s.\-]*(\d{3})[\s.\-]*(\d{4})|\b(\d{3})[ .\-]+(\d{3})[ .\-]+(\d{4}))\b/g;
    return text.replace(one, function (all, a1, b1, c1, a2, b2, c2) {
      return '(' + (a1 || a2) + ') ' + (b1 || b2) + '-' + (c1 || c2);
    });
  }

  /** Cleans up the text of a document so contact details are whole again. */
  function repairText(text) {
    var t = String(text == null ? '' : text)
      .replace(/\r\n?/g, '\n')
      .replace(/[\u00a0\u2007\u202f]/g, ' ')
      .replace(/[\u200b\u200c\u200d\ufeff]/g, '');
    t = repairEmails(t);
    t = repairPhones(t);
    // Spaces collapse, but tabs stay: they mark the cells of a table.
    return t.split('\n').map(function (line) {
      return line.replace(/ *\t[ \t]*/g, '\t').replace(/ {2,}/g, ' ').replace(/^[ \t]+|[ \t]+$/g, '');
    }).join('\n');
  }

  // ---------------------------------------------------------------------------
  // Names
  // ---------------------------------------------------------------------------

  function stripEdges(line) {
    return String(line || '').replace(/^[\s·•*\-–—|,;:]+/, '').replace(/[\s·•*\-–—|,;:]+$/, '').trim();
  }

  function tidyName(value) {
    var v = String(value || '').replace(/\s+/g, ' ').replace(/\s*[-–—]\s*/g, '-').trim();
    var flipped = v.match(/^([A-Za-z][A-Za-z'’.\-]*),\s*([A-Za-z][A-Za-z'’.\-]*(?:\s+[A-Za-z][A-Za-z'’.\-]*)*)$/);
    if (flipped) v = flipped[2] + ' ' + flipped[1];
    return v;
  }

  function isNameLine(line) {
    var v = stripEdges(line);
    if (!v || v.length < 2 || v.length > 60) return false;
    if (/[0-9@:()\/\\|+&=<>_#%$*✕←→]/.test(v)) return false;
    if (!/^[A-Za-z][A-Za-z'’.]*(?:(?:\s+[-–—]\s+|[-–—]|\s+)[A-Za-z][A-Za-z'’.]*)*$/.test(v)) return false;
    var words = v.split(/\s*[-–—]\s*|\s+/).filter(Boolean);
    if (words.length < 1 || words.length > 5) return false;
    for (var i = 0; i < words.length; i++) {
      var word = words[i].replace(/[.'’]+$/, '');
      var lower = word.toLowerCase();
      if (STOP[lower]) return false;
      if (PARTICLES[lower] && i > 0) continue;
      if (!/^[A-Z]/.test(word)) return false;
    }
    return true;
  }

  // Words that start a sentence ("Please add Sam Ortiz", "Also Sam Ortiz") and are not part of the name.
  var LEADING = {};
  ('also and please add new hi hello dear thanks thank contact call email reach regarding re fwd from to cc for the my our his her their').split(' ').forEach(function (word) { LEADING[word] = true; });

  function dropLeading(name) {
    var words = String(name || '').split(/\s+/).filter(Boolean);
    while (words.length > 1 && LEADING[words[0].toLowerCase()]) words.shift();
    return words.join(' ');
  }

  /** Finds a person's name inside a sentence: the last run of two or more capitalized words. */
  function nameFromSentence(text) {
    var runs = String(text || '').match(/[A-Z][A-Za-z'’\-]*(?:[ ]+[A-Z][A-Za-z'’\-]*)+/g) || [];
    for (var i = runs.length - 1; i >= 0; i--) {
      var candidate = dropLeading(runs[i]);
      if (candidate.indexOf(' ') > 0 && isNameLine(candidate)) return candidate;
    }
    return '';
  }

  // ---------------------------------------------------------------------------
  // Reading one line
  // ---------------------------------------------------------------------------

  var EMAIL = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}/g;
  var PHONE = /\(\d{3}\) \d{3}-\d{4}/g;
  var DOB = /\b(?:D\.?O\.?B\.?|date of birth|birth ?date|birthday|born)\b\s*[:\-]?\s*(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})/i;
  var WEEKDAY = /\b(?:mon|tues?|wed(?:nes)?|thu(?:rs?)?|fri|sat(?:ur)?|sun)(?:day)?s?\b/i;

  function scanLine(line) {
    var found = { emails: [], phones: [], dob: '', lead: '', rest: '' };
    var starts = [];
    var m;
    EMAIL.lastIndex = 0;
    while ((m = EMAIL.exec(line))) { found.emails.push(m[0].toLowerCase()); starts.push(m.index); }
    PHONE.lastIndex = 0;
    while ((m = PHONE.exec(line))) { found.phones.push(m[0]); starts.push(m.index); }
    var d = line.match(DOB);
    if (d) { found.dob = normalizeDate(d[1]); starts.push(d.index); }
    found.any = !!(found.emails.length || found.phones.length || found.dob);
    if (found.any) {
      var first = Math.min.apply(null, starts);
      found.lead = stripEdges(line.slice(0, first));
      found.rest = stripEdges(line
        .replace(EMAIL, ' ').replace(PHONE, ' ').replace(DOB, ' ')
        .replace(/\b(?:email|e-mail|phone|cell|mobile|tel|telephone|dob)\b\s*:?/ig, ' ')
        .replace(/[·•|]/g, ' ').replace(/\s+/g, ' '));
    }
    return found;
  }

  // ---------------------------------------------------------------------------
  // Whole documents
  // ---------------------------------------------------------------------------

  function addSessions(record, text) {
    var t = stripEdges(text);
    if (!t || !WEEKDAY.test(t) || t.length > 120) return;
    if (/✕|←|→|\b(?:rename|slot|color|colour)\b/i.test(t)) return; // page buttons, not a session
    record.sessions = record.sessions ? (record.sessions.indexOf(t) === -1 ? record.sessions + '; ' + t : record.sessions) : t;
  }

  function fillRecord(record, found) {
    found.emails.forEach(function (email) {
      if (STAFF_DOMAIN.test(email)) return;
      if (!record.email) record.email = email;
    });
    found.phones.forEach(function (phone) {
      if (!record.phone) record.phone = phone;
      else if (record.phone.indexOf(phone) === -1) record.phone += ' / ' + phone;
    });
    if (found.dob && !record.dob) record.dob = found.dob;
    addSessions(record, found.rest);
  }

  function clientsFromFreeText(text) {
    // Table cells (split by tabs) are read one after another, left to right, like separate lines.
    var lines = [];
    repairText(text).split('\n').forEach(function (line) {
      line.split('\t').forEach(function (cell) {
        cell = cell.trim();
        if (cell) lines.push(cell);
      });
    });
    var records = [];
    var current = null;
    var pending = [];
    lines.forEach(function (line) {
      var found = scanLine(line);
      if (found.any) {
        var name = pending.length ? pending[pending.length - 1]
          : (isNameLine(found.lead) ? dropLeading(found.lead) : nameFromSentence(found.lead));
        if (current && !name) {
          fillRecord(current, found);
        } else {
          current = { name: tidyName(name) };
          records.push(current);
          fillRecord(current, found);
        }
        pending = [];
        return;
      }
      if (isNameLine(line)) {
        pending.push(stripEdges(line));
        if (pending.length > 3) pending.shift();
        return;
      }
      if (current && WEEKDAY.test(line)) addSessions(current, line);
    });

    var named = records.filter(function (record) { return record.name; });
    var kept = (named.length ? named : records).filter(function (record) { return record.email || record.phone || record.dob || record.name; });
    // Households share one email address and phone number, so only the same name counts as the same person.
    var seen = {};
    var out = [];
    kept.forEach(function (record) {
      var key = (record.name || '').toLowerCase() + '|' + (record.name ? '' : (record.email || record.phone || ''));
      if (seen[key]) {
        Object.keys(record).forEach(function (field) { if (!seen[key][field] && record[field]) seen[key][field] = record[field]; });
        return;
      }
      var clean = {};
      Object.keys(record).forEach(function (field) { if (record[field]) clean[field] = record[field]; });
      seen[key] = clean;
      out.push(clean);
    });
    return out.slice(0, 1000);
  }

  // ---------------------------------------------------------------------------
  // Other kinds of document
  // ---------------------------------------------------------------------------

  /** Turns a web page saved as a file into plain text. Table cells become tabs and rows become lines. */
  function htmlToText(html) {
    var t = String(html || '')
      .replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<!--[\s\S]*?-->/g, ' ')
      .replace(/<\/(td|th)>/gi, '\t')
      .replace(/<br\s*\/?>|<\/(p|div|tr|li|h[1-6]|section|article|table)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'")
      .replace(/&#(\d+);/g, function (all, code) { return String.fromCharCode(Number(code)); });
    return t.split('\n').map(function (line) { return line.replace(/[ \u00a0]+/g, ' ').replace(/ ?\t ?/g, '\t').replace(/\t+$/, '').trim(); }).join('\n');
  }

  /** Turns a Rich Text Format document into plain text. */
  function rtfToText(rtf) {
    var t = String(rtf || '');
    t = t.replace(/\{\\\*[^{}]*(?:\{[^{}]*\}[^{}]*)*\}/g, '');
    t = t.replace(/\\par[d]?\b ?/g, '\n').replace(/\\(?:line|tab)\b ?/g, function (m) { return /tab/.test(m) ? '\t' : '\n'; });
    t = t.replace(/\\'([0-9a-fA-F]{2})/g, function (all, hex) { return String.fromCharCode(parseInt(hex, 16)); });
    t = t.replace(/\\u(-?\d+)\??/g, function (all, code) { var n = Number(code); return String.fromCharCode(n < 0 ? n + 65536 : n); });
    t = t.replace(/\\[a-zA-Z]+-?\d* ?/g, '').replace(/\\([\\{}])/g, '$1').replace(/[{}]/g, '');
    return t.split('\n').map(function (line) { return line.trim(); }).join('\n');
  }

  /** Reads contact cards (.vcf files, such as an exported phone or email contact list). */
  function clientsFromVcard(text) {
    var unfolded = String(text || '').replace(/\r\n?/g, '\n').replace(/\n[ \t]/g, '');
    var cards = unfolded.split(/BEGIN:VCARD/i).slice(1);
    var out = [];
    cards.forEach(function (card) {
      var client = {};
      card.split('\n').forEach(function (raw) {
        var cut = raw.indexOf(':');
        if (cut < 0) return;
        var key = raw.slice(0, cut).split(';')[0].toUpperCase();
        var value = raw.slice(cut + 1).trim();
        if (!value) return;
        if (key === 'FN' && !client.name) client.name = tidyName(value);
        else if (key === 'N' && !client.name) {
          var parts = value.split(';');
          client.name = tidyName(((parts[1] || '') + ' ' + (parts[0] || '')).trim());
        } else if (key === 'TEL' && !client.phone) client.phone = repairPhones(value);
        else if (key === 'EMAIL' && !client.email) client.email = value.toLowerCase();
        else if (key === 'BDAY' && !client.dob) {
          var iso = value.match(/^(\d{4})-?(\d{2})-?(\d{2})/);
          client.dob = iso ? iso[2] + '/' + iso[3] + '/' + iso[1] : normalizeDate(value);
        }
      });
      if (client.name || client.email || client.phone) out.push(client);
    });
    return out;
  }

  return {
    clientsFromFreeText: clientsFromFreeText,
    clientsFromVcard: clientsFromVcard,
    htmlToText: htmlToText,
    rtfToText: rtfToText,
    pdfItemsToText: pdfItemsToText,
    repairText: repairText,
    isNameLine: isNameLine,
  };
}));
