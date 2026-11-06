/**
 * Turns the guidebook Word HTML into the sections the themed page shows.
 * The Word file stays the source. This only keeps headings, text, lists, tables, and links.
 */

const ALLOWED = {
  p: 1, h1: 1, h2: 1, h3: 1, h4: 1, ul: 1, ol: 1, li: 1,
  strong: 1, em: 1, b: 1, i: 1, u: 1, a: 1, br: 1,
  table: 1, thead: 1, tbody: 1, tr: 1, th: 1, td: 1,
  blockquote: 1, img: 1, sup: 1, sub: 1,
};

function decodeBasic(value) {
  return String(value || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"');
}

function sanitize(html) {
  let out = String(html || '');
  out = out.replace(/<\s*(script|style|iframe|object|embed|link|meta)[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, '');
  out = out.replace(/<\s*(script|style|iframe|object|embed|link|meta)[^>]*\/?\s*>/gi, '');
  out = out.replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '');
  out = out.replace(/\sstyle\s*=\s*("[^"]*"|'[^']*')/gi, '');
  out = out.replace(/\s(href|src)\s*=\s*("([^"]*)"|'([^']*)')/gi, function (_full, attr, _quoted, dq, sq) {
    const value = String(dq || sq || '').trim();
    const ok = /^(https?:|mailto:)/i.test(value) || /^data:image\/(png|jpeg|gif|webp);base64,/i.test(value);
    if (!ok) return '';
    return ' ' + attr + '="' + value.replace(/"/g, '%22') + '"';
  });
  out = out.replace(/<\/?([a-z0-9]+)([^>]*)>/gi, function (full, tag) {
    if (!ALLOWED[String(tag || '').toLowerCase()]) return '';
    return full;
  });
  out = out.replace(/<a\s+/gi, '<a target="_blank" rel="noopener noreferrer" ');
  return out.trim();
}

function textOf(html) {
  return decodeBasic(String(html || '')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<\/(p|div|li|tr|h[1-6])>/gi, ' ')
    .replace(/<[^>]+>/g, ''))
    .replace(/\s+/g, ' ')
    .trim();
}

function splitHtmlAt(html, count) {
  const src = String(html || '');
  let seen = 0;
  let i = 0;
  while (i < src.length && seen < count) {
    if (src.charAt(i) === '<') {
      const end = src.indexOf('>', i);
      if (end < 0) break;
      i = end + 1;
      continue;
    }
    if (src.charAt(i) === '&') {
      const end = src.indexOf(';', i);
      const entity = end < 0 ? src.charAt(i) : src.slice(i, end + 1);
      seen += decodeBasic(entity).length || 1;
      i = end < 0 ? i + 1 : end + 1;
      continue;
    }
    seen += 1;
    i += 1;
  }
  if (seen < count) return null;
  return { head: src.slice(0, i), rest: src.slice(i) };
}

function promotePlainHeadings(html) {
  return String(html || '').replace(/<p>([\s\S]*?)<\/p>/gi, function (_full, inner) {
    let title = textOf(inner);
    if (!/^\d+[a-z]?\.\s+\S/.test(title)) return '<p>' + inner + '</p>';
    let headHtml = inner;
    let restHtml = '';
    if (title.length > 140) {
      const cut = title.match(/^(\d+[a-z]?\.\s+[^:]{2,100}?):\s+/);
      if (!cut) return '<p>' + inner + '</p>';
      const parts = splitHtmlAt(inner, cut[1].length);
      if (!parts || !textOf(parts.rest)) return '<p>' + inner + '</p>';
      headHtml = parts.head;
      restHtml = parts.rest.replace(/^(?:\s|<br>)*:?\s*/i, '');
      title = textOf(headHtml);
    }
    const level = /^\d+[a-z]\./i.test(title) ? 'h3' : 'h2';
    const heading = '<' + level + '>' + headHtml + '</' + level + '>';
    return restHtml && textOf(restHtml) ? heading + '<p>' + restHtml + '</p>' : heading;
  });
}

function isWordContents(html) {
  const withoutList = String(html || '').replace(/<ul[\s\S]*<\/ul>/gi, ' ');
  const rest = textOf(withoutList);
  return /table of contents/i.test(textOf(html)) && rest.length < 140 && /guidebook|contents/i.test(rest);
}

function promoteNumberedSections(html) {
  return String(html || '').replace(/<h3>([\s\S]*?)<\/h3>/gi, function (full, inner) {
    const title = textOf(inner);
    if (/^\d+\.\s+\S/.test(title) && !/^\d+[a-z]\./i.test(title)) return '<h2>' + inner + '</h2>';
    return full;
  });
}

function attachSubsections(section, index) {
  const subsections = [];
  let count = 0;
  section.html = String(section.html || '').replace(/<h([34])([^>]*)>([\s\S]*?)<\/h\1>/gi, function (full, level, attrs, inner) {
    const title = textOf(inner);
    if (!title) return full;
    if (level === '4' && !/^\d/.test(title)) return full;
    count += 1;
    const id = 's' + (index + 1) + '-' + count;
    subsections.push({ title: title, id: id });
    return '<h' + level + ' id="' + id + '"' + attrs + '>' + inner + '</h' + level + '>';
  });
  section.subsections = subsections;
  return section;
}

function splitSections(html) {
  const cleaned = promoteNumberedSections(promotePlainHeadings(sanitize(html)));
  const chunks = cleaned.split(/(?=<h[12]\b)/i).map(function (part) { return part.trim(); }).filter(Boolean);
  const sections = [];
  let coverTitle = '';
  chunks.forEach(function (chunk) {
    const heading = chunk.match(/^<h([12])[^>]*>([\s\S]*?)<\/h\1>/i);
    if (!heading) {
      if (isWordContents(chunk)) {
        if (!coverTitle) coverTitle = 'Provider Guidebook';
        return;
      }
      if (textOf(chunk)) sections.push({ title: 'Overview', html: chunk });
      return;
    }
    const title = textOf(heading[2]) || 'Section';
    const body = chunk.slice(heading[0].length).trim();
    if (/guidebook/i.test(title) && !textOf(body) && !coverTitle) {
      coverTitle = title;
      return;
    }
    sections.push({ title: title, html: body });
  });
  if (!sections.length && cleaned) sections.push({ title: 'Guidebook', html: cleaned });
  sections.forEach(attachSubsections);
  return { coverTitle: coverTitle, sections: sections };
}

module.exports = { sanitize, splitSections, textOf };
