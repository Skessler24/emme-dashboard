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
  return decodeBasic(String(html || '').replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
}

function promotePlainHeadings(html) {
  return String(html || '').replace(/<p>\s*<strong>([^<]{2,90})<\/strong>\s*<\/p>/gi, function (_full, inner) {
    const title = textOf(inner);
    if (/^\d+[a-z]\.\s+\S/.test(title)) return '<h3>' + inner + '</h3>';
    if (/^\d+\.\s+\S/.test(title)) return '<h2>' + inner + '</h2>';
    return '<p><strong>' + inner + '</strong></p>';
  });
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
