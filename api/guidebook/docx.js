/**
 * Reads headings, paragraphs, lists, tables, and links out of a .docx file.
 * The dashboard theme renders that text. The Word editor is not shown.
 */
const zlib = require('zlib');

function xmlText(value) {
  return String(value || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, function (_m, n) { return String.fromCharCode(Number(n)); })
    .replace(/&amp;/g, '&');
}

function escapeHtml(value) {
  return xmlText(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function unzip(buffer) {
  const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  const files = {};
  let i = 0;
  while (i + 30 <= buf.length && buf.readUInt32LE(i) === 0x04034b50) {
    const method = buf.readUInt16LE(i + 8);
    const compSize = buf.readUInt32LE(i + 18);
    const nameLen = buf.readUInt16LE(i + 26);
    const extraLen = buf.readUInt16LE(i + 28);
    const name = buf.slice(i + 30, i + 30 + nameLen).toString('utf8');
    const start = i + 30 + nameLen + extraLen;
    const compressed = buf.slice(start, start + compSize);
    let data = compressed;
    if (method === 8) data = zlib.inflateRawSync(compressed);
    else if (method !== 0) data = Buffer.alloc(0);
    files[name] = data;
    i = start + compSize;
  }
  return files;
}

function attr(tag, name) {
  const match = String(tag || '').match(new RegExp('\\b' + name + '="([^"]*)"', 'i'));
  return match ? match[1] : '';
}

function relationships(xml) {
  const links = {};
  const images = {};
  const re = /<Relationship\b[^>]*\/?>/g;
  let match;
  while ((match = re.exec(xml || ''))) {
    const tag = match[0];
    const id = attr(tag, 'Id');
    const target = xmlText(attr(tag, 'Target'));
    if (!id || !target) continue;
    if (/hyperlink/i.test(tag)) links[id] = target;
    else if (/\/image\b/i.test(attr(tag, 'Type'))) images[id] = target;
  }
  return { links: links, images: images };
}

function mediaDataUri(files, target) {
  const clean = String(target || '').replace(/^\/+/, '');
  const base = clean.split('/').pop();
  const key = [clean, 'word/' + clean, base].concat(Object.keys(files)).find(function (name) {
    return files[name] && (name === clean || name === 'word/' + clean || name === base || name.endsWith('/' + base));
  });
  const data = key ? files[key] : null;
  if (!data || !data.length) return '';
  const ext = String(key.split('.').pop() || 'png').toLowerCase();
  const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : 'image/png';
  return 'data:' + mime + ';base64,' + data.toString('base64');
}

function runHtml(run) {
  let text = '';
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>/g;
  let match;
  while ((match = re.exec(run))) {
    if (match[1] != null) text += escapeHtml(match[1]);
    else if (/<w:tab/.test(match[0])) text += ' ';
    else text += '<br>';
  }
  if (!text) return '';
  const bold = /<w:b\b/.test(run) && !/w:val="(?:0|false)"/.test(run.slice(0, run.indexOf('>') + 80));
  const italic = /<w:i\b/.test(run) && !/<w:i\b[^>]*w:val="(?:0|false)"/.test(run);
  if (bold) text = '<strong>' + text + '</strong>';
  if (italic) text = '<em>' + text + '</em>';
  return text;
}

function imageTags(fragment, images) {
  let html = '';
  const re = /r:embed="([^"]+)"/g;
  let match;
  while ((match = re.exec(fragment || ''))) {
    if (images[match[1]]) html += '<img src="' + images[match[1]] + '" alt="">';
  }
  return html;
}

function paragraphParts(paragraph, links, images) {
  let html = '';
  const re = /<w:hyperlink\b[^>]*r:id="([^"]+)"[^>]*>([\s\S]*?)<\/w:hyperlink>|<w:r\b[\s\S]*?<\/w:r>/gi;
  let match;
  while ((match = re.exec(paragraph))) {
    if (match[1]) {
      const href = links[match[1]] || '';
      const label = (runHtml(match[2]) || escapeHtml(href)) + imageTags(match[2], images);
      if (/^https?:|^mailto:/i.test(href)) html += '<a href="' + escapeHtml(href) + '">' + label + '</a>';
      else html += label;
    } else {
      html += runHtml(match[0]) + imageTags(match[0], images);
    }
  }
  return html;
}

function headingTag(paragraph) {
  const style = ((paragraph.match(/<w:pStyle\b[^>]*w:val="([^"]+)"/) || [])[1] || '')
    .toLowerCase()
    .replace(/[\s_-]/g, '');
  if (style === 'title' || style === 'heading1') return 'h2';
  if (style === 'heading2') return 'h3';
  if (style === 'heading3' || style === 'heading4') return 'h4';
  const level = paragraph.match(/<w:outlineLvl\b[^>]*w:val="(\d+)"/);
  if (level) {
    const depth = Number(level[1]);
    if (depth === 0) return 'h2';
    if (depth === 1) return 'h3';
    if (depth <= 3) return 'h4';
  }
  return '';
}

function blocksFromXml(documentXml, relXml, files) {
  const rels = relationships(relXml);
  const links = rels.links;
  const images = {};
  Object.keys(rels.images).forEach(function (id) {
    const uri = mediaDataUri(files || {}, rels.images[id]);
    if (uri) images[id] = uri;
  });
  const body = (documentXml.match(/<w:body\b[\s\S]*<\/w:body>/) || [documentXml])[0];
  const chunks = body.match(/<w:tbl\b[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const html = [];
  let list = [];
  function flushList() {
    if (!list.length) return;
    html.push('<ul>' + list.map(function (item) { return '<li>' + item + '</li>'; }).join('') + '</ul>');
    list = [];
  }
  chunks.forEach(function (chunk) {
    if (chunk.indexOf('<w:tbl') === 0) {
      flushList();
      const rows = chunk.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || [];
      const rendered = rows.map(function (row, index) {
        const cells = row.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || [];
        const tag = index === 0 ? 'th' : 'td';
        const inner = cells.map(function (cell) {
          const paras = cell.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
          return '<' + tag + '>' + paras.map(function (p) { return paragraphParts(p, links, images); }).filter(Boolean).join('<br>') + '</' + tag + '>';
        }).join('');
        return '<tr>' + inner + '</tr>';
      }).join('');
      if (rendered) html.push('<table>' + rendered + '</table>');
      return;
    }
    const text = paragraphParts(chunk, links, images);
    if (!text) return;
    const tag = headingTag(chunk);
    if (!tag && /<w:numPr\b/.test(chunk)) {
      list.push(text);
      return;
    }
    flushList();
    html.push(tag ? '<' + tag + '>' + text + '</' + tag + '>' : '<p>' + text + '</p>');
  });
  flushList();
  return html.join('');
}

function docxToHtml(buffer) {
  const files = unzip(buffer);
  const documentXml = files['word/document.xml'] ? files['word/document.xml'].toString('utf8') : '';
  if (!documentXml) throw new Error('The guidebook file could not be read.');
  const relXml = files['word/_rels/document.xml.rels'] ? files['word/_rels/document.xml.rels'].toString('utf8') : '';
  return blocksFromXml(documentXml, relXml, files);
}

module.exports = { docxToHtml };
