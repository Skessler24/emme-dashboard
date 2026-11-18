const test = require('node:test');
const assert = require('node:assert/strict');
const lib = require('../../client-import.js');

// All names, emails and phone numbers below are made up.

function piece(str, x, y) {
  return { str, transform: [1, 0, 0, 10, x, y], width: str.length * 5, height: 10 };
}

test('a printed schedule page (a table with no labels) gives full names and details', () => {
  const items = [
    // page chrome that must not be taken for clients
    piece('Monday', 60, 470), piece('+', 120, 470), piece('Slot', 140, 470), piece('Rename', 200, 470), piece('Color', 260, 470), piece('✕', 320, 470),
    piece('TIME', 90, 440), piece('STATUS', 184, 440), piece('NOTES', 313, 440), piece('LAST VISIT', 604, 440),
    // one client row: first name on top, full name, then details, then days
    piece('Pat', 313, 323),
    piece('4:30', 91, 299), piece('PM', 117, 299), piece('Open', 184, 298), piece('(2&4)', 228, 298),
    piece('Pat', 313, 300), piece('Rivera-Stone', 342, 300),
    piece('DOB', 313, 284), piece('01/15/2024', 334, 284), piece('·', 392, 284),
    piece('pat.rivera@', 400, 284), piece('example.com', 450, 284), piece('·', 505, 284), piece('(303) 555-0101', 515, 284),
    piece('Mondays', 313, 272), piece('4:00', 353, 272), piece('EOW', 374, 272),
    // a second client sharing the same family email
    piece('Lee', 313, 223),
    piece('4:30', 91, 200), piece('PM', 117, 200),
    piece('Lee', 313, 200), piece('Rivera-Stone', 342, 200),
    piece('DOB', 313, 184), piece('03/02/2022', 334, 184), piece('·', 392, 184),
    piece('pat.rivera@', 400, 184), piece('example.com', 450, 184), piece('·', 505, 184), piece('(303) 555-0101', 515, 184),
    piece('Fridays', 313, 172), piece('9:00', 353, 172),
  ];
  const clients = lib.clientsFromFreeText(lib.pdfItemsToText(items));
  assert.equal(clients.length, 2);
  assert.equal(clients[0].name, 'Pat Rivera-Stone');
  assert.equal(clients[0].dob, '01/15/2024');
  assert.equal(clients[0].email, 'pat.rivera@example.com');
  assert.equal(clients[0].phone, '(303) 555-0101');
  assert.match(clients[0].sessions, /Mondays/);
  assert.doesNotMatch(clients[0].sessions, /Rename|Slot|✕/);
  assert.equal(clients[1].name, 'Lee Rivera-Stone');
  assert.equal(clients[1].email, 'pat.rivera@example.com');
});

test('emails split by a PDF are put back together', () => {
  const text = [
    'Jordan Avery',
    'DOB 05/06/2023 · jordan . avery @ example . com · (303) 555 0102',
    'Casey Morgan',
    'DOB 07/08/2022 · casey14 @ example . com · 303.555.0103',
  ].join('\n');
  const clients = lib.clientsFromFreeText(text);
  assert.deepEqual(clients.map((c) => c.email), ['jordan.avery@example.com', 'casey14@example.com']);
  assert.deepEqual(clients.map((c) => c.phone), ['(303) 555-0102', '(303) 555-0103']);
});

test('phone numbers wrapped across lines are joined', () => {
  const clients = lib.clientsFromFreeText('Riley Quinn\nriley.q@example.com\n(303) 555-\n0104');
  assert.equal(clients.length, 1);
  assert.equal(clients[0].phone, '(303) 555-0104');
});

test('plain sentences with no rows and columns still work', () => {
  const text = 'Please add Taylor Brooks. You can reach the family at taylor.family@example.com or 720-555-0105.\n\nAlso Sam Ortiz, sam.o@example.com, 720 555 0106.';
  const clients = lib.clientsFromFreeText(text);
  assert.equal(clients.length, 2);
  assert.equal(clients[0].name, 'Taylor Brooks');
  assert.equal(clients[0].email, 'taylor.family@example.com');
  assert.equal(clients[0].phone, '(720) 555-0105');
  assert.equal(clients[1].name, 'Sam Ortiz');
});

test('Emme staff addresses are not treated as client emails', () => {
  const clients = lib.clientsFromFreeText('Morgan Lane\nDOB 01/01/2024 · staff.person@emmeco.org · (303) 555-0107');
  assert.equal(clients.length, 1);
  assert.equal(clients[0].email, undefined);
  assert.equal(clients[0].phone, '(303) 555-0107');
});

test('text with no clients in it finds nothing', () => {
  assert.deepEqual(lib.clientsFromFreeText('Meeting notes\nRemember to bring snacks.'), []);
});

test('contact files (.vcf) are read', () => {
  const vcf = [
    'BEGIN:VCARD', 'VERSION:3.0', 'N:Hart;Jamie;;;', 'FN:Jamie Hart', 'TEL;TYPE=CELL:303-555-0108', 'EMAIL:Jamie.Hart@example.com', 'BDAY:2023-04-05', 'END:VCARD',
    'BEGIN:VCARD', 'FN:Alex Wu', 'END:VCARD',
  ].join('\r\n');
  const clients = lib.clientsFromVcard(vcf);
  assert.equal(clients.length, 2);
  assert.deepEqual(clients[0], { name: 'Jamie Hart', phone: '(303) 555-0108', email: 'jamie.hart@example.com', dob: '04/05/2023' });
});

test('saved web pages and rich text become readable text', () => {
  const html = '<html><head><style>p{}</style></head><body><table><tr><td>Dana Fox</td><td>dana.fox@example.com</td><td>(303) 555-0109</td></tr></table></body></html>';
  const clients = lib.clientsFromFreeText(lib.htmlToText(html));
  assert.equal(clients.length, 1);
  assert.equal(clients[0].name, 'Dana Fox');
  assert.equal(clients[0].email, 'dana.fox@example.com');

  const rtf = '{\\rtf1\\ansi Eli Park\\par DOB 02/03/2024 \\par eli.park@example.com \\par (303) 555-0110\\par}';
  const fromRtf = lib.clientsFromFreeText(lib.rtfToText(rtf));
  assert.equal(fromRtf.length, 1);
  assert.equal(fromRtf[0].name, 'Eli Park');
  assert.equal(fromRtf[0].dob, '02/03/2024');
});
