import { esc, run, tab, para, textPara, emptyPara, labelPara, valuePara } from './ooxml.js';

// Bump this string whenever smartFormatParty's logic changes. Shown on screen
// so a stale cached copy is obvious instead of silently reproducing old bugs.
export const PARTY_FORMAT_BUILD = 'party-fmt-5 (2026-09-10, wrap fix + duplicate warning)';

export const COURT_TITLE =
  'IN THE COURT OF THE PRL. SENIOR CIVIL JUDGE, BENGALURU RURAL DISTRICT, BENGALURU.';
const JUDGE_NAME = 'Abdul Saleem';
export const DISPOSAL_PREFIX =
  'This Petition coming on this day for final disposal before Sri. Abdul Saleem B.A. (LAW) L.L.B. N. Prl. Senior Civil Judge, Bengaluru Rural District, Bengaluru, in the presence of ';

export const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

function parseIsoDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || '').trim());
  if (!m) return null;
  return { year: Number(m[1]), month: Number(m[2]), day: Number(m[3]) };
}

/** `03-07-2026` */
export function formatDMY(iso) {
  const d = parseIsoDate(iso);
  if (!d) return '';
  return `${String(d.day).padStart(2, '0')}-${String(d.month).padStart(2, '0')}-${d.year}`;
}

export function ordinalSuffix(day) {
  if (day % 100 >= 11 && day % 100 <= 13) return 'th';
  return { 1: 'st', 2: 'nd', 3: 'rd' }[day % 10] || 'th';
}

/** Runs for `25`+superscript `th`+` August 2026.` */
export function ordinalDateRuns(iso, trailing) {
  const d = parseIsoDate(iso);
  if (!d) return [];
  return [
    run(String(d.day)),
    run(ordinalSuffix(d.day), { superscript: true }),
    run(` ${MONTHS[d.month - 1]} ${d.year}${trailing || ''}`),
  ];
}

/**
 * Pasted text carries the indentation of the document it was copied from, and a
 * tab inside a run is a real tab in Word — it would shift the whole line off the
 * value column. Strip that whitespace before it reaches the document.
 */
export function cleanText(text) {
  return String(text == null ? '' : text)
    .replace(/[\t\u00a0\u200b]+/g, ' ')
    .trim();
}

/** Split a pasted block into cleaned lines, dropping leading/trailing blanks. */
export function toLines(text) {
  const lines = String(text == null ? '' : text)
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map(cleanText);
  while (lines.length && lines[lines.length - 1] === '') lines.pop();
  while (lines.length && lines[0] === '') lines.shift();
  return lines;
}

// Matches "1.", "2.", "1(a)." etc. at the start of a new party/defendant entry.
const PARTY_START = /^\d+\s*(\([a-z0-9]+\))?\s*[.)]/i;

// Words that always begin their own detail line in a cause-title address, so a
// line starting with one of these is never merged into the line before it.
const DETAIL_START =
  /^(s\/o|d\/o|w\/o|c\/o|h\/o|daughter|son|wife|husband|aged|age|major|minor|r\/at|r\/o|residing|resident|rep\.?\s*by|represented|since|presently|permanent|through|occupation)\b/i;

function endsWithPunctuation(line) {
  return /[,.;:]$/.test(line);
}

// True for a fragment with no lowercase letters ("INDIA,", "BAGALUR-562149"). A
// name/title split by a photo scan reliably stays all-caps across the split,
// which is what tells it apart from the next real address line (mixed case).
function isShouty(line) {
  return /[A-Z]/.test(line) && !/[a-z]/.test(line);
}

/**
 * Reconstructs a party/defendant block pasted from a photo (Google Lens, phone
 * screenshots, etc.), which tends to add blank lines between every visual line
 * and to wrap a long name across two lines. Two passes, both conservative so a
 * cleanly pasted block is left untouched:
 *  1. Drop blank lines that are just scan noise, keeping exactly one as a
 *     separator when it precedes a new numbered party ("2. Rama,").
 *  2. Re-join a party's header line with the next line when the header lacks
 *     ending punctuation and the next line continues the same all-caps name
 *     ("1. Ms. UNION BANK OF" + "INDIA,") - names in these documents are
 *     conventionally capitals, so this never touches a genuine next detail
 *     like "Konappa Agrahara" or "Daughter of Venkatappa", which are mixed case.
 */
export function smartFormatParty(text) {
  const rawLines = String(text == null ? '' : text)
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.replace(/[ \t\u00a0\u200b]+/g, ' ').trim());

  const collapsed = [];
  for (let i = 0; i < rawLines.length; i += 1) {
    const line = rawLines[i];
    if (line !== '') { collapsed.push(line); continue; }
    let j = i;
    while (j < rawLines.length && rawLines[j] === '') j += 1;
    const next = rawLines[j];
    if (next && PARTY_START.test(next) && collapsed.length && collapsed[collapsed.length - 1] !== '') {
      collapsed.push('');
    }
  }
  while (collapsed.length && collapsed[0] === '') collapsed.shift();
  while (collapsed.length && collapsed[collapsed.length - 1] === '') collapsed.pop();

  const out = [];
  for (let i = 0; i < collapsed.length; i += 1) {
    let line = collapsed[i];
    const isHeader = out.length === 0 || PARTY_START.test(line);
    if (isHeader) {
      while (!endsWithPunctuation(line)) {
        const next = collapsed[i + 1];
        if (!next || PARTY_START.test(next) || DETAIL_START.test(next) || !isShouty(next)) break;
        line = `${line} ${next}`;
        i += 1;
      }
    }
    out.push(line);
  }
  return out;
}

export function documentOpen() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
    'xmlns:v="urn:schemas-microsoft-com:vml" ' +
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
    'mc:Ignorable="w14 w15 wp14"><w:body>'
  );
}

const SECT_PR =
  '<w:sectPr>' +
  '<w:headerReference w:type="default" r:id="rId8"/>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="2267" w:right="1440" w:bottom="1440" w:left="2160" w:header="1728" w:footer="0" w:gutter="0"/>' +
  '<w:cols w:space="720"/>' +
  '<w:formProt w:val="0"/>' +
  '<w:docGrid w:linePitch="360" w:charSpace="-2049"/>' +
  '</w:sectPr>';

export function documentClose() {
  return SECT_PR + '</w:body></w:document>';
}

/** `  DECREE` + centred underlined court title + underlined case number. */
export function headingBlock(caseTitle, decreeWord) {
  return [
    emptyPara(),
    para([run('  '), run(decreeWord || 'DECREE', { bold: true })]),
    textPara(COURT_TITLE, { jc: 'center', runOpts: { bold: true, underline: true } }),
    textPara(caseTitle, {
      ind: { left: 2160, firstLine: 720 },
      runOpts: { bold: true, underline: true },
    }),
  ].join('');
}

/**
 * `Label : first line` followed by one paragraph per remaining pasted line,
 * all aligned under the value column.
 */
export function partyBlock(label, text) {
  const lines = smartFormatParty(text);
  if (!lines.length) lines.push('Nil');
  const out = [labelPara(label, [run(lines[0])])];
  for (let i = 1; i < lines.length; i++) out.push(valuePara(lines[i]));
  return out.join('');
}

/** The `V/S` separator between the two party blocks. */
export function versusBlock() {
  return (
    valuePara('') +
    textPara('V/S', { ind: { left: 2880, firstLine: 720 }, jc: 'both' }) +
    emptyPara()
  );
}

/** Judge signature block used at the end of every section. */
export function signatureBlock() {
  const sixTabs = [tab({ bold: true }), tab({ bold: true }), tab({ bold: true }),
    tab({ bold: true }), tab({ bold: true }), tab({ bold: true })];
  return [
    emptyPara({ after: 0, ind: { firstLine: 720 }, jc: 'both' }),
    para(sixTabs, { after: 0, ind: { firstLine: 720 }, jc: 'both' }),
    para(
      [
        run('                                                                   (', { bold: true }),
        run(`${JUDGE_NAME})`),
      ],
      { after: 0, jc: 'both' }
    ),
    para([tab(), tab(), tab(), tab(), tab(), tab(), run('Prl. Senior Civil Judge ')]),
    para([tab(), tab(), tab(), tab(), tab(), run('Bengaluru Rural District, Bengaluru.')]),
  ].join('');
}

/** Centred, bold, underlined `ORDER` heading followed by the typed order text. */
export function orderBlock(orderText) {
  const paras = toLines(orderText).map((line) =>
    line === ''
      ? emptyPara({ after: 0, ind: { firstLine: 720 }, jc: 'both', runOpts: { bold: true } })
      : textPara(line, {
          after: 0,
          ind: { firstLine: 720 },
          jc: 'both',
          runOpts: { bold: true },
        })
  );
  return (
    textPara('ORDER', { jc: 'center', runOpts: { bold: true, underline: true } }) +
    paras.join('')
  );
}

/** `Given under my hand and seal of the court, Dated this the 25th August 2026.` */
export function givenUnderBlock(iso, trailing) {
  return para(
    [run('Given under my hand and seal of the court, Dated this the ')].concat(
      ordinalDateRuns(iso, trailing == null ? '.' : trailing)
    ),
    { after: 0, ind: { firstLine: 720 }, jc: 'both' }
  );
}

/** The right-aligned running header carrying the page number and case number. */
export function headerXml(caseTitle) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:hdr xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">' +
    '<w:p><w:pPr><w:pStyle w:val="Header"/>' +
    '<w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="left" w:pos="5370"/></w:tabs>' +
    '</w:pPr>' +
    '<w:r><w:tab/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText>PAGE</w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:tab/></w:r></w:p>' +
    '<w:p><w:pPr><w:spacing w:after="30"/><w:jc w:val="right"/>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/><w:lang w:val="en-US"/></w:rPr></w:pPr>' +
    '<w:r><w:tab/></w:r><w:r><w:tab/></w:r><w:r><w:tab/></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' +
    '<w:sz w:val="32"/><w:szCs w:val="32"/><w:lang w:val="en-US"/></w:rPr>' +
    '<w:t xml:space="preserve">' + esc(caseTitle) + '</w:t></w:r></w:p></w:hdr>'
  );
}
