import { run, textPara, emptyPara, labelPara } from './ooxml.js';
import * as C from './common.js';

const COST_ROWS = [
  { label: 'Court fee paid on plaint', field: 'courtFee' },
  { label: 'Stamp for power' },
  { label: 'Stamp for Exhibits' },
  { label: 'Fee certificate not filed' },
  { label: 'Subsistence for witnesses' },
  { label: 'Commissioner\u2019s Fee' },
  { label: 'Process Fee', field: 'processFee' },
];

const CELL_WIDTHS = [3860, 2157, 2279];
const CELL_BORDERS =
  '<w:tcBorders>' +
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
  '</w:tcBorders>';

export function caseTitle(caseNumber) {
  const n = String(caseNumber || '').trim();
  return /o\.?s\.?\s*no/i.test(n) ? n : `O.S.NO.${n}`;
}

/** Accepts `61795-00`, `61795.00`, `61,795` and returns the amount in paise. */
export function parseAmount(value) {
  const raw = String(value == null ? '' : value).trim().replace(/,/g, '');
  if (raw === '' || raw === '-') return null;
  const m = /^(\d+)(?:[-.](\d{1,2}))?$/.exec(raw);
  if (!m) return null;
  const paise = m[2] ? Number(m[2].padEnd(2, '0')) : 0;
  return Number(m[1]) * 100 + paise;
}

/** `6181500` -> `61815-00` */
export function formatAmount(paise) {
  if (paise == null) return '-';
  return `${Math.floor(paise / 100)}-${String(paise % 100).padStart(2, '0')}`;
}

function cell(index, text, opts = {}) {
  const body = text === ''
    ? emptyPara({ after: 0, jc: opts.jc || 'center', runOpts: { bold: !!opts.bold } })
    : textPara(text, { after: 0, jc: opts.jc || 'center', runOpts: { bold: !!opts.bold } });
  return (
    '<w:tc><w:tcPr>' +
    `<w:tcW w:w="${CELL_WIDTHS[index]}" w:type="dxa"/>` +
    CELL_BORDERS +
    '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>' +
    '<w:tcMar><w:left w:w="83" w:type="dxa"/></w:tcMar>' +
    '</w:tcPr>' + body + '</w:tc>'
  );
}

function row(cells, height) {
  return `<w:tr><w:trPr><w:trHeight w:val="${height}"/></w:trPr>${cells.join('')}</w:tr>`;
}

function costTable(data) {
  const rows = [
    row([cell(0, ''), cell(1, 'By the plaintiff'), cell(2, 'By the defendant')], 522),
  ];
  let total = null;
  for (const spec of COST_ROWS) {
    const amount = spec.field ? parseAmount(data[spec.field]) : null;
    if (amount != null) total = (total || 0) + amount;
    rows.push(
      row([cell(0, spec.label, { jc: 'left' }), cell(1, formatAmount(amount)), cell(2, '-')], 278)
    );
  }
  rows.push(
    row(
      [
        cell(0, 'Total', { bold: true }),
        cell(1, formatAmount(total), { bold: true }),
        cell(2, '-', { bold: true }),
      ],
      278
    )
  );

  return (
    '<w:tbl><w:tblPr>' +
    '<w:tblW w:w="0" w:type="auto"/>' +
    '<w:tblInd w:w="-25" w:type="dxa"/>' +
    '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="00000A"/>' +
    '</w:tblBorders>' +
    '<w:tblCellMar><w:left w:w="83" w:type="dxa"/></w:tblCellMar>' +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
    '</w:tblPr>' +
    `<w:tblGrid>${CELL_WIDTHS.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    rows.join('') +
    '</w:tbl>'
  );
}

const BOUNDARY_LINE = /^\s*(east|west|north|south)\s*by\s*:/i;

function scheduleBlock(text) {
  const lines = C.toLines(text);
  if (!lines.length) return '';
  const body = lines.map((line) => {
    if (line === '') return emptyPara();
    if (BOUNDARY_LINE.test(line)) return textPara(line, { ind: { left: 1440 } });
    return textPara(line, { ind: { firstLine: 720 }, jc: 'both' });
  });
  return (
    textPara('SCHEDULE ', {
      ind: { firstLine: 720 },
      jc: 'center',
      runOpts: { bold: true, underline: true },
    }) + body.join('')
  );
}

function claimBlock(label, text) {
  const lines = C.toLines(text);
  if (!lines.length) return labelPara(label, [run('')]);
  const out = [labelPara(label, [run(lines[0])])];
  for (let i = 1; i < lines.length; i++) {
    out.push(lines[i] === '' ? emptyPara({ jc: 'both' }) : textPara(lines[i], { jc: 'both' }));
  }
  return out.join('');
}

/** Joins one or more `{advocate, numbers}` entries as "{name} Advocate for {role} No.{n}, ...". */
function advocatePhrase(entries, role) {
  return entries
    .map((e, i) => ({ advocate: C.cleanText(e.advocate), numbers: C.cleanText(e.numbers), i }))
    .filter((e) => e.i === 0 || e.advocate)
    .map((e) => `${e.advocate} Advocate for ${role}` + (e.numbers ? ` No.${e.numbers}` : ''))
    .join(', ');
}

export function disposalSentence(data) {
  const exNo = C.cleanText(data.exparteNumbers);
  const plaintiffEntries = [
    { advocate: data.plaintiffAdvocate, numbers: data.plaintiffNumbers },
    ...(data.plaintiffAdvocatesExtra || []),
  ];
  const defendantEntries = [
    { advocate: data.defendantAdvocate, numbers: data.defendantNumbers },
    ...(data.defendantAdvocatesExtra || []),
  ];
  let s = C.DISPOSAL_PREFIX + advocatePhrase(plaintiffEntries, 'Plaintiff');
  s += ` and ${advocatePhrase(defendantEntries, 'Defendant')}`;
  if (data.exparte) s += `, Defendant${exNo ? ` No. ${exNo}` : ''}  placed Exparte`;
  return s + '.';
}

export function buildDocument(data) {
  const title = caseTitle(data.caseNumber);
  const parts = [C.documentOpen(), C.headingBlock(title, 'Decree ')];

  parts.push(C.partyBlock('Plaintiff', data.plaintiff));
  parts.push(C.versusBlock());
  parts.push(C.partyBlock('Defendant/s   ', data.defendants));
  parts.push(emptyPara({ jc: 'both' }));

  parts.push(labelPara('Suit filed on', [run(C.formatDMY(data.suitFiledOn))]));
  parts.push(labelPara('Suit for', [run(C.cleanText(data.suitFor))], { wrapToValue: true }));
  parts.push(claimBlock('Suit Claim', data.suitClaim));

  parts.push(textPara(disposalSentence(data), { after: 0, ind: { firstLine: 720 }, jc: 'both' }));
  parts.push(C.orderBlock(data.orderText));
  parts.push(emptyPara({ after: 0, ind: { firstLine: 720 }, jc: 'both' }));
  parts.push(C.givenUnderBlock(data.decreeDate, ''));
  parts.push(C.signatureBlock());

  parts.push(
    textPara('COST OF THE SUIT', {
      ind: { left: 1440, firstLine: 720 },
      runOpts: { bold: true, underline: true },
    })
  );
  parts.push(costTable(data));
  parts.push(emptyPara({ after: 0 }));

  if (data.includeSchedule) {
    parts.push(emptyPara());
    parts.push(scheduleBlock(data.scheduleText));
  }

  // The closing signature follows the cost table whether or not a schedule is added.
  parts.push(emptyPara());
  parts.push(C.signatureBlock());

  parts.push(emptyPara());
  parts.push(C.documentClose());
  return parts.join('');
}
