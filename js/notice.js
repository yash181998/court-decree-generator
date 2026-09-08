import { lineCount } from './measure.js';

// Page geometry copied from the original template's <w:sectPr>:
// A4 landscape, 2cm margins, two newspaper columns with 708 twip gutter.
const PAGE_W = 16838;
const PAGE_H = 11906;
const MARGIN = 1134;
const COL_SPACE = 708;
const COL_W_TWIPS = (PAGE_W - MARGIN * 2 - COL_SPACE) / 2; // 6931
const COL_W_PT = COL_W_TWIPS / 20;
export const COL_H_PT = (PAGE_H - MARGIN * 2) / 20;
const SAFETY = 0.96; // leave a sliver so Word's own rounding never spills a line
const LINE_FACTOR = 1.15; // Times New Roman single-spaced line height per em

export { COL_W_PT };

export const ORDER_VARIANTS = {
  LRS: {
    clause: 'U/O 22 Rule 4 of CPC',
    label: "Lr\u2019s of Defendant No.",
  },
  PROPOSED: {
    clause: 'U/O 1 Rule 10(2) of CPC',
    label: 'Proposed Defendant No.',
  },
};

function esc(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function ordinal(n) {
  const num = Number(n);
  if (!Number.isFinite(num)) return String(n);
  const rem100 = num % 100;
  if (rem100 >= 11 && rem100 <= 13) return `${num}th`;
  switch (num % 10) {
    case 1: return `${num}st`;
    case 2: return `${num}nd`;
    case 3: return `${num}rd`;
    default: return `${num}th`;
  }
}

// Font sizes are half-points, matching the original document.
const BASE = { big: 32, body: 28, court: 24 };

function scaleSize(halfPoints, scale) {
  const scaled = Math.round(halfPoints * scale);
  return Math.max(14, scaled % 2 === 0 ? scaled : scaled + 1);
}

// Whitespace is squeezed before type is shrunk, so a long address costs blank
// lines first and legibility last. 0 = original gaps, 1 = tight, 2 = none.
const GAP_RATIO = [0.6, 0.3, 0];

/**
 * Describes one notice as an ordered list of paragraphs. The same description
 * feeds both the height estimator and the XML writer, so what we measure is
 * exactly what gets written.
 */
function describeNotice(form, party, scale, compact = 0) {
  const variant = ORDER_VARIANTS[form.variant] || ORDER_VARIANTS.LRS;
  const big = scaleSize(BASE.big, scale);
  const body = scaleSize(BASE.body, scale);
  const court = scaleSize(BASE.court, scale);
  const ratio = GAP_RATIO[Math.min(compact, GAP_RATIO.length - 1)];
  const gap = ratio ? Math.max(6, Math.round(body * ratio)) : 0;

  const addressLines = (party.useCommonAddress ? form.commonAddress : party.address || '')
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const paras = [];
  const pushGap = () => { if (gap) paras.push({ sz: gap, runs: [] }); };

  paras.push({ align: 'right', sz: big, runs: [{ t: `Hg. Date: ${form.hearingDate}` }] });
  paras.push({ align: 'center', sz: court, runs: [{ t: form.courtLine, b: true, u: true }] });
  pushGap();
  // Centred on the column's midpoint rather than indented from it, so the
  // number sits in the middle whatever its length.
  paras.push({
    sz: body,
    tabs: [{ type: 'center', pos: Math.round(COL_W_TWIPS * 0.5) }],
    runs: [{ tab: true }, { t: form.caseNumber, b: true, u: true }],
  });
  paras.push({
    sz: body,
    tabs: [{ type: 'right', pos: COL_W_TWIPS }],
    runs: [{ t: 'PLAINTIFF', b: true, u: true }, { tab: true }, { t: 'DEFENDANT', b: true, u: true }],
  });
  paras.push({
    sz: body,
    tabs: [{ type: 'center', pos: Math.round(COL_W_TWIPS / 2) }, { type: 'right', pos: COL_W_TWIPS }],
    runs: [{ t: form.plaintiff }, { tab: true }, { t: 'V/s' }, { tab: true }, { t: form.defendant }],
  });
  pushGap();
  paras.push({ align: 'center', sz: body, runs: [{ t: 'NOTICE', b: true, u: true }] });
  pushGap();
  paras.push({
    align: 'both',
    sz: body,
    firstIndent: 720,
    runs: [{ t: `Whereas the above named plaintiff has made an application to this court that ${variant.clause}.` }],
  });
  paras.push({
    align: 'both',
    sz: body,
    firstIndent: 720,
    runs: [{
      t: `You are hereby informed to appear before this court in person or by a pleader duly instructed on ${form.appearDate} at ${form.appearTime} in the forenoon, failing wherein the said application will be heard and determined in your absence.`,
    }],
  });
  paras.push({
    align: 'both',
    sz: body,
    firstIndent: 720,
    runs: [{
      t: `Given under my hand and seal of the court, this ${form.givenDay} day of ${form.givenMonth} ${form.givenYear}`,
    }],
  });
  pushGap();

  const headRuns = [{ t: variant.label, b: true, u: true }];
  if (party.number) headRuns.push({ t: party.number });
  paras.push({ sz: big, runs: headRuns });

  for (const line of addressLines) {
    paras.push({ sz: big, runs: [{ t: line }] });
  }

  return paras;
}

function paraHeightPt(para) {
  const fontPt = para.sz / 2;
  if (!para.runs.length) return fontPt * LINE_FACTOR;

  const widthPt = COL_W_PT - (para.leftIndent || 0) / 20;

  if (para.tabs && para.tabs.length) {
    // Tabbed rows are laid out on a single line; they only wrap if the combined
    // text is wider than the column, which the estimator below catches.
    const text = para.runs.map((r) => (r.tab ? ' ' : r.t)).join(' ');
    const bold = para.runs.some((r) => r.b);
    return Math.max(1, lineCount(text, fontPt, bold, widthPt)) * fontPt * LINE_FACTOR;
  }

  const text = para.runs.map((r) => (r.tab ? ' ' : r.t)).join('');
  const bold = para.runs.every((r) => r.tab || r.b);
  const indentPt = (para.firstIndent || 0) / 20;
  return lineCount(text, fontPt, bold, widthPt, indentPt) * fontPt * LINE_FACTOR;
}

function noticeHeightPt(form, party, scale, compact) {
  return describeNotice(form, party, scale, compact).reduce((sum, p) => sum + paraHeightPt(p), 0);
}

/**
 * Finds the roomiest layout that still fits one notice per column: full size
 * with normal gaps first, then tighter gaps, and only then a smaller font.
 */
export function findFit(form, parties) {
  const limit = COL_H_PT * SAFETY;
  const worstAt = (scale, compact) => Math.max(...parties.map((p) => noticeHeightPt(form, p, scale, compact)));

  for (let step = 100; step >= 50; step -= 2) {
    const scale = step / 100;
    for (let compact = 0; compact < GAP_RATIO.length; compact += 1) {
      const worst = worstAt(scale, compact);
      if (worst <= limit) {
        return { scale, compact, heightPt: worst, limitPt: limit, fits: true };
      }
    }
  }
  return { scale: 0.5, compact: 2, heightPt: worstAt(0.5, 2), limitPt: limit, fits: false };
}

function runXml(run, sz) {
  const props = [
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>',
    run.b ? '<w:b/><w:bCs/>' : '',
    `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`,
    run.u ? '<w:u w:val="single"/>' : '',
  ].join('');
  const content = run.tab ? '<w:tab/>' : `<w:t xml:space="preserve">${esc(run.t)}</w:t>`;
  return `<w:r><w:rPr>${props}</w:rPr>${content}</w:r>`;
}

function paraXml(para) {
  const tabs = para.tabs && para.tabs.length
    ? `<w:tabs>${para.tabs.map((t) => `<w:tab w:val="${t.type}" w:pos="${t.pos}"/>`).join('')}</w:tabs>`
    : '';
  const indentAttrs = [
    para.leftIndent ? ` w:left="${para.leftIndent}"` : '',
    para.firstIndent ? ` w:firstLine="${para.firstIndent}"` : '',
  ].join('');
  const indent = indentAttrs ? `<w:ind${indentAttrs}/>` : '';
  const align = para.align ? `<w:jc w:val="${para.align}"/>` : '';
  const rPr = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="${para.sz}"/><w:szCs w:val="${para.sz}"/></w:rPr>`;
  const pPr = `<w:pPr>${tabs}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${indent}${align}${rPr}</w:pPr>`;
  return `<w:p>${pPr}${para.runs.map((r) => runXml(r, para.sz)).join('')}</w:p>`;
}

function breakXml(type) {
  // A 1pt run keeps the break paragraph from stealing a visible line from the
  // column it terminates.
  return '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    + '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>'
    + `<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:br w:type="${type}"/></w:r></w:p>`;
}

const SECT_PR = `<w:sectPr><w:pgSz w:w="${PAGE_W}" w:h="${PAGE_H}" w:orient="landscape"/>`
  + `<w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="709" w:footer="709" w:gutter="0"/>`
  + `<w:cols w:num="2" w:space="${COL_SPACE}"/><w:docGrid w:linePitch="360"/></w:sectPr>`;

/**
 * Expands the parties into the sequence of columns to print.
 * - RPD absent  -> two identical copies of each notice (they fill one page).
 * - RPD present -> one copy per party; by default the rest of the page is left
 *   blank, or two different parties share a page when packTwoPerPage is set.
 */
export function buildColumnPlan(form, parties) {
  const columns = [];
  if (!form.rpdPresent) {
    for (const party of parties) {
      columns.push({ party, breakAfter: 'column' });
      columns.push({ party, breakAfter: 'page' });
    }
  } else if (form.packTwoPerPage) {
    parties.forEach((party, i) => {
      columns.push({ party, breakAfter: i % 2 === 0 ? 'column' : 'page' });
    });
  } else {
    for (const party of parties) {
      columns.push({ party, breakAfter: 'page' });
    }
  }
  if (columns.length) columns[columns.length - 1].breakAfter = null;
  return columns;
}

export function expectedPages(form, parties) {
  const plan = buildColumnPlan(form, parties);
  return plan.filter((c) => c.breakAfter === 'page').length + 1;
}

function buildCaseXml(form, parties, scale, compact = 0) {
  const plan = buildColumnPlan(form, parties);
  const chunks = [];
  for (const col of plan) {
    for (const para of describeNotice(form, col.party, scale, compact)) {
      chunks.push(paraXml(para));
    }
    if (col.breakAfter === 'column') {
      chunks.push(breakXml('column'));
    } else if (col.breakAfter === 'page') {
      chunks.push(breakXml('page'));
    }
  }
  return chunks.join('');
}

/**
 * Joins any number of independent cases into one document. Each case starts on
 * a new page and keeps its own font scale, so batching never disturbs a layout
 * that already fitted.
 */
export function buildBodyXml(cases) {
  const body = cases
    .map((c) => buildCaseXml(c.form, c.parties, c.scale, c.compact))
    .join(breakXml('page'));
  return `<w:body>${body}${SECT_PR}</w:body>`;
}
