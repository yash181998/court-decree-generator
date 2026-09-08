// Low-level WordprocessingML helpers. Every run/paragraph produced here matches
// the formatting used in the reference decrees: Times New Roman, 14pt (sz 28).

const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';
const SIZE = '<w:sz w:val="28"/><w:szCs w:val="28"/>';
const LANG = '<w:lang w:val="en-US"/>';

export function esc(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export function rPr(opts = {}) {
  let xml = FONT;
  if (opts.bold) xml += '<w:b/><w:bCs/>';
  xml += SIZE;
  if (opts.underline) xml += '<w:u w:val="single"/>';
  if (opts.superscript) xml += '<w:vertAlign w:val="superscript"/>';
  xml += LANG;
  return '<w:rPr>' + xml + '</w:rPr>';
}

export function run(text, opts = {}) {
  if (text === '') return '';
  return '<w:r>' + rPr(opts) + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
}

export function tab(opts = {}) {
  return '<w:r>' + rPr(opts) + '<w:tab/></w:r>';
}

function pPr(opts = {}) {
  let xml = '';
  if (opts.tabs && opts.tabs.length) {
    xml += '<w:tabs>' + opts.tabs.map((p) => `<w:tab w:val="left" w:pos="${p}"/>`).join('') + '</w:tabs>';
  }
  xml += `<w:spacing w:after="${opts.after == null ? 30 : opts.after}"/>`;
  if (opts.ind) {
    const i = opts.ind;
    let a = '';
    if (i.left != null) a += ` w:left="${i.left}"`;
    if (i.hanging != null) a += ` w:hanging="${i.hanging}"`;
    if (i.firstLine != null) a += ` w:firstLine="${i.firstLine}"`;
    xml += `<w:ind${a}/>`;
  }
  if (opts.jc) xml += `<w:jc w:val="${opts.jc}"/>`;
  xml += rPr(opts.runOpts || {});
  return '<w:pPr>' + xml + '</w:pPr>';
}

/** Build a paragraph from an array of run XML strings. */
export function para(runs, opts = {}) {
  const body = Array.isArray(runs) ? runs.join('') : runs || '';
  return '<w:p>' + pPr(opts) + body + '</w:p>';
}

/** Plain text paragraph. */
export function textPara(text, opts = {}) {
  return para([run(text, opts.runOpts || {})], opts);
}

/** Empty spacing paragraph. */
export function emptyPara(opts = {}) {
  return para([], opts);
}

// Values in the label block start at 2880 twips (2 inches) from the left margin.
const LABEL_TABS = [2160, 2880];
const VALUE_LEFT = 2880;

/**
 * `Label` <tab> `:` <tab> value — the heading rows (Petitioner/s, Suit filed on, ...).
 * Wrapped lines fall back to the left margin unless `wrapToValue` is set.
 */
export function labelPara(label, valueRuns, opts = {}) {
  const runs = [
    run(label, { bold: true }),
    tab({ bold: true }),
    run(':', { bold: true }),
    tab(),
  ].concat(valueRuns || []);
  return para(runs, {
    tabs: LABEL_TABS,
    jc: opts.jc || 'both',
    after: opts.after,
    ind: opts.wrapToValue ? { left: VALUE_LEFT, hanging: VALUE_LEFT } : undefined,
  });
}

/** Continuation line of a label block, aligned under the value column. */
export function valuePara(text, opts = {}) {
  return para([run(text, opts.runOpts || {})], {
    ind: { left: 2160, firstLine: 720 },
    jc: opts.jc || 'both',
    after: opts.after,
  });
}
