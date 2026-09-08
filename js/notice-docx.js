import { findFit, buildBodyXml, expectedPages } from './notice.js';
import { formatAddress } from './address.js';
import { NOTICE_TEMPLATE_BASE64 } from './notice-template.js';

export const DEFAULT_COURT_LINE =
  'IN THE COURT OF PRL. SENIOR CIVIL JUDGE BENGALURU RURAL DISTRICT , BENGALURU.';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

// The desktop tool asks Word to confirm the page count once a layout passes this
// share of the column. A phone has no Word, so it says so instead.
const TIGHT_RATIO = 0.85;

function str(value, fallback = '') {
  const s = typeof value === 'string' ? value.trim() : '';
  return s || fallback;
}

function normalise(payload) {
  const body = payload && typeof payload === 'object' ? payload : {};
  const rawParties = Array.isArray(body.parties) ? body.parties.slice(0, 50) : [];

  const form = {
    courtLine: str(body.courtLine, DEFAULT_COURT_LINE),
    hearingDate: str(body.hearingDate),
    caseNumber: str(body.caseNumber),
    plaintiff: str(body.plaintiff),
    defendant: str(body.defendant),
    appearDate: str(body.appearDate) || str(body.hearingDate),
    appearTime: str(body.appearTime, "11.00 O\u2019clock"),
    givenDay: str(body.givenDay),
    givenMonth: str(body.givenMonth),
    givenYear: str(body.givenYear),
    variant: body.variant === 'PROPOSED' ? 'PROPOSED' : 'LRS',
    rpdPresent: Boolean(body.rpdPresent),
    packTwoPerPage: Boolean(body.packTwoPerPage),
    commonAddress: formatAddress(String(body.commonAddress || '').slice(0, 4000)),
  };

  const parties = rawParties.map((p) => ({
    number: str(p && p.number),
    address: formatAddress(String((p && p.address) || '').slice(0, 4000)),
    useCommonAddress: Boolean(p && p.useCommonAddress),
  })).filter((p) => p.number || p.address || p.useCommonAddress);

  if (parties.length === 0) {
    parties.push({ number: '', address: '', useCommonAddress: true });
  }

  return { form, parties };
}

/** Accepts either a single case object or `{ cases: [...] }`. */
export function normaliseAll(payload) {
  const list = payload && Array.isArray(payload.cases) && payload.cases.length
    ? payload.cases.slice(0, 100)
    : [payload];
  return list.map(normalise);
}

function templateBytes() {
  const binary = atob(NOTICE_TEMPLATE_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** Layout summary for the status line, without producing a file. */
export function previewFit(payload) {
  const cases = normaliseAll(payload);
  const perCase = cases.map(({ form, parties }) => {
    const fit = findFit(form, parties);
    return {
      caseNumber: form.caseNumber,
      scale: fit.scale,
      compact: fit.compact,
      fits: fit.fits,
      usedPt: Math.round(fit.heightPt),
      limitPt: Math.round(fit.limitPt),
      pages: expectedPages(form, parties),
    };
  });

  const usedPt = Math.max(...perCase.map((c) => c.usedPt));
  const limitPt = perCase[0].limitPt;

  return {
    cases: perCase.length,
    scale: Math.min(...perCase.map((c) => c.scale)),
    compact: Math.max(...perCase.map((c) => c.compact)),
    fits: perCase.every((c) => c.fits),
    tight: usedPt > limitPt * TIGHT_RATIO,
    usedPt,
    limitPt,
    pages: perCase.reduce((n, c) => n + c.pages, 0),
  };
}

/**
 * Rebuilds the document body inside a copy of the original .docx, so every
 * style, font and theme part of the source file is preserved untouched.
 *
 * The desktop tool can ask Word to confirm the page count; on a phone there is
 * no Word, so the font-metric estimate in notice.js is the final word.
 */
export async function generateDocx(payload) {
  const cases = normaliseAll(payload);
  const estimates = cases.map((c) => findFit(c.form, c.parties));

  const zip = await window.JSZip.loadAsync(templateBytes());
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) throw new Error('The template does not contain word/document.xml.');

  const original = await documentFile.async('string');
  const bodyStart = original.indexOf('<w:body>');
  if (bodyStart === -1) throw new Error('Unexpected template structure: no <w:body> element.');
  const header = original.slice(0, bodyStart);

  const layouts = cases.map((c, i) => ({
    form: c.form,
    parties: c.parties,
    scale: estimates[i].scale,
    compact: estimates[i].compact,
  }));

  zip.file('word/document.xml', `${header}${buildBodyXml(layouts)}</w:document>`);
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME,
    compression: 'DEFLATE',
  });

  const label = cases.length > 1
    ? `Batch_${cases.length}_cases`
    : (cases[0].form.caseNumber || 'notice');
  const filename = `Notice_${label.replace(/[^a-z0-9._-]+/gi, '_').slice(0, 60)}.docx`;

  return { blob, filename, fit: previewFit(payload) };
}
