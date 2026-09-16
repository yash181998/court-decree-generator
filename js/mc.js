import { run, textPara, emptyPara, labelPara } from './ooxml.js';
import * as C from './common.js';

export const NATURE_PRESETS = {
  'divorce-act': 'Under section 10(1) (ix) of Indian Divorce Act 1869',
  'hindu-marriage-act': 'Under section 13(1) (i-a) of Hindu Marriage Act',
  'hindu-marriage-act-13ab': 'Under section 13 (1), (i-a)(i-b) of the Hindu Marriage Act',
  'hindu-marriage-act-13b': 'Under Section 13B of the Hindu Marriage Act',
};

export function caseTitle(caseNumber) {
  const n = String(caseNumber || '').trim();
  return /m\.?c\.?\s*no/i.test(n) ? n : `M.C.NO.${n}`;
}

function natureText(data) {
  if (data.natureOfPetition === 'custom') return C.cleanText(data.natureCustom);
  return NATURE_PRESETS[data.natureOfPetition] || '';
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

function disposalSentence(data) {
  if (data.bothPetitioners) {
    return `${C.DISPOSAL_PREFIX}${C.cleanText(data.petitionerAdvocate)}. Advocate for the Petitioners.`;
  }
  return (
    `${C.DISPOSAL_PREFIX}${C.cleanText(data.petitionerAdvocate)} Advocate for the Petitioner ` +
    `and ${C.cleanText(data.respondentAdvocate)} Advocate for Respondent.`
  );
}

export function buildDocument(data) {
  const title = caseTitle(data.caseNumber);
  const parts = [C.documentOpen(), C.headingBlock(title, 'DECREE')];

  parts.push(C.partyBlock('Petitioner/s', data.petitioners));
  parts.push(C.versusBlock());
  parts.push(C.partyBlock('Respondent/s   ', data.bothPetitioners ? 'Nil' : data.respondents));
  parts.push(emptyPara({ jc: 'both' }));

  parts.push(labelPara('Suit filed on', [run(C.formatDMY(data.suitFiledOn))]));
  parts.push(labelPara('Nature of Petition', [run(natureText(data))], { wrapToValue: true }));
  parts.push(claimBlock('Petition Claim', data.petitionClaim));

  parts.push(textPara(disposalSentence(data), { ind: { firstLine: 720 }, jc: 'both' }));
  parts.push(emptyPara({ ind: { firstLine: 720 }, jc: 'both' }));

  parts.push(C.orderBlock(data.orderText));
  parts.push(emptyPara({ after: 0, ind: { firstLine: 720 }, jc: 'both' }));
  parts.push(C.givenUnderBlock(data.decreeDate));
  parts.push(C.signatureBlock());

  parts.push(emptyPara());
  parts.push(C.documentClose());
  return parts.join('');
}
