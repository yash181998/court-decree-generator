import * as mc from './mc.js';
import * as os from './os.js';
import { headerXml, PARTY_FORMAT_BUILD, smartFormatParty } from './common.js';
import { packDocx, safeFileName } from './docx.js';
import { setupNotice, refreshNotice, buildNotice, batchSize, resetNotice } from './notice-ui.js';

const $ = (id) => document.getElementById(id);
const form = $('decree-form');
const statusEl = $('status');
const resultEl = $('result');
const resultLink = $('result-link');

const DISPOSAL_PREFIX =
  'This Petition coming on this day for final disposal before Sri. Abdul Saleem B.A. (LAW) L.L.B. N. ' +
  'Prl. Senior Civil Judge, Bengaluru Rural District, Bengaluru, in the presence of ';

const buildTagEl = $('build-tag');
// A missing/stale build-info.js (e.g. forgotten commit) must not break the rest of the app.
import('./build-info.js')
  .then(({ BUILD_TIME }) => {
    if (buildTagEl) buildTagEl.textContent = `\u00b7 Deployed ${BUILD_TIME}`;
  })
  .catch(() => {
    if (buildTagEl) buildTagEl.textContent = `\u00b7 ${PARTY_FORMAT_BUILD}`;
  });

function setStatus(message, isError) {
  statusEl.textContent = message;
  statusEl.classList.toggle('is-error', !!isError);
}

/** Reading a field that a stale cached page does not have must not kill the form. */
function val(id) {
  const el = $(id);
  if (!el) {
    console.warn(`Missing field #${id}`);
    return '';
  }
  return el.type === 'checkbox' ? el.checked : el.value;
}

/* ------------------------------- tabs -------------------------------- */

function selectCase(caseType) {
  $('caseType').value = caseType;
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('is-active', tab.dataset.case === caseType);
  });
  document.querySelectorAll('.panel').forEach((panel) => {
    panel.classList.toggle('is-hidden', panel.dataset.panel !== caseType);
  });
  refresh();
}

document.querySelectorAll('.tab').forEach((tab) => {
  tab.addEventListener('click', () => selectCase(tab.dataset.case));
});

/* ------------------------ conditional visibility ---------------------- */

function toggle(selector, visible) {
  document.querySelectorAll(selector).forEach((el) => el.classList.toggle('is-hidden', !visible));
}

function refresh() {
  const both = val('mc-bothPetitioners');
  toggle('[data-only="separate"]', !both);
  toggle('[data-only="nature-custom"]', val('mc-natureOfPetition') === 'custom');
  toggle('[data-only="exparte"]', val('os-exparte'));
  toggle('[data-only="schedule"]', val('os-includeSchedule'));
  syncSuitFor();

  const petitionerAdvocate = val('mc-petitionerAdvocate').trim() || '\u2026';
  $('mc-disposal-preview').textContent = both
    ? `${DISPOSAL_PREFIX}${petitionerAdvocate}. Advocate for the Petitioners.`
    : `${DISPOSAL_PREFIX}${petitionerAdvocate} Advocate for the Petitioner ` +
      `and ${val('mc-respondentAdvocate').trim() || '\u2026'} Advocate for Respondent.`;

  $('os-disposal-preview').textContent = os.disposalSentence({
    plaintiffAdvocate: val('os-plaintiffAdvocate').trim() || '\u2026',
    plaintiffNumbers: val('os-plaintiffNumbers'),
    plaintiffAdvocatesExtra: extraAdvocates('plaintiff'),
    defendantAdvocate: val('os-defendantAdvocate').trim() || '\u2026',
    defendantNumbers: val('os-defendantNumbers'),
    defendantAdvocatesExtra: extraAdvocates('defendant'),
    exparte: val('os-exparte'),
    exparteNumbers: val('os-exparteNumbers'),
  });

  const court = os.parseAmount(val('os-courtFee'));
  const process = os.parseAmount(val('os-processFee'));
  const total = court == null && process == null ? null : (court || 0) + (process || 0);
  $('os-total').textContent = os.formatAmount(total);

  showFormatPreview('mc-petitioners');
  showFormatPreview('mc-respondents');
  showFormatPreview('os-plaintiff');
  showFormatPreview('os-defendants');

  refreshNotice();
}

/** The dropdown drives the free-text field; "Other" leaves it for manual typing. */
function syncSuitFor() {
  const choice = val('os-suitForSelect');
  toggle('[data-only="suitfor-custom"]', choice === 'custom');
  if (choice && choice !== 'custom') $('os-suitFor').value = choice;
}

/** Advocate rows added beyond the first (which carries the fixed field ids). */
function extraAdvocates(role) {
  const container = $(`os-${role}Advocates`);
  if (!container) return [];
  return Array.from(container.querySelectorAll('.advocate-row'))
    .slice(1)
    .map((row) => ({
      advocate: row.querySelector('.adv-name').value,
      numbers: row.querySelector('.adv-no').value,
    }))
    .filter((entry) => entry.advocate.trim());
}

function addAdvocateRow(role) {
  const container = $(`os-${role}Advocates`);
  if (!container) return;
  const row = document.createElement('div');
  row.className = 'advocate-row';
  row.innerHTML =
    '<input type="text" class="adv-name" placeholder="Sri. A.B.C">' +
    '<input type="text" class="adv-no" placeholder="No.">' +
    '<button type="button" class="remove-advocate" aria-label="Remove advocate">\u00d7</button>';
  container.appendChild(row);
}

document.querySelectorAll('.add-advocate').forEach((btn) => {
  btn.addEventListener('click', () => {
    addAdvocateRow(btn.dataset.add);
    refresh();
  });
});

form.addEventListener('click', (event) => {
  if (event.target.classList.contains('remove-advocate')) {
    event.target.closest('.advocate-row').remove();
    refresh();
  }
});

/**
 * Renders exactly what smartFormatParty will print, so scrambled paste (a
 * numbering marker stuck mid-line, missing markers, duplicated fragments -
 * common when copying from a PDF table) is visible and fixable before
 * generating, instead of only showing up in the finished .docx.
 */
function showFormatPreview(fieldId) {
  const preview = $(`${fieldId}-preview`);
  if (!preview) return;
  const lines = smartFormatParty(val(fieldId));
  preview.textContent = lines.length ? lines.join('\n') : '';

  const dup = findDuplicateBlock(lines);
  let warning = preview.nextElementSibling;
  if (!warning || !warning.classList.contains('format-warning')) {
    warning = document.createElement('p');
    warning.className = 'format-warning is-hidden';
    preview.insertAdjacentElement('afterend', warning);
  }
  if (dup) {
    warning.textContent =
      `\u26a0 The lines starting with "${dup.text}" appear twice \u2014 check you ` +
      `haven\u2019t pasted the same block in more than once.`;
    warning.classList.remove('is-hidden');
  } else {
    warning.classList.add('is-hidden');
  }
}

// A run of 3+ consecutive non-blank lines that repeats elsewhere almost always
// means the same chunk was pasted in twice by accident. This only warns - a
// genuinely shared address for two different people is common and left alone.
const MIN_DUPLICATE_RUN = 3;

function findDuplicateBlock(lines) {
  const nonBlank = [];
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i] !== '') nonBlank.push(i);
  }
  for (let a = 0; a < nonBlank.length; a += 1) {
    for (let b = a + MIN_DUPLICATE_RUN; b < nonBlank.length; b += 1) {
      let run = 0;
      while (
        a + run < nonBlank.length &&
        b + run < nonBlank.length &&
        lines[nonBlank[a + run]] === lines[nonBlank[b + run]]
      ) {
        run += 1;
      }
      if (run >= MIN_DUPLICATE_RUN) {
        return { text: lines[nonBlank[a]].slice(0, 50) };
      }
    }
  }
  return null;
}

form.addEventListener('input', refresh);
form.addEventListener('change', refresh);

/* ---------------------------- clear data ------------------------------ */

/** Extra advocate rows are added dynamically and aren't reset by clearing input values. */
function removeExtraAdvocateRows() {
  ['plaintiff', 'defendant'].forEach((role) => {
    const container = $(`os-${role}Advocates`);
    if (!container) return;
    Array.from(container.querySelectorAll('.advocate-row')).slice(1).forEach((row) => row.remove());
  });
}

$('clear-data').addEventListener('click', () => {
  const caseType = val('caseType');
  const panel = document.querySelector(`.panel[data-panel="${caseType}"]`);
  if (panel) {
    panel.querySelectorAll('input[type="text"], input[type="date"], textarea').forEach((el) => {
      el.value = el.defaultValue;
    });
    panel.querySelectorAll('select').forEach((el) => {
      const def = Array.from(el.options).find((opt) => opt.defaultSelected);
      el.value = def ? def.value : (el.options[0] ? el.options[0].value : '');
    });
    panel.querySelectorAll('input[type="checkbox"], input[type="radio"]').forEach((el) => {
      el.checked = el.defaultChecked;
    });
  }
  if (caseType === 'os') removeExtraAdvocateRows();
  if (caseType === 'notice') resetNotice();

  resultEl.classList.add('is-hidden');
  setStatus('');
  refresh();
});

/* ------------------------------ collect ------------------------------- */

function collect() {
  const caseType = val('caseType');
  if (caseType === 'mc') {
    return {
      caseType,
      caseNumber: val('mc-caseNumber'),
      bothPetitioners: val('mc-bothPetitioners'),
      petitioners: val('mc-petitioners'),
      respondents: val('mc-respondents'),
      suitFiledOn: val('mc-suitFiledOn'),
      natureOfPetition: val('mc-natureOfPetition'),
      natureCustom: val('mc-natureCustom'),
      petitionClaim: val('mc-petitionClaim'),
      petitionerAdvocate: val('mc-petitionerAdvocate'),
      respondentAdvocate: val('mc-respondentAdvocate'),
      orderText: val('mc-orderText'),
      decreeDate: val('mc-decreeDate'),
    };
  }
  return {
    caseType,
    caseNumber: val('os-caseNumber'),
    plaintiff: val('os-plaintiff'),
    defendants: val('os-defendants'),
    suitFiledOn: val('os-suitFiledOn'),
    suitFor: val('os-suitFor'),
    suitClaim: val('os-suitClaim'),
    plaintiffAdvocate: val('os-plaintiffAdvocate'),
    plaintiffNumbers: val('os-plaintiffNumbers'),
    plaintiffAdvocatesExtra: extraAdvocates('plaintiff'),
    defendantAdvocate: val('os-defendantAdvocate'),
    defendantNumbers: val('os-defendantNumbers'),
    defendantAdvocatesExtra: extraAdvocates('defendant'),
    exparte: val('os-exparte'),
    exparteNumbers: val('os-exparteNumbers'),
    orderText: val('os-orderText'),
    decreeDate: val('os-decreeDate'),
    courtFee: val('os-courtFee'),
    processFee: val('os-processFee'),
    includeSchedule: val('os-includeSchedule'),
    scheduleText: val('os-scheduleText'),
  };
}

/* ----------------------------- generate ------------------------------- */

let lastUrl = null;

function offerDownload(blob, filename) {
  if (lastUrl) URL.revokeObjectURL(lastUrl);
  lastUrl = URL.createObjectURL(blob);

  // The link stays on the page: some mobile browsers ignore a synthetic click.
  resultLink.href = lastUrl;
  resultLink.download = filename;
  resultLink.textContent = `Tap to save ${filename}`;
  resultEl.classList.remove('is-hidden');

  const auto = document.createElement('a');
  auto.href = lastUrl;
  auto.download = filename;
  document.body.appendChild(auto);
  auto.click();
  auto.remove();
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const caseType = val('caseType');
  const button = $('generate');
  button.disabled = true;
  setStatus('Generating\u2026');
  try {
    if (caseType === 'notice') {
      const { blob, filename, fit } = await buildNotice();
      offerDownload(blob, filename);
      const queued = batchSize();
      setStatus(
        `Ready \u2014 ${queued || 1} case(s), ${fit.pages} page(s), font ${Math.round(fit.scale * 100)}%` +
          (fit.fits ? '' : ' \u2014 STILL OVERFLOWS, shorten the address'),
        !fit.fits
      );
      return;
    }

    const payload = collect();
    if (!payload.caseNumber.trim()) {
      const field = $(`${payload.caseType}-caseNumber`);
      setStatus('Enter the case number first.', true);
      if (field) {
        field.scrollIntoView({ block: 'center', behavior: 'smooth' });
        field.focus();
      }
      return;
    }

    const builder = payload.caseType === 'os' ? os : mc;
    const title = builder.caseTitle(payload.caseNumber);
    const blob = await packDocx({
      documentXml: builder.buildDocument(payload),
      headerXml: headerXml(title),
    });

    offerDownload(blob, `${safeFileName(title)}.docx`);
    setStatus(`Ready \u2014 ${(blob.size / 1024).toFixed(0)} KB`);
  } catch (err) {
    setStatus(err.message, true);
    console.error(err);
  } finally {
    button.disabled = false;
  }
});

setupNotice(refresh);
refresh();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch((err) => console.warn('SW failed', err));
  });
}
