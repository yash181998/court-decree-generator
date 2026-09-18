import { generateDocx, previewFit, DEFAULT_COURT_LINE } from './notice-docx.js';
import { formatAddress } from './address.js';

const $ = (id) => document.getElementById(id);

const VARIANT_LABELS = {
  LRS: 'Lr\u2019s of Defendant No.',
  PROPOSED: 'Proposed Defendant No.',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

let batch = [];
let onChange = () => {};

function ordinal(n) {
  const rem100 = n % 100;
  if (rem100 >= 11 && rem100 <= 13) return `${n}th`;
  return `${n}${{ 1: 'st', 2: 'nd', 3: 'rd' }[n % 10] || 'th'}`;
}

function parseIso(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
}

/** `2026-09-11` -> `11-09-2026`, the form printed on the notice. */
function toDmy(iso) {
  const p = parseIso(iso);
  return p ? `${String(p.d).padStart(2, '0')}-${String(p.m).padStart(2, '0')}-${p.y}` : '';
}

function givenParts(iso) {
  const p = parseIso(iso);
  if (!p) return { day: '', month: '', year: '' };
  return { day: ordinal(p.d), month: MONTHS[p.m - 1], year: String(p.y) };
}

function currentVariant() {
  const checked = document.querySelector('input[name="n-variant"]:checked');
  return checked ? checked.value : 'LRS';
}

/* ------------------------------ parties ------------------------------- */

function refreshPartyLabels() {
  const label = VARIANT_LABELS[currentVariant()];
  $('n-parties').querySelectorAll('fieldset.party').forEach((el, i) => {
    el.querySelector('legend').textContent = `Defendant ${i + 1}`;
    el.querySelector('.p-number-label').textContent = label;
    const number = el.querySelector('.p-number').value.trim();
    el.querySelector('.p-preview').textContent = `${label}${number || '___'}`;
  });
}

function tidyField(el) {
  const tidied = formatAddress(el.value);
  if (tidied && tidied !== el.value) {
    el.value = tidied;
    onChange();
  }
}

/**
 * Shows exactly what formatAddress will print, without touching what the user
 * typed — rewriting the field on every paste/blur silently discarded manual
 * line breaks and made typing on mobile feel broken.
 */
function updateAddressPreview(el, previewEl) {
  if (!previewEl) return;
  previewEl.textContent = formatAddress(el.value);
}

function addParty(values = {}) {
  const node = $('n-partyTemplate').content.firstElementChild.cloneNode(true);
  node.querySelector('.p-number').value = values.number || '';
  node.querySelector('.p-address').value = values.address || '';
  node.querySelector('.p-same').checked = values.useCommonAddress !== false;

  const same = node.querySelector('.p-same');
  const addrWrap = node.querySelector('.p-address-wrap');
  const syncAddr = () => addrWrap.classList.toggle('is-hidden', same.checked);
  same.addEventListener('change', () => { syncAddr(); onChange(); });
  syncAddr();

  node.querySelector('.remove').addEventListener('click', () => {
    node.remove();
    refreshPartyLabels();
    onChange();
  });

  node.querySelector('.duplicate').addEventListener('click', () => {
    addParty({
      number: '',
      address: node.querySelector('.p-address').value,
      useCommonAddress: same.checked,
    });
    const inputs = $('n-parties').querySelectorAll('.p-number');
    inputs[inputs.length - 1].focus();
    onChange();
  });

  node.addEventListener('input', onChange);
  $('n-parties').appendChild(node);
  refreshPartyLabels();
}

/* ------------------------------ collect ------------------------------- */

export function collectNotice() {
  const hearing = toDmy($('n-hearingPicker').value);
  const given = givenParts($('n-givenPicker').value);
  return {
    courtLine: $('n-courtLine').value,
    hearingDate: hearing,
    caseNumber: $('n-caseNumber').value,
    plaintiff: $('n-plaintiff').value,
    defendant: $('n-defendant').value,
    appearDate: $('n-appearSameAsHg').checked ? hearing : toDmy($('n-appearPicker').value),
    givenDay: given.day,
    givenMonth: given.month,
    givenYear: given.year,
    variant: currentVariant(),
    rpdPresent: $('n-rpdPresent').checked,
    packTwoPerPage: $('n-packTwoPerPage').checked,
    commonAddress: $('n-commonAddress').value,
    parties: [...$('n-parties').querySelectorAll('fieldset.party')].map((el) => ({
      number: el.querySelector('.p-number').value,
      address: el.querySelector('.p-address').value,
      useCommonAddress: el.querySelector('.p-same').checked,
    })),
  };
}

/* ------------------------------- batch -------------------------------- */

function renderBatch() {
  const list = $('n-batchList');
  list.innerHTML = '';
  batch.forEach((item, i) => {
    const li = document.createElement('li');
    li.textContent = `${item.caseNumber || '(no case number)'} \u2014 ${item.parties.length} defendant(s)`;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'secondary';
    btn.textContent = 'Remove';
    btn.addEventListener('click', () => {
      batch.splice(i, 1);
      renderBatch();
      onChange();
    });
    li.appendChild(btn);
    list.appendChild(li);
  });
}

export function batchSize() {
  return batch.length;
}

/* ------------------------------ display ------------------------------- */

export function refreshNotice() {
  const rpd = $('n-rpdPresent').checked;
  $('n-rpdHint').textContent = rpd
    ? 'RPD present: one notice per page (single side).'
    : 'RPD absent: two identical copies side by side on the same page.';
  document.querySelectorAll('[data-only="pack"]').forEach((el) => el.classList.toggle('is-hidden', !rpd));
  document.querySelectorAll('[data-only="appear"]').forEach((el) =>
    el.classList.toggle('is-hidden', $('n-appearSameAsHg').checked)
  );

  $('n-hearingEcho').textContent = toDmy($('n-hearingPicker').value) || '\u2014';
  const given = givenParts($('n-givenPicker').value);
  $('n-givenEcho').textContent = given.day
    ? `this ${given.day} day of ${given.month} ${given.year}`
    : '\u2014';

  refreshPartyLabels();

  updateAddressPreview($('n-commonAddress'), $('n-commonAddress-preview'));
  $('n-parties').querySelectorAll('fieldset.party').forEach((el) => {
    updateAddressPreview(el.querySelector('.p-address'), el.querySelector('.p-address-preview'));
  });

  try {
    const data = previewFit(collectNotice());
    const pct = Math.round(data.scale * 100);
    const spacing = ['normal spacing', 'tight spacing', 'no blank lines'][data.compact] || '';
    const queued = batch.length ? ` \u00b7 ${batch.length} queued` : '';
    let text;
    if (!data.fits) {
      text = `Too long even at minimum size \u2014 shorten an address. (${data.usedPt}pt of ${data.limitPt}pt)`;
    } else {
      text = `${data.pages} page(s) \u00b7 font ${pct}%, ${spacing} (${data.usedPt}pt of ${data.limitPt}pt)${queued}`;
      if (data.tight) text += ' \u00b7 close to the limit \u2014 check the page count before printing';
    }
    $('n-fit').textContent = text;
    $('n-fit').classList.toggle('is-warn', !data.fits || data.tight);
  } catch (err) {
    $('n-fit').textContent = err.message;
    $('n-fit').classList.add('is-warn');
  }
}

/** Builds the queued batch when there is one, otherwise the current case. */
export function buildNotice() {
  return generateDocx(batch.length ? { cases: batch } : collectNotice());
}

function applyDateDefaults() {
  $('n-courtLine').value = DEFAULT_COURT_LINE;
  const today = new Date();
  $('n-givenPicker').value = [
    today.getFullYear(),
    String(today.getMonth() + 1).padStart(2, '0'),
    String(today.getDate()).padStart(2, '0'),
  ].join('-');
}

/** Wipes case-specific data (parties, batch) back to a single blank defendant. */
export function resetNotice() {
  applyDateDefaults();
  batch = [];
  renderBatch();
  $('n-parties').innerHTML = '';
  addParty();
}

export function setupNotice(handleChange) {
  onChange = handleChange;

  applyDateDefaults();
  addParty();

  $('n-addParty').addEventListener('click', () => { addParty(); onChange(); });
  $('n-tidyAddress').addEventListener('click', () => {
    tidyField($('n-commonAddress'));
    $('n-parties').querySelectorAll('.p-address').forEach(tidyField);
  });

  $('n-addToBatch').addEventListener('click', () => {
    const current = collectNotice();
    if (!current.caseNumber.trim() && !current.plaintiff.trim()) {
      $('n-fit').textContent = 'Fill the case details before adding to the batch.';
      $('n-fit').classList.add('is-warn');
      return;
    }
    batch.push(current);
    renderBatch();
    onChange();
  });

  $('n-clearBatch').addEventListener('click', () => {
    batch = [];
    renderBatch();
    onChange();
  });

  renderBatch();
}
