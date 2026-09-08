'use strict';

// Runs every fixture through the original Node apps and records a hash of the
// XML each one produces. The browser build hashes the same fixtures; if any
// hash differs, the phone app would produce a different document.
// Run with: node tools/check-parity.js

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const DECREE = path.join(__dirname, '..', '..', 'decree-generator', 'src');
const NOTICE = path.join(__dirname, '..', '..', 'notice-generator', 'src');

const mc = require(path.join(DECREE, 'mc.js'));
const os = require(path.join(DECREE, 'os.js'));
const noticeDocx = require(path.join(NOTICE, 'docx.js'));
const notice = require(path.join(NOTICE, 'notice.js'));

const fixtures = JSON.parse(fs.readFileSync(path.join(__dirname, 'parity-fixtures.json'), 'utf8'));

function sha256(text) {
  return crypto.createHash('sha256').update(text, 'utf8').digest('hex');
}

function buildNoticeXml(data) {
  const cases = noticeDocx.normaliseAll(data);
  const layouts = cases.map((c) => {
    const fit = notice.findFit(c.form, c.parties);
    return { form: c.form, parties: c.parties, scale: fit.scale, compact: fit.compact };
  });
  return notice.buildBodyXml(layouts);
}

const results = {};
for (const fixture of fixtures) {
  let xml;
  if (fixture.kind === 'mc') xml = mc.buildDocument(fixture.data);
  else if (fixture.kind === 'os') xml = os.buildDocument(fixture.data);
  else xml = buildNoticeXml(fixture.data);
  results[fixture.name] = { length: xml.length, sha256: sha256(xml) };
}

const target = path.join(__dirname, 'parity-node.json');
fs.writeFileSync(target, `${JSON.stringify(results, null, 2)}\n`, 'utf8');

for (const [name, r] of Object.entries(results)) {
  console.log(`${name.padEnd(32)} ${String(r.length).padStart(7)}  ${r.sha256.slice(0, 16)}`);
}
console.log(`\nWrote ${target}`);
