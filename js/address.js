// Segments that conventionally begin their own line in an Indian cause-title address.
const LINE_STARTERS = /^(s\/o|d\/o|w\/o|c\/o|h\/o|s\/o\.|aged|age|major|minor|r\/at|r\/o|residing|resident|rep\.?\s*by|represented|since|presently|permanent)\b/i;
const PIN_CODE = /\b\d{6}\b\s*[.,]?\s*$/;
const TARGET_LINE_CHARS = 46;

function tidy(line) {
  return line
    .replace(/\s+/g, ' ')
    .replace(/\s+([,.])/g, '$1')
    .replace(/,\s*,/g, ',')
    .replace(/^[,\s]+|[\s]+$/g, '')
    .trim();
}

/**
 * Turns an address pasted in any shape - one long line, comma soup, ragged
 * copy-paste with stray blank lines - into the line-per-detail layout the
 * notice expects. Text already broken into sensible lines is left alone.
 */
export function formatAddress(raw, targetChars = TARGET_LINE_CHARS) {
  const text = String(raw === undefined || raw === null ? '' : raw).replace(/\r/g, '').trim();
  if (!text) return '';

  const given = text.split('\n').map(tidy).filter(Boolean);
  if (given.length > 1 && given.every((l) => l.length <= targetChars * 1.6)) {
    return given.join('\n');
  }

  const flat = tidy(given.join(', '));
  const segments = flat.split(/\s*,\s*/).map(tidy).filter(Boolean);

  const lines = [];
  let current = '';
  const flush = () => { if (current) { lines.push(current); current = ''; } };

  for (const segment of segments) {
    const tooLong = current && `${current}, ${segment}`.length > targetChars;
    if (current && (LINE_STARTERS.test(segment) || tooLong)) flush();
    current = current ? `${current}, ${segment}` : segment;
    if (PIN_CODE.test(segment)) flush();
  }
  flush();

  return lines
    .map((line, i) => (i < lines.length - 1 && !/[,.]$/.test(line) ? `${line},` : line))
    .join('\n');
}
