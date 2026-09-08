// Advance widths (per 1000 em) for Times New Roman, from the Adobe AFM metrics.
// Used to predict how many lines a paragraph will occupy inside a column so the
// generator can pick a font scale that always fits on one page.
const REGULAR = {
  ' ': 250, '!': 333, '"': 408, '#': 500, $: 500, '%': 833, '&': 778, "'": 333,
  '(': 333, ')': 333, '*': 500, '+': 564, ',': 250, '-': 333, '.': 250, '/': 278,
  0: 500, 1: 500, 2: 500, 3: 500, 4: 500, 5: 500, 6: 500, 7: 500, 8: 500, 9: 500,
  ':': 278, ';': 278, '<': 564, '=': 564, '>': 564, '?': 444, '@': 921,
  A: 722, B: 667, C: 667, D: 722, E: 611, F: 556, G: 722, H: 722, I: 333, J: 389,
  K: 722, L: 611, M: 889, N: 722, O: 722, P: 556, Q: 722, R: 667, S: 556, T: 611,
  U: 722, V: 722, W: 944, X: 722, Y: 722, Z: 611,
  '[': 333, '\\': 278, ']': 333, '^': 469, _: 500, '`': 333,
  a: 444, b: 500, c: 444, d: 500, e: 444, f: 333, g: 500, h: 500, i: 278, j: 278,
  k: 500, l: 278, m: 778, n: 500, o: 500, p: 500, q: 500, r: 333, s: 389, t: 278,
  u: 500, v: 500, w: 722, x: 500, y: 500, z: 444,
  '{': 480, '|': 200, '}': 480, '~': 541, '\u2019': 333, '\u2018': 333, '\u2013': 500, '\u2014': 1000,
};

const BOLD = {
  ' ': 250, '!': 333, '"': 555, '#': 500, $: 500, '%': 1000, '&': 833, "'": 333,
  '(': 333, ')': 333, '*': 500, '+': 570, ',': 250, '-': 333, '.': 250, '/': 278,
  0: 500, 1: 500, 2: 500, 3: 500, 4: 500, 5: 500, 6: 500, 7: 500, 8: 500, 9: 500,
  ':': 333, ';': 333, '<': 570, '=': 570, '>': 570, '?': 500, '@': 930,
  A: 722, B: 667, C: 722, D: 722, E: 667, F: 611, G: 778, H: 778, I: 389, J: 500,
  K: 778, L: 667, M: 944, N: 722, O: 778, P: 611, Q: 778, R: 722, S: 556, T: 667,
  U: 722, V: 722, W: 1000, X: 722, Y: 722, Z: 667,
  '[': 333, '\\': 278, ']': 333, '^': 581, _: 500, '`': 333,
  a: 500, b: 556, c: 444, d: 556, e: 444, f: 333, g: 500, h: 556, i: 278, j: 333,
  k: 556, l: 278, m: 833, n: 556, o: 500, p: 556, q: 556, r: 444, s: 389, t: 333,
  u: 556, v: 500, w: 722, x: 500, y: 500, z: 444,
  '{': 394, '|': 220, '}': 394, '~': 520, '\u2019': 333, '\u2018': 333, '\u2013': 500, '\u2014': 1000,
};

function charEm(ch, bold) {
  const table = bold ? BOLD : REGULAR;
  const w = table[ch];
  return (w === undefined ? 500 : w) / 1000;
}

export function textWidthPt(text, fontPt, bold) {
  let em = 0;
  for (const ch of text) em += charEm(ch, bold);
  return em * fontPt;
}

/**
 * Greedy line-break count for a run of text laid out in a fixed-width column.
 * firstIndentPt shortens only the first line (Word's first-line indent).
 */
export function lineCount(text, fontPt, bold, columnPt, firstIndentPt = 0) {
  const words = String(text).split(/\s+/).filter(Boolean);
  if (words.length === 0) return 1;

  const spacePt = charEm(' ', bold) * fontPt;
  let lines = 1;
  let avail = columnPt - firstIndentPt;
  let used = 0;

  for (const word of words) {
    const w = textWidthPt(word, fontPt, bold);
    if (used === 0) {
      used = w;
    } else if (used + spacePt + w <= avail) {
      used += spacePt + w;
    } else {
      lines += 1;
      avail = columnPt;
      used = w;
    }
    // A single word wider than the column forces extra wrapped lines.
    if (used > avail) {
      const overflow = Math.ceil(used / avail) - 1;
      lines += overflow;
      used = used - overflow * avail;
      avail = columnPt;
    }
  }
  return lines;
}
