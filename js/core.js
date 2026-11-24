// mpdf core — pure logic, no DOM. Used by the site and by tests.
import { PDFDocument } from '../vendor/pdf-lib.esm.min.js';

export const MODES = {
  FILES: 'files', // each uploaded file is one document: Inv1, DN1, PO1, Inv2, ...
  TYPES: 'types', // one file per document type: all invoices, all DNs, all POs
  PAGES: 'pages', // one combined PDF (or several, joined): every N pages is a set
};

/** Suggest a mode from what was uploaded. files: [{pageCount}] */
export function suggestMode(files, perSet) {
  if (files.length === 1) return MODES.PAGES;
  if (files.length === perSet && files.length > 1) {
    const p = files[0].pageCount;
    if (p > 1 && files.every((f) => f.pageCount === p)) return MODES.TYPES;
  }
  return MODES.FILES;
}

/**
 * Build sets. Returns { sets: [[{file, page}]], warnings: [string] }
 * file = index into files, page = 0-based page index.
 */
export function buildSets(files, mode, perSet) {
  const warnings = [];
  const sets = [];
  perSet = Math.max(1, Math.floor(perSet) || 1);
  if (!files.length) return { sets, warnings };

  if (mode === MODES.FILES) {
    if (files.length % perSet !== 0)
      warnings.push(`${files.length} files don't divide evenly into sets of ${perSet}. The last set has only ${files.length % perSet} file(s).`);
    for (let i = 0; i < files.length; i += perSet) {
      const set = [];
      for (let f = i; f < Math.min(i + perSet, files.length); f++)
        for (let p = 0; p < files[f].pageCount; p++) set.push({ file: f, page: p });
      sets.push(set);
    }
  } else if (mode === MODES.TYPES) {
    const counts = files.map((f) => f.pageCount);
    const min = Math.min(...counts);
    if (counts.some((c) => c !== min))
      warnings.push(`Files have different page counts (${counts.join(', ')}). Only the first ${min} pages of each are used — check that each document is exactly one page.`);
    for (let i = 0; i < min; i++) sets.push(files.map((_, f) => ({ file: f, page: i })));
  } else {
    const all = [];
    files.forEach((f, fi) => { for (let p = 0; p < f.pageCount; p++) all.push({ file: fi, page: p }); });
    if (all.length % perSet !== 0)
      warnings.push(`${all.length} pages don't divide evenly into sets of ${perSet}. The last set has only ${all.length % perSet} page(s).`);
    for (let i = 0; i < all.length; i += perSet) sets.push(all.slice(i, i + perSet));
  }
  return { sets, warnings };
}

/* ---------- Text layout from pdf.js ---------- */

/** Turn pdf.js textContent into lines of words with positions (PDF units, y up). */
export function toLines(textContent, viewport) {
  const items = textContent.items
    .filter((it) => it.str && it.str.trim())
    .map((it) => ({ str: it.str, x: it.transform[4], y: it.transform[5], w: it.width || 0, h: Math.abs(it.transform[3]) || 10 }));
  items.sort((a, b) => b.y - a.y || a.x - b.x);
  const lines = [];
  for (const it of items) {
    let line = lines.find((l) => Math.abs(l.y - it.y) <= Math.max(2, it.h * 0.4));
    if (!line) { line = { y: it.y, items: [] }; lines.push(line); }
    line.items.push(it);
  }
  lines.sort((a, b) => b.y - a.y);
  const W = viewport?.width || 1, H = viewport?.height || 1;
  return lines.map((l) => {
    l.items.sort((a, b) => a.x - b.x);
    const words = [];
    for (const it of l.items) {
      const len = it.str.length || 1;
      const re = /\S+/g; let m;
      while ((m = re.exec(it.str))) {
        const x = it.x + (it.w * m.index) / len;
        words.push({ text: m[0], x, y: it.y, w: (it.w * m[0].length) / len, h: it.h, rx: x / W, ry: it.y / H });
      }
    }
    return { y: l.y, words, text: words.map((w) => w.text).join(' ') };
  });
}

/* ---------- Invoice number detection ---------- */

const LABEL = String.raw`(?:tax\s*invoice|invoice|inv|bill)\s*(?:no|number|num|nr|#|ref(?:erence)?)?\s*\.?\s*[:#\-–]?\s*`;
const VALUE = String.raw`([A-Za-z]{0,6}[\-\/_.]?\d[\w\-\/.]*)`;
const AUTO_RE = new RegExp(LABEL + VALUE, 'i');
const AUTO_RE_STRICT = new RegExp(String.raw`(?:invoice|inv)\s*(?:no|number|num|nr|#)\s*\.?\s*[:#\-–]?\s*` + VALUE, 'i');

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const trimValue = (v) => v.replace(/^[:#\-–]+/, '').replace(/[.,;:)\]]+$/, '');

/** Automatic guess: look for "Invoice No: XXX" style text, incl. value on the next line. */
export function autoDetect(lines) {
  const text = lines.map((l) => l.text);
  for (const re of [AUTO_RE_STRICT, AUTO_RE]) {
    for (const t of text) { const m = t.match(re); if (m) return trimValue(m[1]); }
  }
  // Table header layout: "Invoice No" label, value on the line below at the same x
  for (let i = 0; i < lines.length - 1; i++) {
    const ws = lines[i].words;
    for (let j = 0; j < ws.length; j++) {
      const t = ws[j].text, n = ws[j + 1]?.text || '';
      const isLabel = /^(invoice|inv)\.?(no\.?|#|number)?:?$/i.test(t) &&
        (/^(no|number|num|#)\.?:?$/i.test(n) || /(no|#|number)\.?:?$/i.test(t.replace(/^(invoice|inv)/i, '')));
      if (!isLabel) continue;
      for (const below of lines.slice(i + 1, i + 3).map((l) => l.words)) {
        if (!below.length) continue;
        const near = below.reduce((a, b) => (Math.abs(b.x - ws[j].x) < Math.abs(a.x - ws[j].x) ? b : a));
        if (/\d/.test(near.text) && Math.abs(near.x - ws[j].x) < 60) return trimValue(near.text);
      }
    }
  }
  return '';
}

/**
 * Build a rule from the word the user clicked on a sample invoice.
 * Rule = { label: regex source or null, rx, ry, anchorText, dx, dy }
 */
export function makeRule(lines, lineIdx, wordIdx) {
  const line = lines[lineIdx];
  const word = line.words[wordIdx];
  const before = line.words.slice(Math.max(0, wordIdx - 3), wordIdx).map((w) => w.text);
  // Keep only the label-ish tail (drop leading pure-value words)
  while (before.length > 1 && /\d/.test(before[0]) && !/[a-z]{2}/i.test(before[0])) before.shift();
  let label = null;
  if (before.length) {
    label = before.map((t) => escapeRe(t)).join('\\s*') + String.raw`\s*[:#\-–]?\s*(\S+)`;
  }
  // Header-above anchor (for table layouts)
  let anchorText = null, dx = 0, dy = 0;
  if (!before.length && lineIdx > 0) {
    const above = lines[lineIdx - 1].words;
    if (above.length) {
      const a = above.reduce((p, c) => (Math.abs(c.x - word.x) < Math.abs(p.x - word.x) ? c : p));
      anchorText = a.text; dx = word.x - a.x; dy = word.y - a.y;
    }
  }
  return { label, rx: word.rx, ry: word.ry, anchorText, dx, dy, sample: word.text };
}

/** Apply a user rule to a page's lines. */
export function applyRule(rule, lines) {
  if (!rule) return '';
  if (rule.label) {
    const re = new RegExp(rule.label, 'i');
    for (const l of lines) { const m = l.text.match(re); if (m) return trimValue(m[1]); }
  }
  const words = lines.flatMap((l) => l.words);
  if (!words.length) return '';
  if (rule.anchorText) {
    const anchors = words.filter((w) => w.text === rule.anchorText);
    if (anchors.length) {
      const a = anchors.reduce((p, c) => (dist2(c.rx, c.ry, rule.rx, rule.ry) < dist2(p.rx, p.ry, rule.rx, rule.ry) ? c : p));
      const tx = a.x + rule.dx, ty = a.y + rule.dy;
      const w = words.reduce((p, c) => (dist2(c.x, c.y, tx, ty) < dist2(p.x, p.y, tx, ty) ? c : p));
      if (dist2(w.x, w.y, tx, ty) < 40 * 40) return trimValue(w.text);
    }
  }
  // Fallback: nearest word to the same spot on the page
  const w = words.reduce((p, c) => (dist2(c.rx, c.ry, rule.rx, rule.ry) < dist2(p.rx, p.ry, rule.rx, rule.ry) ? c : p));
  return dist2(w.rx, w.ry, rule.rx, rule.ry) < 0.03 * 0.03 ? trimValue(w.text) : '';
}

/** Apply a custom regex typed by the user (first capture group, or whole match). */
export function applyRegex(src, lines) {
  let re;
  try { re = new RegExp(src, 'i'); } catch { return ''; }
  for (const l of lines) { const m = l.text.match(re); if (m) return trimValue(m[1] ?? m[0]); }
  return '';
}

const dist2 = (x1, y1, x2, y2) => (x1 - x2) ** 2 + (y1 - y2) ** 2;

/* ---------- File names ---------- */

export function safeName(s) {
  return String(s || '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/^[\s.\-]+|[\s.]+$/g, '')
    .slice(0, 120);
}

/** Make final names unique and non-empty. names: array of strings. */
export function finalNames(names, fallbackPrefix = 'Set') {
  const used = new Map();
  const pad = String(names.length).length < 3 ? 3 : String(names.length).length;
  return names.map((n, i) => {
    let base = safeName(n) || `${fallbackPrefix}_${String(i + 1).padStart(pad, '0')}`;
    const key = base.toLowerCase();
    const c = (used.get(key) || 0) + 1;
    used.set(key, c);
    if (c > 1) base = `${base}_${c}`;
    return base + '.pdf';
  });
}

/* ---------- Output ---------- */

/** srcDocs: array of loaded pdf-lib PDFDocument (one per file). */
export async function buildSetPdf(srcDocs, set) {
  const out = await PDFDocument.create();
  for (const { file, page } of set) {
    const [p] = await out.copyPages(srcDocs[file], [page]);
    out.addPage(p);
  }
  return out.save();
}

export async function loadPdfLib(bytes) {
  return PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
}
