import * as pdfjsLib from '../vendor/pdf.min.mjs';
import {
  MODES, suggestMode, buildSets, toLines, autoDetect, makeRule, applyRule, applyRegex,
  finalNames, safeName, buildSetPdf, loadPdfLib,
} from './core.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;
const CMAP_URL = new URL('../vendor/cmaps/', import.meta.url).href;

const $ = (s) => document.querySelector(s);
const state = {
  files: [],          // {id, name, size, buf, pageCount, js (pdf.js doc), lib (pdf-lib doc)}
  mode: null,
  modeTouched: false,
  perSet: 3,
  method: 'auto',
  rule: null,
  regex: '',
  sets: [],
  detected: [],
  overrides: new Map(),
  token: 0,
};
let nextId = 1;
const lineCache = new Map();

/* ---------------- Files ---------------- */

const input = $('#file-input');
const drop = $('#drop');

input.addEventListener('change', () => { addFiles([...input.files]); input.value = ''; });
$('#add-more').addEventListener('click', () => input.click());
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => addFiles([...e.dataTransfer.files]));
// Allow dropping anywhere on the page
let dragDepth = 0;
window.addEventListener('dragenter', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) { dragDepth++; document.body.classList.add('file-over'); } });
window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('file-over'); } });
window.addEventListener('drop', () => { dragDepth = 0; document.body.classList.remove('file-over'); }, true);
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => { if (!drop.contains(e.target) && e.dataTransfer?.files?.length) { e.preventDefault(); addFiles([...e.dataTransfer.files]); } });

$('#clear').addEventListener('click', () => {
  state.files.forEach(closeDoc);
  state.files = []; state.rule = null; state.modeTouched = false; lineCache.clear();
  refresh();
});
$('#sort-az').addEventListener('click', () => {
  const coll = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  state.files.sort((a, b) => coll.compare(a.name, b.name));
  refresh();
});

async function addFiles(list, at = null) {
  const pdfs = list.filter((f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name));
  if (!pdfs.length) return;
  const errors = [];
  const added = [];
  for (const f of pdfs) {
    try {
      const buf = await f.arrayBuffer();
      const js = await pdfjsLib.getDocument({ data: new Uint8Array(buf.slice(0)), cMapUrl: CMAP_URL, cMapPacked: true, isEvalSupported: false }).promise;
      added.push({ id: nextId++, name: f.name, size: f.size, buf, pageCount: js.numPages, js, lib: null });
    } catch (err) {
      errors.push(`${f.name}: ${err?.name === 'PasswordException' ? 'password protected' : 'could not be read'}`);
    }
  }
  if (at === null || at > state.files.length) state.files.push(...added);
  else state.files.splice(at, 0, ...added);
  if (errors.length) alert('Skipped:\n' + errors.join('\n'));
  refresh();
}

const ZOOM_ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="10.5" cy="10.5" r="6.5"/><path d="m20 20-4.8-4.8M10.5 7.5v6M7.5 10.5h6"/></svg>';

function fmtSize(b) { return b >= 1048576 ? (b / 1048576).toFixed(2) + ' MB' : (b / 1024).toFixed(2) + ' KB'; }

function renderFileList() {
  const grid = $('#file-list');
  grid.innerHTML = '';
  const showTags = state.mode === MODES.FILES || state.mode === MODES.TYPES;
  const k = state.mode === MODES.TYPES ? state.files.length : state.perSet;
  state.files.forEach((f, i) => {
    const card = document.createElement('div');
    card.className = 'fcard';
    card.draggable = true;
    card.dataset.i = i;
    card.tabIndex = 0;
    const tag = showTags ? `t${(i % k) % 3}` : '';
    card.innerHTML = `<span class="pos ${tag}">${i + 1}</span>
      <span class="tip">${fmtSize(f.size)} · ${f.pageCount} page${f.pageCount === 1 ? '' : 's'}</span>
      <span class="acts">
        <button type="button" data-a="rm" aria-label="Remove" title="Remove">✕</button>
        <button type="button" data-a="zoom" aria-label="Preview" title="Preview">${ZOOM_ICON}</button>
        <button type="button" data-a="up" aria-label="Move earlier" title="Move earlier">←</button>
        <button type="button" data-a="down" aria-label="Move later" title="Move later">→</button>
      </span>
      <div class="page"></div>
      <span class="fname" title="${esc(f.name)}">${esc(f.name)}</span>`;
    const page = card.querySelector('.page');
    if (f.thumb) page.appendChild(cloneCanvas(f.thumb));
    else { page.innerHTML = '<div class="spinner"></div>'; fileThumbObserver.observe(page); }
    grid.appendChild(card);
  });
  const pages = state.files.reduce((s, f) => s + f.pageCount, 0);
  $('#file-summary').textContent = `${state.files.length} file${state.files.length === 1 ? '' : 's'} · ${pages} page${pages === 1 ? '' : 's'}`;
}

const fileThumbObserver = new IntersectionObserver((entries) => {
  for (const e of entries) if (e.isIntersecting) {
    fileThumbObserver.unobserve(e.target);
    const i = +e.target.closest('.fcard').dataset.i;
    const f = state.files[i];
    if (!f) continue;
    renderPageCanvas(f, 0, 300).then((c) => { f.thumb = c; if (e.target.isConnected) { e.target.innerHTML = ''; e.target.appendChild(cloneCanvas(c)); } }).catch(() => {});
  }
}, { rootMargin: '300px' });

/** Render a page to a canvas whose height is `h` CSS px (sharp on hi-dpi). */
async function renderPageCanvas(f, pageIdx, h) {
  const page = await f.js.getPage(pageIdx + 1);
  const base = page.getViewport({ scale: 1 });
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  const vp = page.getViewport({ scale: (h * dpr) / base.height });
  const c = document.createElement('canvas');
  c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
  return c;
}
function cloneCanvas(src) {
  const c = document.createElement('canvas');
  c.width = src.width; c.height = src.height;
  c.getContext('2d').drawImage(src, 0, 0);
  return c;
}

$('#file-list').addEventListener('click', (e) => {
  const card = e.target.closest('.fcard'); if (!card) return;
  const i = +card.dataset.i;
  const b = e.target.closest('button');
  if (!b) { openFileViewer(i); return; }
  const a = b.dataset.a;
  if (a === 'zoom') { openFileViewer(i); return; }
  if (a === 'rm') { closeDoc(state.files[i]); state.files.splice(i, 1); }
  if (a === 'up' && i > 0) [state.files[i - 1], state.files[i]] = [state.files[i], state.files[i - 1]];
  if (a === 'down' && i < state.files.length - 1) [state.files[i + 1], state.files[i]] = [state.files[i], state.files[i + 1]];
  refresh();
});
$('#file-list').addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.classList.contains('fcard')) openFileViewer(+e.target.dataset.i);
});
let dragFrom = null;
const clearDropMarks = () => document.querySelectorAll('.drop-before').forEach((x) => x.classList.remove('drop-before'));
$('#file-list').addEventListener('dragstart', (e) => {
  const card = e.target.closest('.fcard'); if (!card) return;
  dragFrom = +card.dataset.i; card.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', '');
});
const isFileDrag = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
$('#file-list').addEventListener('dragover', (e) => {
  if (dragFrom === null && !isFileDrag(e)) return;
  e.preventDefault(); e.stopPropagation();
  clearDropMarks();
  e.target.closest('.fcard')?.classList.add('drop-before');
});
$('#file-list').addEventListener('drop', (e) => {
  if (dragFrom === null) {
    if (!e.dataTransfer?.files?.length) return;
    e.preventDefault(); e.stopPropagation();
    clearDropMarks(); document.body.classList.remove('file-over');
    const c = e.target.closest('.fcard');
    addFiles([...e.dataTransfer.files], c ? +c.dataset.i : null);
    return;
  }
  e.preventDefault(); e.stopPropagation();
  const card = e.target.closest('.fcard');
  let to = card ? +card.dataset.i : state.files.length;
  const [m] = state.files.splice(dragFrom, 1);
  if (to > dragFrom) to--;
  state.files.splice(to, 0, m);
  dragFrom = null; refresh();
});
$('#file-list').addEventListener('dragend', () => { dragFrom = null; clearDropMarks(); document.querySelectorAll('.dragging').forEach((x) => x.classList.remove('dragging')); });

/* ---------------- Viewer (zoom) ---------------- */

const viewer = { pages: [], zoom: 1, fit: true, token: 0, obs: null };
const vBody = $('#viewer-body');

function openFileViewer(i) {
  const f = state.files[i];
  const pages = Array.from({ length: f.pageCount }, (_, p) => ({ file: i, page: p }));
  openViewer(f.name, `${fmtSize(f.size)} · ${f.pageCount} page${f.pageCount === 1 ? '' : 's'}`, pages);
}
function openSetViewer(i) {
  const s = state.sets[i];
  openViewer(currentNames()[i], `Set ${i + 1} · ${s.length} page${s.length === 1 ? '' : 's'}`, s);
}

async function openViewer(title, sub, pages) {
  viewer.pages = pages; viewer.fit = true;
  $('#viewer-title').textContent = title;
  $('#viewer-sub').textContent = sub;
  $('#viewer').hidden = false;
  document.body.classList.add('no-scroll');
  // base sizes (PDF points) for layout
  viewer.sizes = await Promise.all(pages.map(async ({ file, page }) => {
    const p = await state.files[file].js.getPage(page + 1);
    const v = p.getViewport({ scale: 1 });
    return { w: v.width, h: v.height };
  }));
  layoutViewer();
  $('#viewer-close').focus();
}

function closeViewer() {
  $('#viewer').hidden = true;
  document.body.classList.remove('no-scroll');
  viewer.token++; viewer.obs?.disconnect(); vBody.innerHTML = '';
}

function fitZoom() {
  const maxW = Math.max(...viewer.sizes.map((s) => s.w));
  return Math.max(0.25, Math.min(1.3, (vBody.clientWidth - 48) / maxW));
}

function layoutViewer(anchor) {
  if (viewer.fit) viewer.zoom = fitZoom();
  $('#viewer-zoom').textContent = Math.round(viewer.zoom * 100) + '%';
  const token = ++viewer.token;
  // keep the reading position when zooming
  const rel = anchor ?? (vBody.scrollHeight ? vBody.scrollTop / vBody.scrollHeight : 0);
  viewer.obs?.disconnect();
  vBody.innerHTML = '';
  viewer.obs = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) { viewer.obs.unobserve(e.target); drawViewerPage(e.target, token); }
  }, { root: vBody, rootMargin: '600px 0px' });
  viewer.pages.forEach((p, k) => {
    const s = viewer.sizes[k];
    const d = document.createElement('div');
    d.className = 'vpage';
    d.dataset.k = k;
    d.style.width = Math.round(s.w * viewer.zoom) + 'px';
    d.style.height = Math.round(s.h * viewer.zoom) + 'px';
    const label = viewer.pages.length > 1 ? `<span class="vlabel">${k + 1} / ${viewer.pages.length} · ${esc(state.files[p.file].name)}</span>` : '';
    d.innerHTML = label;
    vBody.appendChild(d);
    viewer.obs.observe(d);
  });
  vBody.scrollTop = rel * vBody.scrollHeight;
}

async function drawViewerPage(el, token) {
  const k = +el.dataset.k;
  const { file, page } = viewer.pages[k];
  try {
    const c = await renderPageCanvas(state.files[file], page, viewer.sizes[k].h * viewer.zoom);
    if (token !== viewer.token) return;
    el.prepend(c);
  } catch { /* ignore */ }
}

function setZoom(z) {
  viewer.fit = false;
  viewer.zoom = Math.max(0.25, Math.min(4, z));
  layoutViewer();
}
document.querySelectorAll('.viewer-tools [data-z]').forEach((b) => b.addEventListener('click', () => {
  const z = b.dataset.z;
  if (z === 'in') setZoom(viewer.zoom * 1.25);
  else if (z === 'out') setZoom(viewer.zoom / 1.25);
  else { viewer.fit = true; layoutViewer(); }
}));
$('#viewer-close').addEventListener('click', closeViewer);
$('#viewer').addEventListener('click', (e) => { if (e.target === vBody) closeViewer(); });
document.addEventListener('keydown', (e) => {
  if ($('#viewer').hidden) return;
  if (e.key === 'Escape') closeViewer();
  else if (e.key === '+' || e.key === '=') setZoom(viewer.zoom * 1.25);
  else if (e.key === '-') setZoom(viewer.zoom / 1.25);
  else if (e.key === '0') { viewer.fit = true; layoutViewer(); }
});
let wheelTimer, wheelZoom = null;
vBody.addEventListener('wheel', (e) => {
  if (!e.ctrlKey && !e.metaKey) return;
  e.preventDefault();
  wheelZoom = (wheelZoom ?? viewer.zoom) * (e.deltaY < 0 ? 1.1 : 1 / 1.1);
  $('#viewer-zoom').textContent = Math.round(Math.max(0.25, Math.min(4, wheelZoom)) * 100) + '%';
  clearTimeout(wheelTimer);
  wheelTimer = setTimeout(() => { setZoom(wheelZoom); wheelZoom = null; }, 120);
}, { passive: false });
window.addEventListener('resize', () => { if (!$('#viewer').hidden && viewer.fit) layoutViewer(); });

/* ---------------- Layout ---------------- */

document.querySelectorAll('input[name=mode]').forEach((r) => r.addEventListener('change', () => { state.mode = r.value; state.modeTouched = true; refresh(); }));
const perSetInput = $('#per-set');
perSetInput.addEventListener('input', () => { const v = parseInt(perSetInput.value, 10); if (v >= 1) { state.perSet = Math.min(50, v); refresh(); } });
document.querySelectorAll('.stepper button').forEach((b) => b.addEventListener('click', () => {
  state.perSet = Math.max(1, Math.min(50, state.perSet + +b.dataset.d)); perSetInput.value = state.perSet; refresh();
}));

/* ---------------- Naming ---------------- */

const hints = {
  auto: 'Looks for “Invoice No”, “Invoice #”, “INV” and similar on the first page of each set.',
  pick: 'Best when auto-detect misses — works for any invoice layout with selectable text.',
  regex: 'For power users. Example: <code>Bill\\s*Ref\\s*:?\\s*(\\S+)</code>',
};
document.querySelectorAll('.seg button').forEach((b) => b.addEventListener('click', () => {
  document.querySelectorAll('.seg button').forEach((x) => x.classList.toggle('on', x === b));
  state.method = b.dataset.method;
  $('#method-hint').innerHTML = hints[state.method];
  $('#pick-panel').hidden = state.method !== 'pick';
  $('#regex-panel').hidden = state.method !== 'regex';
  if (state.method === 'pick') renderPicker();
  detectAll();
}));
let regexTimer;
$('#regex').addEventListener('input', (e) => { state.regex = e.target.value; clearTimeout(regexTimer); regexTimer = setTimeout(detectAll, 250); });

async function getLines(fileIdx, page) {
  const f = state.files[fileIdx];
  const key = `${f.id}:${page}`;
  if (!lineCache.has(key)) {
    lineCache.set(key, (async () => {
      const p = await f.js.getPage(page + 1);
      const vp = p.getViewport({ scale: 1 });
      const tc = await p.getTextContent();
      return toLines(tc, vp);
    })());
  }
  return lineCache.get(key);
}

function detectFrom(lines) {
  if (state.method === 'pick') return state.rule ? applyRule(state.rule, lines) : autoDetect(lines);
  if (state.method === 'regex') return state.regex ? applyRegex(state.regex, lines) : autoDetect(lines);
  return autoDetect(lines);
}

async function detectAll() {
  const token = state.token;
  state.overrides.clear();
  for (let i = 0; i < state.sets.length; i++) {
    const first = state.sets[i][0];
    let name = '';
    try { name = detectFrom(await getLines(first.file, first.page)); } catch { /* scanned or broken page */ }
    if (token !== state.token) return;
    state.detected[i] = safeName(name);
    updateRow(i);
  }
  updateStats();
}

async function renderPicker() {
  const stage = $('#pick-stage');
  if (!state.sets.length) return;
  stage.innerHTML = '<div class="spinner"></div>';
  const first = state.sets[0][0];
  const f = state.files[first.file];
  const page = await f.js.getPage(first.page + 1);
  const lines = await getLines(first.file, first.page);
  const base = page.getViewport({ scale: 1 });
  const scale = Math.min(2.5, (Math.min(640, stage.clientWidth || 640) * (window.devicePixelRatio || 1)) / base.width);
  const vp = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = vp.width; canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
  const layer = document.createElement('div');
  layer.className = 'pick-layer';
  layer.style.setProperty('--ar', `${vp.width} / ${vp.height}`);
  let count = 0;
  lines.forEach((l, li) => l.words.forEach((w, wi) => {
    const [px, py] = vp.convertToViewportPoint(w.x, w.y);
    const b = document.createElement('button');
    b.type = 'button';
    b.title = w.text;
    b.dataset.l = li; b.dataset.w = wi;
    const h = w.h * scale;
    Object.assign(b.style, {
      left: (px / vp.width) * 100 + '%',
      top: ((py - h * 0.95) / vp.height) * 100 + '%',
      width: ((Math.max(w.w, 4) * scale) / vp.width) * 100 + '%',
      height: ((h * 1.25) / vp.height) * 100 + '%',
    });
    layer.appendChild(b); count++;
  }));
  stage.innerHTML = '';
  stage.append(canvas, layer);
  if (!count) {
    stage.insertAdjacentHTML('beforeend', '<p class="warn" style="position:absolute;left:12px;right:12px;bottom:12px">This page has no selectable text (probably a scan), so names can’t be read automatically. Type them in the table below.</p>');
  }
  layer.addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    layer.querySelectorAll('.chosen').forEach((x) => x.classList.remove('chosen'));
    b.classList.add('chosen');
    state.rule = makeRule(lines, +b.dataset.l, +b.dataset.w);
    detectAll();
  });
}

/* ---------------- Results table ---------------- */

const thumbObserver = new IntersectionObserver((entries) => {
  for (const e of entries) if (e.isIntersecting) { thumbObserver.unobserve(e.target); renderThumb(e.target); }
}, { rootMargin: '200px' });

async function renderThumb(el) {
  const i = +el.dataset.set;
  const s = state.sets[i]; if (!s) return;
  try {
    const f = state.files[s[0].file];
    const page = await f.js.getPage(s[0].page + 1);
    const base = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: 136 / base.height });
    const c = document.createElement('canvas');
    c.width = vp.width; c.height = vp.height;
    await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    el.innerHTML = ''; el.appendChild(c);
  } catch { /* ignore */ }
}

function renderRows() {
  const tb = $('#rows');
  tb.innerHTML = '';
  state.sets.forEach((s, i) => {
    const tr = document.createElement('tr');
    tr.dataset.i = i;
    const pageDots = s.slice(0, 12).map((p) => `<i class="${state.mode === MODES.TYPES ? 't' + (p.file % 3) : ''}" title="${esc(state.files[p.file].name)} p${p.page + 1}"></i>`).join('');
    tr.innerHTML = `<td>${i + 1}</td>
      <td><div class="thumb" data-set="${i}" title="Preview this set"></div></td>
      <td><div class="name-cell"><span class="badge"></span><input type="text" spellcheck="false" aria-label="File name for set ${i + 1}"><span class="ext">.pdf</span></div></td>
      <td><span class="pages">${pageDots}</span> ${s.length}</td>
      <td><button type="button" class="dl-one" title="Download this file"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11m-5-5 5 5 5-5M5 20h14"/></svg></button></td>`;
    tb.appendChild(tr);
    thumbObserver.observe(tr.querySelector('.thumb'));
  });
}

$('#rows').addEventListener('input', (e) => {
  if (e.target.tagName !== 'INPUT') return;
  const i = +e.target.closest('tr').dataset.i;
  state.overrides.set(i, e.target.value);
  updateBadges(); updateStats();
});
$('#rows').addEventListener('click', async (e) => {
  const th = e.target.closest('.thumb');
  if (th) { openSetViewer(+th.dataset.set); return; }
  const b = e.target.closest('.dl-one'); if (!b) return;
  const i = +b.closest('tr').dataset.i;
  b.disabled = true;
  try {
    await ensureLib();
    const bytes = await buildSetPdf(state.files.map((f) => f.lib), state.sets[i]);
    saveBlob(new Blob([bytes], { type: 'application/pdf' }), currentNames()[i]);
  } finally { b.disabled = false; }
});

const nameOf = (i) => (state.overrides.has(i) ? state.overrides.get(i) : state.detected[i] || '');
const currentNames = () => finalNames(state.sets.map((_, i) => nameOf(i)), 'Invoice');

function updateRow(i) {
  const tr = $(`#rows tr[data-i="${i}"]`); if (!tr) return;
  const inp = tr.querySelector('input');
  if (document.activeElement !== inp) inp.value = nameOf(i);
  inp.placeholder = finalNames(state.sets.map(() => ''), 'Invoice')[i].replace(/\.pdf$/, '');
  updateBadges();
}

function updateBadges() {
  const counts = new Map();
  state.sets.forEach((_, i) => { const k = safeName(nameOf(i)).toLowerCase(); if (k) counts.set(k, (counts.get(k) || 0) + 1); });
  document.querySelectorAll('#rows tr').forEach((tr) => {
    const i = +tr.dataset.i;
    const k = safeName(nameOf(i)).toLowerCase();
    const b = tr.querySelector('.badge');
    if (!k) { b.className = 'badge miss'; b.textContent = '?'; b.title = 'No invoice number found — type one or leave for a numbered name'; }
    else if (counts.get(k) > 1) { b.className = 'badge dup'; b.textContent = '2'; b.title = 'Same name as another set — a suffix will be added'; }
    else { b.className = 'badge ok'; b.textContent = '✓'; b.title = 'Invoice number found'; }
  });
}

function updateStats() {
  const n = state.sets.length;
  const found = state.sets.filter((_, i) => safeName(nameOf(i))).length;
  $('#results-title').textContent = `${n} set${n === 1 ? '' : 's'} → ${n} PDF${n === 1 ? '' : 's'}`;
  $('#found-stat').innerHTML = n ? `<b class="${found < n ? 'bad' : ''}">${found}/${n}</b> invoice numbers found` : '';
  $('#dl-hint').textContent = found < n ? 'Sets without a number will be saved as Invoice_001.pdf, Invoice_002.pdf…' : '';
  $('#download-zip').textContent = n === 1 ? 'Download PDF' : `Download ${n} PDFs as ZIP`;
}

/* ---------------- Download ---------------- */

async function ensureLib() {
  for (const f of state.files) if (!f.lib) f.lib = await loadPdfLib(f.buf.slice(0));
}

$('#download-zip').addEventListener('click', async () => {
  const btn = $('#download-zip');
  const bar = $('#progress-bar');
  btn.disabled = true; $('#progress').hidden = false; bar.style.width = '0%';
  try {
    await ensureLib();
    const names = currentNames();
    const docs = state.files.map((f) => f.lib);
    if (state.sets.length === 1) {
      const bytes = await buildSetPdf(docs, state.sets[0]);
      saveBlob(new Blob([bytes], { type: 'application/pdf' }), names[0]);
    } else {
      const zip = new window.JSZip();
      for (let i = 0; i < state.sets.length; i++) {
        zip.file(names[i], await buildSetPdf(docs, state.sets[i]));
        bar.style.width = `${((i + 1) / state.sets.length) * 85}%`;
        if (i % 5 === 4) await new Promise((r) => setTimeout(r));
      }
      const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' }, (m) => { bar.style.width = `${85 + m.percent * 0.15}%`; });
      saveBlob(blob, `mpdf-${new Date().toISOString().slice(0, 10)}.zip`);
    }
    bar.style.width = '100%';
  } catch (err) {
    console.error(err);
    alert('Something went wrong building the files: ' + (err?.message || err));
  } finally {
    btn.disabled = false;
    setTimeout(() => { $('#progress').hidden = true; }, 1200);
  }
});

function saveBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}

/* ---------------- Refresh ---------------- */

function refresh() {
  const has = state.files.length > 0;
  document.body.classList.toggle('has-files', has);
  $('#file-bar').hidden = !has;
  ['#step-layout', '#step-names', '#step-dl'].forEach((s) => ($(s).hidden = !has));
  if (!has) { $('#file-list').innerHTML = ''; return; }

  if (!state.modeTouched) state.mode = suggestMode(state.files, state.perSet);
  document.querySelectorAll('input[name=mode]').forEach((r) => (r.checked = r.value === state.mode));

  const types = state.mode === MODES.TYPES;
  $('#per-set-row').style.display = types ? 'none' : '';
  $('#per-set-label').textContent = state.mode === MODES.PAGES ? 'Pages in each set' : 'Files in each set';
  $('#per-set-hint').textContent = state.mode === MODES.PAGES ? 'Invoice + DN + PO, one page each = 3' : 'Invoice + delivery note + PO = 3';

  renderFileList();
  const { sets, warnings } = buildSets(state.files, state.mode, state.perSet);
  if (types) warnings.unshift(`Set N = page N of each file, in this order: ${state.files.map((f) => f.name).join(' → ')}. The invoice file should be first.`);
  state.sets = sets;
  state.detected = new Array(sets.length).fill('');
  state.token++;
  const w = $('#warnings');
  w.hidden = !warnings.length;
  w.innerHTML = warnings.map((x) => `<p>${esc(x)}</p>`).join('');

  renderRows();
  updateStats();
  if (state.method === 'pick') renderPicker();
  detectAll();
}

function closeDoc(f) { try { (f.js?.loadingTask || f.js)?.destroy?.(); } catch { /* ignore */ } }
function esc(s) { return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
