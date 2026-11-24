// Tool 1: invoice sets — group invoice + DN + PO and name each file by invoice number.
import {
  MODES, suggestMode, buildSets, toLines, autoDetect, makeRule, applyRule, applyRegex,
  finalNames, safeName, buildSetPdf, loadPdfLib,
} from './core.js';
import {
  $, ZOOM_ICON, fmtSize, plural, esc, isPdf, naturalCompare, openPdf, closePdf, renderPageCanvas, cloneCanvas,
  lazy, pdfViewerItems, openViewer, saveBlob, onPageDrop, wireDropZone, wireCardGrid, moveItem,
} from './shared.js';

const state = {
  files: [],          // {id, name, size, buf, pageCount, js (pdf.js doc), lib (pdf-lib doc), thumb}
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
const closeDoc = (f) => closePdf(f.js);

/* ---------------- Files ---------------- */

wireDropZone($('#drop'), $('#file-input'), (files) => addFiles(files));
onPageDrop('invoices', (files) => addFiles(files));
$('#add-more').addEventListener('click', () => $('#file-input').click());

$('#clear').addEventListener('click', () => {
  state.files.forEach(closeDoc);
  state.files = []; state.rule = null; state.modeTouched = false; lineCache.clear();
  refresh();
});
$('#sort-az').addEventListener('click', () => {
  state.files.sort((a, b) => naturalCompare(a.name, b.name));
  refresh();
});

async function addFiles(list, at = null) {
  const pdfs = list.filter(isPdf);
  if (!pdfs.length) return;
  const errors = [];
  const added = [];
  for (const f of pdfs) {
    try {
      const buf = await f.arrayBuffer();
      const js = await openPdf(buf);
      added.push({ id: nextId++, name: f.name, size: f.size, buf, pageCount: js.numPages, js, lib: null });
    } catch (err) {
      errors.push(`${f.name}: ${err.message}`);
    }
  }
  if (at === null || at > state.files.length) state.files.push(...added);
  else state.files.splice(at, 0, ...added);
  if (errors.length) alert('Skipped:\n' + errors.join('\n'));
  refresh();
}

const observeThumb = lazy();

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
      <span class="tip">${fmtSize(f.size)} · ${plural(f.pageCount, 'page')}</span>
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
    else {
      page.innerHTML = '<div class="spinner"></div>';
      observeThumb(page, (el) => renderPageCanvas(f.js, 0, 300).then((c) => { f.thumb = c; el.innerHTML = ''; el.appendChild(cloneCanvas(c)); }).catch(() => {}));
    }
    grid.appendChild(card);
  });
  const pages = state.files.reduce((s, f) => s + f.pageCount, 0);
  $('#file-summary').textContent = `${plural(state.files.length, 'file')} · ${plural(pages, 'page')}`;
}

$('#file-list').addEventListener('click', (e) => {
  const card = e.target.closest('.fcard'); if (!card) return;
  const i = +card.dataset.i;
  const b = e.target.closest('button');
  if (!b || b.dataset.a === 'zoom') { openFileViewer(i); return; }
  const a = b.dataset.a;
  if (a === 'rm') { closeDoc(state.files[i]); state.files.splice(i, 1); }
  if (a === 'up' && i > 0) [state.files[i - 1], state.files[i]] = [state.files[i], state.files[i - 1]];
  if (a === 'down' && i < state.files.length - 1) [state.files[i + 1], state.files[i]] = [state.files[i], state.files[i + 1]];
  refresh();
});
$('#file-list').addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.classList.contains('fcard')) openFileViewer(+e.target.dataset.i);
});
wireCardGrid($('#file-list'), {
  move: (from, to) => { moveItem(state.files, from, to); refresh(); },
  insertFiles: (files, at) => addFiles(files, at),
});

async function openFileViewer(i) {
  const f = state.files[i];
  const items = await pdfViewerItems(Array.from({ length: f.pageCount }, (_, p) => ({ js: f.js, page: p })));
  openViewer(f.name, `${fmtSize(f.size)} · ${plural(f.pageCount, 'page')}`, items);
}
async function openSetViewer(i) {
  const s = state.sets[i];
  const items = await pdfViewerItems(s.map(({ file, page }) => ({ js: state.files[file].js, page, label: state.files[file].name })));
  openViewer(currentNames()[i], `Set ${i + 1} · ${plural(s.length, 'page')}`, items);
}

/* ---------------- Layout ---------------- */

document.querySelectorAll('input[name=mode]').forEach((r) => r.addEventListener('change', () => { state.mode = r.value; state.modeTouched = true; refresh(); }));
const perSetInput = $('#per-set');
perSetInput.addEventListener('input', () => { const v = parseInt(perSetInput.value, 10); if (v >= 1) { state.perSet = Math.min(50, v); refresh(); } });
document.querySelectorAll('#per-set-row .stepper button').forEach((b) => b.addEventListener('click', () => {
  state.perSet = Math.max(1, Math.min(50, state.perSet + +b.dataset.d)); perSetInput.value = state.perSet; refresh();
}));

/* ---------------- Naming ---------------- */

const hints = {
  auto: 'Looks for “Invoice No”, “Invoice #”, “INV” and similar on the first page of each set.',
  pick: 'Best when auto-detect misses — works for any invoice layout with selectable text.',
  regex: 'For power users. Example: <code>Bill\\s*Ref\\s*:?\\s*(\\S+)</code>',
};
document.querySelectorAll('#naming-seg button').forEach((b) => b.addEventListener('click', () => {
  document.querySelectorAll('#naming-seg button').forEach((x) => x.classList.toggle('on', x === b));
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


/* ---------------- Refresh ---------------- */

function refresh() {
  const has = state.files.length > 0;
  $('#drop').classList.toggle('compact', has);
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

