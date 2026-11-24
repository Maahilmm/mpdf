// Tool 2: Split PDF — custom ranges, fixed ranges, extract all / selected pages.
import { computeSplit, parsePageList, pagesLabel, buildPagesPdf, loadPdfLib } from './core.js';
import {
  $, $$, ZOOM_ICON, fmtSize, plural, esc, baseName, isPdf, openPdf, closePdf, renderPageCanvas,
  lazy, pdfViewerItems, openViewer, saveOutputs, onPageDrop, wireDropZone, tick,
} from './shared.js';

const st = {
  file: null,            // {name, size, buf, js, lib, pageCount, thumbs: []}
  mode: 'range',
  rangeType: 'custom',
  ranges: [],
  every: 1,
  extractType: 'all',
  select: '',
  merge: false,
  result: { outputs: [], errors: [] },
};

/* ---------- File ---------- */

wireDropZone($('#sp-drop'), $('#sp-input'), loadFile);
onPageDrop('split', loadFile);
$('#sp-clear').addEventListener('click', () => { if (st.file) closePdf(st.file.js); st.file = null; render(); });

async function loadFile(files) {
  const f = files.find(isPdf);
  if (!f) { if (files.length) alert('Please choose a PDF file.'); return; }
  if (files.filter(isPdf).length > 1) alert(`Split works on one PDF at a time — using “${f.name}”.`);
  try {
    const buf = await f.arrayBuffer();
    const js = await openPdf(buf);
    if (st.file) closePdf(st.file.js);
    st.file = { name: f.name, size: f.size, buf, js, lib: null, pageCount: js.numPages, thumbs: [] };
  } catch (err) { alert(`${f.name}: ${err.message}`); return; }
  const n = st.file.pageCount;
  st.ranges = [{ from: 1, to: Math.min(n, Math.max(1, Math.ceil(n / 2))) }];
  if (n > 1 && st.ranges[0].to < n) st.ranges.push({ from: st.ranges[0].to + 1, to: n });
  st.every = Math.min(st.every, n) || 1;
  $('#sp-every').value = st.every;
  st.select = '';
  $('#sp-select').value = '';
  renderRanges();
  render();
}

/* ---------- Options ---------- */

function wireSeg(id, key, after) {
  $$(`#${id} button`).forEach((b) => b.addEventListener('click', () => {
    $$(`#${id} button`).forEach((x) => x.classList.toggle('on', x === b));
    st[key] = b.dataset.v;
    after?.();
    update();
  }));
}
wireSeg('sp-mode', 'mode');
wireSeg('sp-range-type', 'rangeType');
wireSeg('sp-extract-type', 'extractType');

$('#sp-merge').addEventListener('change', (e) => { st.merge = e.target.checked; update(); });
$('#sp-select').addEventListener('input', (e) => { st.select = e.target.value; update(); });

const every = $('#sp-every');
every.addEventListener('input', () => { const v = parseInt(every.value, 10); if (v >= 1) { st.every = v; update(); } });
$$('#sp-every-stepper button').forEach((b) => b.addEventListener('click', () => {
  const max = st.file?.pageCount || 1;
  st.every = Math.max(1, Math.min(max, st.every + +b.dataset.d));
  every.value = st.every; update();
}));

function renderRanges() {
  const box = $('#sp-ranges');
  const n = st.file?.pageCount || 1;
  box.innerHTML = st.ranges.map((r, i) => `
    <div class="range-row" data-i="${i}">
      <span class="range-name"><i class="g${i % 6}"></i>Range ${i + 1}</span>
      <label>from page <input type="number" min="1" max="${n}" value="${r.from}" data-k="from" inputmode="numeric"></label>
      <label>to <input type="number" min="1" max="${n}" value="${r.to}" data-k="to" inputmode="numeric"></label>
      <button type="button" class="rm" aria-label="Remove range" ${st.ranges.length === 1 ? 'disabled' : ''}>✕</button>
    </div>`).join('');
}
$('#sp-ranges').addEventListener('input', (e) => {
  const row = e.target.closest('.range-row'); if (!row) return;
  st.ranges[+row.dataset.i][e.target.dataset.k] = e.target.value;
  update();
});
$('#sp-ranges').addEventListener('click', (e) => {
  if (!e.target.closest('.rm')) return;
  st.ranges.splice(+e.target.closest('.range-row').dataset.i, 1);
  renderRanges(); update();
});
$('#sp-add-range').addEventListener('click', () => {
  const n = st.file?.pageCount || 1;
  const last = st.ranges[st.ranges.length - 1];
  const from = Math.min(n, (+last?.to || 0) + 1);
  st.ranges.push({ from, to: n });
  renderRanges(); update();
});

/* ---------- Page grid ---------- */

const observe = lazy('400px');

function renderPages() {
  const grid = $('#sp-pages');
  grid.innerHTML = '';
  if (!st.file) return;
  const { js, pageCount } = st.file;
  for (let p = 0; p < pageCount; p++) {
    const cell = document.createElement('div');
    cell.className = 'pcell';
    cell.dataset.p = p;
    cell.innerHTML = `<div class="pthumb"><div class="spinner"></div></div>
      <button type="button" class="pzoom" aria-label="Preview page ${p + 1}" title="Preview">${ZOOM_ICON}</button>
      <span class="pcheck" aria-hidden="true">✓</span>
      <span class="pchip"></span>
      <span class="pnum">${p + 1}</span>`;
    grid.appendChild(cell);
    const th = cell.querySelector('.pthumb');
    if (st.file.thumbs[p]) { th.innerHTML = ''; th.appendChild(st.file.thumbs[p]); }
    else observe(th, (el) => renderPageCanvas(js, p, 190).then((c) => { st.file.thumbs[p] = c; el.innerHTML = ''; el.appendChild(c); }).catch(() => {}));
  }
}

$('#sp-pages').addEventListener('click', async (e) => {
  const cell = e.target.closest('.pcell'); if (!cell) return;
  const p = +cell.dataset.p;
  const selecting = st.mode === 'extract' && st.extractType === 'select';
  if (e.target.closest('.pzoom') || !selecting) {
    const items = await pdfViewerItems(Array.from({ length: st.file.pageCount }, (_, i) => ({ js: st.file.js, page: i })));
    openViewer(st.file.name, `${fmtSize(st.file.size)} · ${plural(st.file.pageCount, 'page')}`, items, p);
    return;
  }
  // toggle page in the selection
  const set = new Set(parsePageList(st.select, st.file.pageCount).pages);
  set.has(p) ? set.delete(p) : set.add(p);
  st.select = pagesLabel([...set].sort((a, b) => a - b));
  $('#sp-select').value = st.select;
  update();
});

function paintPages() {
  const { outputs } = st.result;
  const groupOf = new Map();
  outputs.forEach((o, gi) => o.pages.forEach((p) => { if (!groupOf.has(p)) groupOf.set(p, gi); }));
  const selecting = st.mode === 'extract' && st.extractType === 'select';
  const showGroups = st.mode === 'range';
  $$('#sp-pages .pcell').forEach((cell) => {
    const p = +cell.dataset.p;
    const g = groupOf.get(p);
    cell.className = 'pcell' + (g === undefined ? ' out' : '') + (showGroups && g !== undefined ? ` g${g % 6}` : '') + (selecting ? ' selectable' : '') + (selecting && g !== undefined ? ' picked' : '');
    const chip = cell.querySelector('.pchip');
    const first = g !== undefined && outputs[g].pages[0] === p;
    chip.textContent = first && outputs.length > 1 && showGroups ? `File ${g + 1}` : '';
  });
}

/* ---------- Summary + download ---------- */

function outputName(o) {
  const base = baseName(st.file.name);
  if (o.pages.length === 1) return `${base}_page-${o.pages[0] + 1}.pdf`;
  return `${base}_${o.label.replace(/,/g, '_')}.pdf`;
}

function update() {
  if (!st.file) return;
  const n = st.file.pageCount;
  $('#sp-range').hidden = st.mode !== 'range';
  $('#sp-extract').hidden = st.mode !== 'extract';
  $('#sp-custom').hidden = st.rangeType !== 'custom';
  $('#sp-fixed').hidden = st.rangeType !== 'fixed';
  $('#sp-select-box').hidden = st.extractType !== 'select';
  $('#sp-every').max = n;

  // merge option only makes sense where several files would be made
  const mergeable = (st.mode === 'range' && st.rangeType === 'custom') || (st.mode === 'extract' && st.extractType === 'select');
  $('#sp-merge-row').hidden = !mergeable;
  $('#sp-merge-label').textContent = st.mode === 'range' ? 'Merge all ranges into one PDF' : 'Merge extracted pages into one PDF';

  st.result = computeSplit(n, { ...st, merge: mergeable && st.merge });
  const { outputs, errors } = st.result;
  const err = $('#sp-errors');
  err.hidden = !errors.length;
  err.innerHTML = errors.map((x) => `<p>${esc(x)}</p>`).join('');

  const pagesOut = outputs.reduce((s, o) => s + o.pages.length, 0);
  $('#sp-pages-title').textContent = plural(n, 'page');
  $('#sp-pages-hint').textContent = st.mode === 'extract' && st.extractType === 'select' ? 'Click pages to select them' : 'Click a page to zoom in';
  const go = $('#sp-go');
  go.disabled = !outputs.length;
  go.textContent = outputs.length === 1 ? `Download 1 PDF (${plural(pagesOut, 'page')})` : outputs.length ? `Download ${outputs.length} PDFs as ZIP` : 'Nothing selected';
  $('#sp-dl-hint').textContent = outputs.length > 1 ? `${outputName(outputs[0])}, ${outputName(outputs[1])}${outputs.length > 2 ? '…' : ''}` : outputs[0] ? outputName(outputs[0]) : '';
  paintPages();
}

$('#sp-go').addEventListener('click', async () => {
  const btn = $('#sp-go');
  const bar = $('#sp-progress div');
  const { outputs } = st.result;
  if (!outputs.length) return;
  btn.disabled = true; $('#sp-progress').hidden = false; bar.style.width = '0%';
  try {
    if (!st.file.lib) st.file.lib = await loadPdfLib(st.file.buf.slice(0));
    const files = [];
    for (let i = 0; i < outputs.length; i++) {
      files.push({ name: outputName(outputs[i]), bytes: await buildPagesPdf(st.file.lib, outputs[i].pages) });
      bar.style.width = `${((i + 1) / outputs.length) * 85}%`;
      if (i % 5 === 4) await tick();
    }
    await saveOutputs(files, `${baseName(st.file.name)}_split.zip`, (f) => { bar.style.width = `${85 + f * 15}%`; });
    bar.style.width = '100%';
  } catch (err) {
    console.error(err);
    alert('Something went wrong splitting the PDF: ' + (err?.message || err));
  } finally {
    btn.disabled = false;
    setTimeout(() => { $('#sp-progress').hidden = true; }, 1200);
  }
});

function render() {
  const has = !!st.file;
  $('#sp-bar').hidden = !has;
  ['#sp-options', '#sp-pages-step', '#sp-dl'].forEach((s) => ($(s).hidden = !has));
  $('#sp-drop-label').textContent = has ? 'Drop a different PDF' : 'Drop a PDF here';
  $('#sp-drop').classList.toggle('compact', has);
  if (!has) { $('#sp-pages').innerHTML = ''; return; }
  $('#sp-summary').innerHTML = `<strong>${esc(st.file.name)}</strong> · ${plural(st.file.pageCount, 'page')} · ${fmtSize(st.file.size)}`;
  renderPages();
  update();
}
