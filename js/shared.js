// Shared helpers for all mpdf tools: pdf.js setup, thumbnails, zoom viewer,
// downloads, tab routing and page-wide file drops.
import * as pdfjsLib from '../vendor/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;
const CMAP_URL = new URL('../vendor/cmaps/', import.meta.url).href;

export const $ = (s, root = document) => root.querySelector(s);
export const $$ = (s, root = document) => [...root.querySelectorAll(s)];

export const ZOOM_ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="10.5" cy="10.5" r="6.5"/><path d="m20 20-4.8-4.8M10.5 7.5v6M7.5 10.5h6"/></svg>';
export const ROTATE_ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 12a8 8 0 1 1-2.6-5.9M20 4v5h-5"/></svg>';

export const fmtSize = (b) => (b >= 1048576 ? (b / 1048576).toFixed(2) + ' MB' : (b / 1024).toFixed(2) + ' KB');
export const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
export const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const baseName = (n) => String(n).replace(/\.[^.]+$/, '');
export const naturalCompare = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }).compare;
export const isPdf = (f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name);
export const isImage = (f) => /^image\//.test(f.type) || /\.(jpe?g|png|webp|gif|bmp)$/i.test(f.name);

/** Open a PDF with pdf.js. Throws a friendly Error on failure. */
export async function openPdf(buf) {
  try {
    return await pdfjsLib.getDocument({ data: new Uint8Array(buf.slice(0)), cMapUrl: CMAP_URL, cMapPacked: true, isEvalSupported: false }).promise;
  } catch (err) {
    throw new Error(err?.name === 'PasswordException' ? 'password protected' : 'could not be read');
  }
}
export function closePdf(js) { try { (js?.loadingTask || js)?.destroy?.(); } catch { /* ignore */ } }

/** Render a pdf.js page to a canvas `h` CSS px tall (sharp on hi-dpi). */
export async function renderPageCanvas(js, pageIdx, h) {
  const page = await js.getPage(pageIdx + 1);
  const base = page.getViewport({ scale: 1 });
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  const vp = page.getViewport({ scale: (h * dpr) / base.height });
  const c = document.createElement('canvas');
  c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
  return c;
}
export function cloneCanvas(src) {
  const c = document.createElement('canvas');
  c.width = src.width; c.height = src.height;
  c.getContext('2d').drawImage(src, 0, 0);
  return c;
}

/** Lazy renderer: call observe(el, fn) and fn(el) runs once el is near the viewport. */
export function lazy(rootMargin = '300px') {
  const fns = new WeakMap();
  const io = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) { io.unobserve(e.target); fns.get(e.target)?.(e.target); }
  }, { rootMargin });
  return (el, fn) => { fns.set(el, fn); io.observe(el); };
}

/** Viewer items for pages of pdf.js documents. pages: [{js, page, label}] */
export async function pdfViewerItems(pages) {
  return Promise.all(pages.map(async ({ js, page, label }) => {
    const v = (await js.getPage(page + 1)).getViewport({ scale: 1 });
    return { w: v.width, h: v.height, label, render: (h) => renderPageCanvas(js, page, h) };
  }));
}

/* ---------------- Zoom viewer ---------------- */

const viewer = { items: [], zoom: 1, fit: true, token: 0, obs: null };
const vBody = () => $('#viewer-body');

/** items: [{w, h, label?, render: async (cssHeight) => canvas}] */
export function openViewer(title, sub, items, start = 0) {
  viewer.items = items; viewer.fit = true;
  $('#viewer-title').textContent = title;
  $('#viewer-sub').textContent = sub || '';
  $('#viewer').hidden = false;
  document.body.classList.add('no-scroll');
  vBody().scrollTop = 0;
  layoutViewer(0);
  if (start) { const el = vBody().children[start]; if (el) vBody().scrollTop = el.offsetTop - 24; }
  $('#viewer-close').focus();
}

function closeViewer() {
  $('#viewer').hidden = true;
  document.body.classList.remove('no-scroll');
  viewer.token++; viewer.obs?.disconnect(); vBody().innerHTML = '';
}

function fitZoom() {
  const maxW = Math.max(...viewer.items.map((s) => s.w));
  return Math.max(0.1, Math.min(1.3, (vBody().clientWidth - 48) / maxW));
}

function layoutViewer(anchor) {
  const body = vBody();
  if (viewer.fit) viewer.zoom = fitZoom();
  $('#viewer-zoom').textContent = Math.round(viewer.zoom * 100) + '%';
  const token = ++viewer.token;
  const rel = anchor ?? (body.scrollHeight ? body.scrollTop / body.scrollHeight : 0);
  viewer.obs?.disconnect();
  body.innerHTML = '';
  viewer.obs = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) {
      viewer.obs.unobserve(e.target);
      const it = viewer.items[+e.target.dataset.k];
      it.render(it.h * viewer.zoom).then((c) => { if (token === viewer.token) e.target.prepend(c); }).catch(() => {});
    }
  }, { root: body, rootMargin: '600px 0px' });
  viewer.items.forEach((it, k) => {
    const d = document.createElement('div');
    d.className = 'vpage';
    d.dataset.k = k;
    d.style.width = Math.round(it.w * viewer.zoom) + 'px';
    d.style.height = Math.round(it.h * viewer.zoom) + 'px';
    if (viewer.items.length > 1 || it.label) d.innerHTML = `<span class="vlabel">${viewer.items.length > 1 ? `${k + 1} / ${viewer.items.length}` : ''}${it.label ? ' · ' + esc(it.label) : ''}</span>`;
    body.appendChild(d);
    viewer.obs.observe(d);
  });
  body.scrollTop = rel * body.scrollHeight;
}

function setZoom(z) {
  viewer.fit = false;
  viewer.zoom = Math.max(0.1, Math.min(4, z));
  layoutViewer();
}

function initViewer() {
  $$('.viewer-tools [data-z]').forEach((b) => b.addEventListener('click', () => {
    const z = b.dataset.z;
    if (z === 'in') setZoom(viewer.zoom * 1.25);
    else if (z === 'out') setZoom(viewer.zoom / 1.25);
    else { viewer.fit = true; layoutViewer(); }
  }));
  $('#viewer-close').addEventListener('click', closeViewer);
  $('#viewer').addEventListener('click', (e) => { if (e.target === vBody()) closeViewer(); });
  document.addEventListener('keydown', (e) => {
    if ($('#viewer').hidden) return;
    if (e.key === 'Escape') closeViewer();
    else if (e.key === '+' || e.key === '=') setZoom(viewer.zoom * 1.25);
    else if (e.key === '-') setZoom(viewer.zoom / 1.25);
    else if (e.key === '0') { viewer.fit = true; layoutViewer(); }
  });
  let wheelTimer, wheelZoom = null;
  vBody().addEventListener('wheel', (e) => {
    if (!e.ctrlKey && !e.metaKey) return;
    e.preventDefault();
    wheelZoom = (wheelZoom ?? viewer.zoom) * (e.deltaY < 0 ? 1.1 : 1 / 1.1);
    $('#viewer-zoom').textContent = Math.round(Math.max(0.1, Math.min(4, wheelZoom)) * 100) + '%';
    clearTimeout(wheelTimer);
    wheelTimer = setTimeout(() => { setZoom(wheelZoom); wheelZoom = null; }, 120);
  }, { passive: false });
  window.addEventListener('resize', () => { if (!$('#viewer').hidden && viewer.fit) layoutViewer(); });
}

/* ---------------- Downloads ---------------- */

export function saveBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}

/** Make names unique within a list (adds _2, _3 before the extension). */
export function uniqueNames(names) {
  const used = new Map();
  return names.map((n) => {
    const m = n.match(/^(.*?)(\.[^.]*)?$/);
    const key = n.toLowerCase();
    const c = (used.get(key) || 0) + 1;
    used.set(key, c);
    return c > 1 ? `${m[1]}_${c}${m[2] || ''}` : n;
  });
}

/**
 * Save outputs [{name, bytes}]: one file directly, several as a ZIP.
 * onProgress(0..1) is optional.
 */
export async function saveOutputs(outputs, zipName, onProgress) {
  if (outputs.length === 1) {
    saveBlob(new Blob([outputs[0].bytes], { type: 'application/pdf' }), outputs[0].name);
    return;
  }
  const zip = new window.JSZip();
  const names = uniqueNames(outputs.map((o) => o.name));
  outputs.forEach((o, i) => zip.file(names[i], o.bytes));
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' }, (m) => onProgress?.(m.percent / 100));
  saveBlob(blob, zipName);
}

export const tick = () => new Promise((r) => setTimeout(r));

/* ---------------- Tabs + page-wide drops ---------------- */

const dropHandlers = {};
/** handler(files) receives files dropped anywhere on the page while `tool` is open. */
export function onPageDrop(tool, handler) { dropHandlers[tool] = handler; }
export const activeTool = () => document.body.dataset.tool;

const TITLES = {
  invoices: 'mpdf — split invoice bundles into one PDF per invoice',
  split: 'mpdf — split PDF',
  images: 'mpdf — JPG / PNG to PDF',
};

function showTool() {
  const t = (location.hash.slice(1) || 'invoices');
  const tool = TITLES[t] ? t : 'invoices';
  document.body.dataset.tool = tool;
  $$('[data-tool-panel]').forEach((p) => (p.hidden = p.dataset.toolPanel !== tool));
  $$('.tabs a').forEach((a) => a.classList.toggle('on', a.dataset.tab === tool));
  document.title = TITLES[tool];
}

function initTabsAndDrops() {
  window.addEventListener('hashchange', showTool);
  showTool();
  let depth = 0;
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  window.addEventListener('dragenter', (e) => { if (hasFiles(e)) { depth++; document.body.classList.add('file-over'); } });
  window.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; document.body.classList.remove('file-over'); } });
  window.addEventListener('drop', () => { depth = 0; document.body.classList.remove('file-over'); }, true);
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    if (!e.dataTransfer?.files?.length) return;
    e.preventDefault();
    dropHandlers[activeTool()]?.([...e.dataTransfer.files]);
  });
}

/**
 * Wire a drop zone: a <label class="drop"> containing a hidden file input.
 * Files dropped on it or chosen in it go to onFiles.
 */
export function wireDropZone(zone, input, onFiles) {
  input.addEventListener('change', () => { onFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((t) => zone.addEventListener(t, (e) => { e.preventDefault(); zone.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => zone.addEventListener(t, (e) => { e.preventDefault(); zone.classList.remove('over'); }));
  zone.addEventListener('drop', (e) => { e.stopPropagation(); document.body.classList.remove('file-over'); onFiles([...e.dataTransfer.files]); });
}

/**
 * Reorderable card grid: drag cards to reorder, drop files onto a card to insert there.
 * opts: { move(from, to), insertFiles(files, at) }
 */
export function wireCardGrid(grid, opts) {
  let from = null;
  const clear = () => $$('.drop-before', grid).forEach((x) => x.classList.remove('drop-before'));
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  grid.addEventListener('dragstart', (e) => {
    const card = e.target.closest('.fcard'); if (!card) return;
    from = +card.dataset.i; card.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', '');
  });
  grid.addEventListener('dragover', (e) => {
    if (from === null && !hasFiles(e)) return;
    e.preventDefault(); e.stopPropagation();
    clear();
    e.target.closest('.fcard')?.classList.add('drop-before');
  });
  grid.addEventListener('drop', (e) => {
    const card = e.target.closest('.fcard');
    if (from === null) {
      if (!e.dataTransfer?.files?.length) return;
      e.preventDefault(); e.stopPropagation();
      clear(); document.body.classList.remove('file-over');
      opts.insertFiles([...e.dataTransfer.files], card ? +card.dataset.i : null);
      return;
    }
    e.preventDefault(); e.stopPropagation();
    let to = card ? +card.dataset.i : Infinity;
    const f = from; from = null;
    opts.move(f, to);
  });
  grid.addEventListener('dragend', () => { from = null; clear(); $$('.dragging', grid).forEach((x) => x.classList.remove('dragging')); });
}

/** Move item in array (to = index of the card it was dropped before). */
export function moveItem(arr, from, to) {
  const [m] = arr.splice(from, 1);
  if (to > from) to--;
  arr.splice(Math.min(to, arr.length), 0, m);
}

initViewer();
initTabsAndDrops();
