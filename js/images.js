// Tool 3: JPG / PNG to PDF — order, rotate, page size, orientation, margin, merge.
import { layoutImage, imagesToPdfs, imageKind, jpegOrientation } from './core.js';
import {
  $, $$, ZOOM_ICON, ROTATE_ICON, fmtSize, plural, esc, baseName, isImage, naturalCompare,
  openViewer, saveOutputs, onPageDrop, wireDropZone, wireCardGrid, moveItem, tick,
} from './shared.js';

const st = {
  images: [],   // {id, name, size, file, w, h (upright px), rot, thumb (upright canvas)}
  opts: { orientation: 'auto', size: 'a4', margin: 'none', merge: true },
};
let nextId = 1;

/* ---------- Adding images ---------- */

wireDropZone($('#im-drop'), $('#im-input'), (files) => addImages(files));
onPageDrop('images', (files) => addImages(files));
$('#im-add').addEventListener('click', () => $('#im-input').click());
$('#im-clear').addEventListener('click', () => { st.images = []; render(); });
$('#im-sort').addEventListener('click', () => { st.images.sort((a, b) => naturalCompare(a.name, b.name)); render(); });

async function decode(file) {
  try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch {
    // Fallback through an <img> element
    const url = URL.createObjectURL(file);
    try {
      const img = new Image(); img.src = url; await img.decode();
      return await createImageBitmap(img);
    } finally { URL.revokeObjectURL(url); }
  }
}

function drawRotated(src, sw, sh, rot, maxSide = Infinity, bg = null) {
  const s = Math.min(1, maxSide / Math.max(sw, sh));
  const w = Math.round(sw * s), h = Math.round(sh * s);
  const c = document.createElement('canvas');
  const side = rot % 180 !== 0;
  c.width = side ? h : w; c.height = side ? w : h;
  const ctx = c.getContext('2d');
  if (bg) { ctx.fillStyle = bg; ctx.fillRect(0, 0, c.width, c.height); }
  ctx.translate(c.width / 2, c.height / 2);
  ctx.rotate((rot * Math.PI) / 180);
  ctx.drawImage(src, -w / 2, -h / 2, w, h);
  return c;
}

async function addImages(list, at = null) {
  const files = list.filter(isImage);
  const skipped = list.filter((f) => !isImage(f)).map((f) => f.name);
  const added = [];
  for (const file of files) {
    try {
      const bmp = await decode(file);
      const thumb = drawRotated(bmp, bmp.width, bmp.height, 0, 420);
      added.push({ id: nextId++, name: file.name, size: file.size, file, w: bmp.width, h: bmp.height, rot: 0, thumb });
      bmp.close?.();
    } catch { skipped.push(file.name); }
  }
  if (at === null || at > st.images.length) st.images.push(...added);
  else st.images.splice(at, 0, ...added);
  if (skipped.length) alert('Skipped (not a JPG/PNG image or unreadable):\n' + skipped.join('\n'));
  render();
}

/* ---------- Options ---------- */

function wireSeg(id, key) {
  $$(`#${id} button`).forEach((b) => b.addEventListener('click', () => {
    $$(`#${id} button`).forEach((x) => x.classList.toggle('on', x === b));
    st.opts[key] = b.dataset.v;
    $('#im-orientation').classList.toggle('disabled', st.opts.size === 'fit');
    render();
  }));
}
wireSeg('im-orientation', 'orientation');
wireSeg('im-size', 'size');
wireSeg('im-margin', 'margin');
$('#im-merge').addEventListener('change', (e) => { st.opts.merge = e.target.checked; updateButton(); });

/* ---------- Grid ---------- */

const dims = (im) => (im.rot % 180 ? [im.h, im.w] : [im.w, im.h]);

/** Page preview: white page box with the image placed exactly as in the PDF. */
function pagePreview(im) {
  const [iw, ih] = dims(im);
  const L = layoutImage(iw, ih, st.opts);
  const box = 3 / 4, pa = L.pw / L.ph;
  const bw = pa > box ? 100 : (pa / box) * 100;
  const bh = pa > box ? (box / pa) * 100 : 100;
  const page = document.createElement('div');
  page.className = 'ipage';
  page.style.width = bw + '%';
  page.style.height = bh + '%';
  const c = drawRotated(im.thumb, im.thumb.width, im.thumb.height, im.rot);
  Object.assign(c.style, {
    left: (L.x / L.pw) * 100 + '%',
    top: ((L.ph - L.y - L.h) / L.ph) * 100 + '%',
    width: (L.w / L.pw) * 100 + '%',
    height: (L.h / L.ph) * 100 + '%',
  });
  page.appendChild(c);
  return page;
}

function renderGrid() {
  const grid = $('#im-list');
  grid.innerHTML = '';
  st.images.forEach((im, i) => {
    const card = document.createElement('div');
    card.className = 'fcard';
    card.draggable = true;
    card.dataset.i = i;
    card.tabIndex = 0;
    card.innerHTML = `<span class="pos">${i + 1}</span>
      <span class="tip">${fmtSize(im.size)} · ${dims(im).join(' × ')} px</span>
      <span class="acts">
        <button type="button" data-a="rm" aria-label="Remove" title="Remove">✕</button>
        <button type="button" data-a="zoom" aria-label="Preview" title="Preview">${ZOOM_ICON}</button>
        <button type="button" data-a="rot" aria-label="Rotate" title="Rotate 90°">${ROTATE_ICON}</button>
        <button type="button" data-a="up" aria-label="Move earlier" title="Move earlier">←</button>
        <button type="button" data-a="down" aria-label="Move later" title="Move later">→</button>
      </span>
      <div class="page"></div>
      <span class="fname" title="${esc(im.name)}">${esc(im.name)}</span>`;
    card.querySelector('.page').appendChild(pagePreview(im));
    grid.appendChild(card);
  });
}

$('#im-list').addEventListener('click', (e) => {
  const card = e.target.closest('.fcard'); if (!card) return;
  const i = +card.dataset.i;
  const a = e.target.closest('button')?.dataset.a;
  if (!a || a === 'zoom') { openImageViewer(i); return; }
  if (a === 'rm') st.images.splice(i, 1);
  if (a === 'rot') st.images[i].rot = (st.images[i].rot + 90) % 360;
  if (a === 'up' && i > 0) [st.images[i - 1], st.images[i]] = [st.images[i], st.images[i - 1]];
  if (a === 'down' && i < st.images.length - 1) [st.images[i + 1], st.images[i]] = [st.images[i], st.images[i + 1]];
  render();
});
$('#im-list').addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.classList.contains('fcard')) openImageViewer(+e.target.dataset.i);
});
wireCardGrid($('#im-list'), {
  move: (from, to) => { moveItem(st.images, from, to); render(); },
  insertFiles: (files, at) => addImages(files, at),
});

/** Viewer shows each finished PDF page (page size, margins, rotation applied). */
function openImageViewer(start) {
  const items = st.images.map((im) => {
    const [iw, ih] = dims(im);
    const L = layoutImage(iw, ih, st.opts);
    return {
      w: L.pw, h: L.ph, label: im.name,
      render: async (cssH) => {
        const dpr = Math.min(2, window.devicePixelRatio || 1);
        const s = (cssH * dpr) / L.ph;
        const c = document.createElement('canvas');
        c.width = Math.round(L.pw * s); c.height = Math.round(L.ph * s);
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
        const bmp = await decode(im.file);
        const r = drawRotated(bmp, bmp.width, bmp.height, im.rot, Math.max(L.w, L.h) * s * 1.2);
        bmp.close?.();
        ctx.drawImage(r, L.x * s, (L.ph - L.y - L.h) * s, L.w * s, L.h * s);
        return c;
      },
    };
  });
  openViewer(st.opts.merge ? 'PDF preview' : st.images[start].name, `${plural(st.images.length, 'page')} · ${st.opts.size === 'fit' ? 'same size as image' : st.opts.size.toUpperCase()}`, items, start);
}

/* ---------- Convert ---------- */

const MAX_PIXELS = 16e6; // stay under Safari's canvas limit

/** Get upright bytes pdf-lib can embed, keeping the original file when possible. */
async function prepare(im) {
  const bytes = new Uint8Array(await im.file.arrayBuffer());
  const kind = imageKind(bytes);
  if (kind === 'jpg' && im.rot === 0 && jpegOrientation(bytes) === 1) return { kind, bytes, w: im.w, h: im.h };
  if (kind === 'png' && im.rot === 0) return { kind, bytes, w: im.w, h: im.h };
  // Re-encode: applies EXIF rotation + user rotation; converts other formats
  const bmp = await decode(im.file);
  const scale = Math.min(1, Math.sqrt(MAX_PIXELS / (bmp.width * bmp.height)));
  const asPng = kind === 'png';
  const c = drawRotated(bmp, bmp.width, bmp.height, im.rot, Math.max(bmp.width, bmp.height) * scale, asPng ? null : '#fff');
  bmp.close?.();
  const blob = await new Promise((r) => c.toBlob(r, asPng ? 'image/png' : 'image/jpeg', 0.92));
  return { kind: asPng ? 'png' : 'jpg', bytes: new Uint8Array(await blob.arrayBuffer()), w: c.width, h: c.height };
}

$('#im-go').addEventListener('click', async () => {
  if (!st.images.length) return;
  const btn = $('#im-go');
  const bar = $('#im-progress div');
  btn.disabled = true; $('#im-progress').hidden = false; bar.style.width = '0%';
  try {
    const prepared = [];
    for (let i = 0; i < st.images.length; i++) {
      prepared.push(await prepare(st.images[i]));
      bar.style.width = `${((i + 1) / st.images.length) * 50}%`;
      await tick();
    }
    const pdfs = await imagesToPdfs(prepared, st.opts, (i) => { bar.style.width = `${50 + ((i + 1) / prepared.length) * 40}%`; });
    const outputs = st.opts.merge
      ? [{ name: st.images.length === 1 ? `${baseName(st.images[0].name)}.pdf` : `images_${new Date().toISOString().slice(0, 10)}.pdf`, bytes: pdfs[0] }]
      : pdfs.map((bytes, i) => ({ name: `${baseName(st.images[i].name)}.pdf`, bytes }));
    await saveOutputs(outputs, `images_${new Date().toISOString().slice(0, 10)}.zip`, (f) => { bar.style.width = `${90 + f * 10}%`; });
    bar.style.width = '100%';
  } catch (err) {
    console.error(err);
    alert('Something went wrong making the PDF: ' + (err?.message || err));
  } finally {
    btn.disabled = false;
    setTimeout(() => { $('#im-progress').hidden = true; }, 1200);
  }
});

function updateButton() {
  const n = st.images.length;
  $('#im-go').textContent = !n ? 'Convert to PDF' : st.opts.merge || n === 1 ? `Convert to PDF (${plural(n, 'page')})` : `Convert to ${n} PDFs (ZIP)`;
}

function render() {
  const has = st.images.length > 0;
  $('#im-bar').hidden = !has;
  ['#im-options', '#im-dl'].forEach((s) => ($(s).hidden = !has));
  $('#im-drop').classList.toggle('compact', has);
  $('#im-summary').textContent = has ? `${plural(st.images.length, 'image')} · ${fmtSize(st.images.reduce((s, x) => s + x.size, 0))}` : '';
  renderGrid();
  updateButton();
}
