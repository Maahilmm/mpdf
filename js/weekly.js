// Weekly pendings report: four Excel files in → weekly report + worked Pending and Clearing files out.
// The processing lives in weekly-core.js (window.WR); this file is the page.
import { $, $$, esc, fmtSize, saveBlob, wireDropZone, onPageDrop } from './shared.js';

const WR = window.WR, JSZip = window.JSZip;
const store = {
  get(k, d) { try { const v = localStorage.getItem('wr.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem('wr.' + k, JSON.stringify(v)); } catch { /* ignore */ } },
};
const MMA = ['AUD', 'CAD', 'DKK', 'EUR', 'HKD', 'JPY', 'NOK', 'SGD', 'SAR', 'GBP', 'CHF', 'SEK', 'LKR', 'INR', 'THB', 'MYR', 'IDR', 'AED', 'CNY'];
const SLOTS = [
  { k: 'open', t: 'Pending items', d: 'SAP open items export' },
  { k: 'cleared', t: 'Cleared items', d: 'SAP clearing export' },
  { k: 'billentry', t: 'Bill Entry', d: 'BILL ENTRY workbook (IGMH + VENDORS)' },
  { k: 'report', t: "Last week's report", d: 'Final Pending as of …' },
];
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const files = {};
let pickFor = null, an = null, built = null;
const now = new Date();
const today = WR.utcDay(now.getFullYear(), now.getMonth(), now.getDate());
const fmt2 = (n) => (n ?? 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const isoD = (d) => d.toISOString().slice(0, 10);
const dmy = (d) => isoD(d).split('-').reverse().join('/');

function showMsg(id, text, bad) { $(id).innerHTML = text ? `<div class="${bad ? 'warn wk-bad' : 'warn wk-note'}"><p>${esc(text)}</p></div>` : ''; }
function resetLater() { an = null; built = null; $('#wk-s2').hidden = true; $('#wk-s3').hidden = true; }

function renderTray() {
  $('#wk-tray').innerHTML = SLOTS.map((s) => {
    const f = files[s.k];
    const cls = f ? (f.error ? 'wk-slot err' : 'wk-slot done') : 'wk-slot';
    const st = f ? (f.error ? esc(f.error) : 'Ready') : 'Empty';
    return `<button type="button" class="${cls}" data-k="${s.k}"><b>${s.t}</b><span class="f">${esc(f ? f.name : s.d)}</span><span class="st">${st}</span></button>`;
  }).join('');
  $('#wk-check').disabled = !SLOTS.every((s) => files[s.k] && files[s.k].det && !files[s.k].error);
}
async function addFiles(list, forced) {
  for (const file of list) {
    if (!/\.xls[xm]$/i.test(file.name)) { showMsg('#wk-m1', `${file.name}: only Excel .xlsx files can be used here.`, true); continue; }
    try {
      const det = await WR.detect(JSZip, await file.arrayBuffer());
      const k = forced || det.kind;
      if (k === 'unknown') { showMsg('#wk-m1', `${file.name}: this doesn't look like any of the four files.`, true); continue; }
      if (forced && det.kind !== forced && det.kind !== 'unknown') showMsg('#wk-m1', `${file.name} looks like the ${SLOTS.find((s) => s.k === det.kind).t} file. It was put in the slot you chose anyway.`, false);
      files[k] = { name: file.name, det };
    } catch (e) {
      if (forced) files[forced] = { name: file.name, error: 'Could not read this file' };
      showMsg('#wk-m1', `${file.name}: ${e.message}`, true);
    }
  }
  resetLater(); renderTray();
}

/* ---------- step 1 ---------- */
const picker = $('#wk-input');
wireDropZone($('#wk-drop'), picker, (fl) => { const k = pickFor; pickFor = null; picker.multiple = true; addFiles(fl, k); });
onPageDrop('weekly', (fl) => addFiles(fl, null));
$('#wk-tray').addEventListener('click', (e) => { const s = e.target.closest('.wk-slot'); if (!s) return; pickFor = s.dataset.k; picker.multiple = false; picker.click(); });
$('#wk-tray').addEventListener('dragover', (e) => e.preventDefault());
$('#wk-tray').addEventListener('drop', (e) => {
  const s = e.target.closest('.wk-slot'); if (!s || !e.dataTransfer?.files?.length) return;
  e.preventDefault(); e.stopPropagation(); document.body.classList.remove('file-over');
  addFiles([...e.dataTransfer.files], s.dataset.k);
});

$('#wk-end').value = isoD(WR.defaultPeriodEnd(today));
$('#wk-arate').value = store.get('ahmedRate', '15.42');
$('#wk-aacct').value = store.get('ahmedAcct', 'A060835');
$('#wk-excl').value = store.get('clearExclude', WR.CLEAR_EXCLUDE.join(', '));
function periodEnd() { const v = $('#wk-end').value; if (!v) return null; const [y, m, d] = v.split('-').map(Number); return WR.utcDay(y, m - 1, d); }
function showSpan() {
  const e = periodEnd(); if (!e) { $('#wk-span').value = ''; return; }
  const s = new Date(e.getTime() - 6 * WR.DAY);
  $('#wk-span').value = WR.dotDate(s) + ' to ' + WR.dotDate(e) + ', aged at ' + WR.dotDate(new Date(e.getTime() + WR.DAY));
  $('#wk-fname').value = WR.fileName(e);
}
$('#wk-end').addEventListener('change', () => { showSpan(); resetLater(); });
showSpan();
function opts() {
  return {
    today, periodEnd: periodEnd(), ahmedRate: parseFloat($('#wk-arate').value) || 15.42,
    ahmedAccounts: $('#wk-aacct').value.split(/[,\s]+/).filter(Boolean),
    clearExclude: $('#wk-excl').value.split(',').map((x) => x.trim()).filter(Boolean),
  };
}

/* ---------- step 2 ---------- */
$('#wk-check').addEventListener('click', async () => {
  if (!periodEnd()) { showMsg('#wk-m1', 'Pick the report period end date first.', true); return; }
  store.set('ahmedRate', $('#wk-arate').value); store.set('ahmedAcct', $('#wk-aacct').value); store.set('clearExclude', $('#wk-excl').value);
  const btn = $('#wk-check'); btn.disabled = true; btn.textContent = 'Checking…'; showMsg('#wk-m1', '');
  try {
    const f = {}; for (const s of SLOTS) f[s.k] = files[s.k].det;
    an = await WR.analyse(JSZip, f, opts());
    renderIssues();
  } catch (e) { showMsg('#wk-m1', 'Something went wrong while reading the files: ' + e.message, true); }
  btn.disabled = false; btn.textContent = 'Check files';
});

const cell = (v) => (v instanceof Date ? dmy(v) : typeof v === 'number' ? (Number.isInteger(v) ? String(v) : fmt2(v)) : (v ?? ''));
const table = (head, rows) => `<div class="table-wrap wk-table"><table class="results"><thead><tr>${head.map((h) => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div>`;

function renderIssues() {
  const I = an.issues, st = an.stats, savedV = store.get('vendorNames', {}), savedR = store.get('rates', {});
  let h = `<div class="wk-stats"><span><b>${st.openItems}</b> pending items (${st.openDropped} subtotal rows dropped)</span><span><b>${st.clearedItems}</b> cleared items, <b>${st.clearedExcluded}</b> payroll, salary or dividend setoff lines left out</span><span><b>${st.keptBills}</b> Bill Entry pendings kept, <b>${st.excludedBills}</b> received on or after the period end left out</span></div>`;
  const ex = an.data.cleared.excluded;
  if (ex.length) {
    h += `<details class="wk-box"><summary>Cleared lines left out of Weekly Clear (${ex.length})</summary>` +
      table(['Account', 'Document', 'Type', 'Amount', 'Currency', 'Text'], ex.map((x) => `<tr><td>${esc(x.acct)}</td><td>${esc(x.docNo)}</td><td>${esc(x.docType)}</td><td class="wk-num">${fmt2(x.amt)}</td><td>${esc(x.cur)}</td><td>${esc(x.text)}</td></tr>`).join('')) + '</details>';
  }
  const P = an.data.prev;
  if (P && P.mvr && P.usd && P.mvr.rows.grand) {
    const g = (k) => fmt2((P[k].rows.grand || {})['Grand Pending']);
    h += `<div class="warn wk-note"><p>Last week's report (${P.periodEnd ? esc(WR.longDate(P.periodEnd)) : 'date not found'}): Grand Pending MVR <b>${g('mvr')}</b>, USD <b>${g('usd')}</b>.` +
      ['mvr', 'usd'].filter((k) => P[k].fromData).map((k) => ` The ${k.toUpperCase()} pivot in that file was filtered, so its totals were worked out from its Pending Data.`).join('') + '</p></div>';
  }
  if (I.prevWarning) h += `<div class="warn wk-bad"><p>${esc(I.prevWarning)} The Previous Total row will be empty.</p></div>`;
  if (I.vendors.length) {
    h += `<div class="wk-box"><h3>Vendor accounts missing from VENDORS</h3><p class="hint">These accounts are in the Pending or Clearing file but not in Bill Entry → VENDORS. Type the vendor name for each one. Names are remembered for next week; add them to your VENDORS sheet too.</p>` +
      table(['Account', 'Found in', 'Lines', 'Vendor name'], I.vendors.map((v) => `<tr><td>${esc(v.acct)}</td><td>${esc(v.sources.join(', '))}</td><td class="wk-num">${v.count}</td><td><input class="wk-in" data-vend="${esc(v.acct)}" value="${esc(savedV[v.acct] || '')}" placeholder="Vendor name"></td></tr>`).join('')) + '</div>';
  }
  if (I.currencies.length) {
    h += `<div class="wk-box"><h3>Exchange rates to USD</h3><p class="hint">Use today's MMA cross rate, in units of the currency per 1 USD, from <a href="https://mma.gov.mv/#/statistics/exchangerates" target="_blank" rel="noopener">mma.gov.mv → Exchange Rates → Cross Rates</a>. For a currency MMA doesn't list, type the rate you want to use.</p>` +
      table(['Currency', 'Found in', 'Units per USD', ''], I.currencies.map((c) => {
        const s = savedR[c.code];
        return `<tr><td><b>${esc(c.code)}</b></td><td>${esc(c.sources.join(', '))}</td><td><input class="wk-in" data-rate="${esc(c.code)}" inputmode="decimal" value="${s ? esc(s.v) : ''}" placeholder="e.g. 0.9165"></td><td>${MMA.includes(c.code) ? '' : '<span class="wk-pill bad">Not on MMA, enter your own</span>'}${s ? ` <span class="wk-pill">last used ${esc(s.d)}, check it</span>` : ''}</td></tr>`;
      }).join('')) + '</div>';
  }
  if (I.bills.length) {
    h += `<div class="wk-box"><h3>Bill Entry rows missing Vendor or Bill Sent To</h3><p class="hint">Fill in the gaps. When Vendor (F) is empty, the name from Remarks (H) is filled in for you to check. Fix these in your Bill Entry too so they don't come up again.</p>` +
      table(['Row', ...an.data.be.headers], I.bills.map((b) => `<tr><td>${esc(b.bill.excelRow)}</td>` + b.bill.cells.map((v, i) => {
        if (i === 4) return `<td><input class="wk-in${b.needVendor && !b.suggestVendor ? ' need' : ''}" data-bv="${b.bill.excelRow}" value="${esc(b.needVendor ? b.suggestVendor : v)}"></td>`;
        if (i === 16) return `<td><select class="wk-in" data-bs="${b.bill.excelRow}"><option value="">Choose…</option>${WR.STATUSES.map((s) => `<option${WR.matchStatus(v) === s ? ' selected' : ''}>${s}</option>`).join('')}</select>${b.needStatus && v ? `<div class="hint">was "${esc(v)}"</div>` : ''}</td>`;
        return `<td${typeof v === 'number' ? ' class="wk-num"' : ''}>${esc(cell(v))}</td>`;
      }).join('') + '</tr>').join('')) + '</div>';
  }
  if (I.noDate.length) {
    h += `<div class="wk-box"><h3>Pending bills with no Received Date</h3><p class="hint">Enter the received date to include a bill, or leave it blank to leave it out.</p>` +
      table(['Row', 'Vendor', 'Invoice', 'Amount', 'Received date'], I.noDate.map((b) => `<tr><td>${b.excelRow}</td><td>${esc(cell(b.cells[4]))}</td><td>${esc(cell(b.cells[5]))}</td><td class="wk-num">${esc(cell(b.cells[9] ?? b.cells[8]))}</td><td><input class="wk-in" type="date" data-bd="${b.excelRow}"></td></tr>`).join('')) + '</div>';
  }
  const n = I.vendors.length + I.currencies.length + I.bills.length;
  $('#wk-s2hint').textContent = n ? 'Everything below has to be filled in before the report can be built.' : 'Nothing needs your input this week.';
  const box = $('#wk-issues');
  box.innerHTML = h;
  box.oninput = box.onchange = (e) => { if (e.target.value) e.target.classList.remove('need'); };
  $('#wk-s2').hidden = false; $('#wk-s3').hidden = false; $('#wk-result').hidden = true; showMsg('#wk-m3', '');
  $('#wk-s2').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function collect() {
  const a = { vendorNames: {}, rates: {}, bills: {} }, missing = [];
  $$('[data-vend]').forEach((i) => { const v = i.value.trim(); i.classList.toggle('need', !v); if (v) a.vendorNames[i.dataset.vend] = v; else missing.push('vendor name for account ' + i.dataset.vend); });
  $$('[data-rate]').forEach((i) => { const v = parseFloat(i.value); i.classList.toggle('need', !(v > 0)); if (v > 0) a.rates[i.dataset.rate] = v; else missing.push('rate for ' + i.dataset.rate); });
  $$('[data-bv]').forEach((i) => { const v = i.value.trim(); i.classList.toggle('need', !v); (a.bills[i.dataset.bv] ||= {}).vendor = v; if (!v) missing.push('vendor for Bill Entry row ' + i.dataset.bv); });
  $$('[data-bs]').forEach((i) => { const v = i.value; i.classList.toggle('need', !v); (a.bills[i.dataset.bs] ||= {}).status = v; if (!v) missing.push('Bill Sent To for row ' + i.dataset.bs); });
  return { a, missing };
}

/* ---------- step 3 ---------- */
$('#wk-build').addEventListener('click', async () => {
  const { a, missing } = collect();
  if (missing.length) { showMsg('#wk-m3', 'Still missing: ' + missing.slice(0, 6).join('; ') + (missing.length > 6 ? ` and ${missing.length - 6} more` : '') + '.', true); return; }
  const extra = [];
  $$('[data-bd]').forEach((i) => {
    if (!i.value) return;
    const b = an.issues.noDate.find((x) => String(x.excelRow) === i.dataset.bd);
    const [y, m, d] = i.value.split('-').map(Number); b.received = WR.utcDay(y, m - 1, d);
    if (b.received < periodEnd()) extra.push(b);
  });
  const vSave = store.get('vendorNames', {}); Object.assign(vSave, a.vendorNames); store.set('vendorNames', vSave);
  const rSave = store.get('rates', {}); for (const k in a.rates) rSave[k] = { v: a.rates[k], d: dmy(today) }; store.set('rates', rSave);
  const btn = $('#wk-build'); btn.disabled = true; btn.textContent = 'Building…'; showMsg('#wk-m3', '');
  try {
    const kept = an.data.keptBills;
    an.data.keptBills = kept.concat(extra);
    const model = WR.buildModel(an, a, opts());
    an.data.keptBills = kept;
    const rep = await WR.writeXlsx(JSZip, model);
    const wf = await WR.writeWorkingFiles(JSZip, model);
    let name = $('#wk-fname').value.trim() || 'Weekly Pending Report.xlsx'; if (!/\.xlsx$/i.test(name)) name += '.xlsx';
    built = { summary: rep.summary, files: [
      { name, data: rep.data, what: 'Weekly report: USD, MVR, Pending Data, Other Pendings, Working Sheet, Previous' },
      { name: wf.pending.name, data: wf.pending.data, what: 'Pending items: Sheet1, Sheet2, MVR and USD sheets, with pivots' },
      { name: wf.cleared.name, data: wf.cleared.data, what: 'Cleared items: Sheet1, MVR and USD sheets, with pivots' + (model.clearExcluded.length ? ', plus the removed lines' : '') },
    ] };
    renderResult(built.summary);
  } catch (e) { showMsg('#wk-m3', 'Could not build the report: ' + e.message, true); }
  btn.disabled = false; btn.textContent = 'Build report';
});

function chg(nowV, prev) {
  if (prev == null) return '';
  const d = nowV - prev, p = prev ? (d / prev) * 100 : 0;
  return `${d >= 0 ? '▲' : '▼'} ${fmt2(Math.abs(d))} (${d >= 0 ? '+' : '−'}${Math.abs(p).toFixed(2)}%) vs last week`;
}
function renderResult(s) {
  $('#wk-figs').innerHTML =
    `<div class="wk-fig key"><div class="lab">MVR Grand Pending</div><div class="big">MVR ${fmt2(s.mvr.grandTotal)}</div><div class="chg">${chg(s.mvr.grandTotal, s.mvr.prevGrand)}</div></div>` +
    `<div class="wk-fig"><div class="lab">USD Grand Pending</div><div class="big">USD ${fmt2(s.usd.grandTotal)}</div><div class="chg">${chg(s.usd.grandTotal, s.usd.prevGrand)}</div></div>`;
  $('#wk-stats').innerHTML = ['mvr', 'usd'].map((k) => `<span>${k.toUpperCase()}: pending at MoF <b>${fmt2(s[k].mof)}</b> · at IGMH <b>${fmt2(s[k].igmh)}</b> · weekly clear <b>${fmt2(s[k].clear)}</b></span>`).join('');
  $('#wk-dls').innerHTML = built.files.map((f, i) => `<div class="wk-dl"><div><b>${esc(f.name)}</b><span>${esc(f.what)} · ${fmtSize(f.data.length)}</span></div><button type="button" class="ghost" data-dl="${i}">Download</button></div>`).join('');
  $('#wk-result').hidden = false;
  $('#wk-result').scrollIntoView({ behavior: 'smooth', block: 'center' });
}
$('#wk-dls').addEventListener('click', (e) => {
  const b = e.target.closest('[data-dl]'); if (!b || !built) return;
  const f = built.files[+b.dataset.dl];
  saveBlob(new Blob([f.data], { type: XLSX_TYPE }), f.name);
});
$('#wk-zip').addEventListener('click', async () => {
  if (!built) return;
  const zip = new JSZip();
  for (const f of built.files) zip.file(f.name, f.data);
  const end = periodEnd();
  saveBlob(await zip.generateAsync({ type: 'blob' }), 'Weekly report ' + (end ? WR.dotDate(end) : '') + '.zip');
});

renderTray();
