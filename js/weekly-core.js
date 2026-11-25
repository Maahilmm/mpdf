/* Weekly Pending Report engine — reads the four workbooks, applies the weekly routine,
   writes the finished report (.xlsx) with real pivot tables. Works in browser and Node. */
(function (root) {
'use strict';

/* ---------------- constants ---------------- */
const AGE_LABELS = ['0-30','31-60','61-90','91-120','120-360','>360','Pendings MoF','Weekly Clear',
  'GRN Pending','Agreement Pending','PO Pending','Budget Pending','Procurement Pending','Docs Missing',
  'Revenue Pending','HR Pending','IT Pending','Estate Pending','Finance Pending','Delivery Pending',
  'On-Hold','Invoice Issue','Received to Park','Report Pending','Sent for Checking','Parked',
  'Pendings IGMH','Grand Pending'];
const STATUSES = AGE_LABELS.slice(8, 26);           // valid "Bill Sent To" values
const MOF_IDX = 6;                                   // AGE item used for vendor sorting
const AGE_FORMULA_TAIL = ',{0,31,61,91,121,361},{"0-30","31-60","61-90","91-120","120-360",">360"})';

/* ---------------- small helpers ---------------- */
const r2 = x => Math.round((x + Number.EPSILON) * 100) / 100;
const isNum = v => typeof v === 'number' && isFinite(v);
const blank = v => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
const cmp = (a, b) => String(a).localeCompare(String(b), 'en', { sensitivity: 'base' });
function colLetter(n) { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
function colIndex(letters) { let n = 0; for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
const DAY = 86400000;
function serialToDate(s) { return new Date(Math.round((s - 25569) * DAY)); }        // UTC date
function dateToSerial(d) { return Math.round(d.getTime() / DAY) + 25569; }
function utcDay(y, m, d) { return new Date(Date.UTC(y, m, d)); }
function parseTextDate(s) {
  const m = String(s).trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/);
  if (!m) return null;
  let y = +m[3]; if (y < 100) y += 2000;
  const d = +m[1], mo = +m[2];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return utcDay(y, mo - 1, d);
}
function toDate(v) {               // number serial, Date, or text → UTC Date | null
  if (v instanceof Date) return v;
  if (isNum(v)) return v > 20000 && v < 80000 ? serialToDate(v) : null;
  if (typeof v === 'string') return parseTextDate(v);
  return null;
}
function ordinal(n) { const s = ['th','st','nd','rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const MON3 = MONTHS.map(m => m.slice(0, 3));
function longDate(d) { return ordinal(d.getUTCDate()) + ' ' + MONTHS[d.getUTCMonth()] + ' ' + d.getUTCFullYear(); }
function dotDate(d) { const p = n => String(n).padStart(2, '0'); return p(d.getUTCDate()) + '.' + p(d.getUTCMonth() + 1) + '.' + d.getUTCFullYear(); }
function normAcct(v) {
  if (blank(v)) return '';
  let s = isNum(v) ? String(Math.round(v)) : String(v).trim().toUpperCase();
  if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));
  return s;
}
function cleanStr(v) { return blank(v) ? '' : (isNum(v) ? String(v) : String(v).trim()); }
function bucket5(days) { return days <= 30 ? '0-30' : days <= 60 ? '31-60' : days <= 90 ? '61-90' : days <= 120 ? '91-120' : '>120'; }
function bucket6(days) { days = Math.max(0, days); return days <= 30 ? '0-30' : days <= 60 ? '31-60' : days <= 90 ? '61-90' : days <= 120 ? '91-120' : days <= 360 ? '120-360' : '>360'; }

/* ---------------- XML helpers ---------------- */
function unesc(s) {
  return s.replace(/&(lt|gt|quot|apos|amp|#x[0-9a-fA-F]+|#\d+);/g, (m, e) =>
    e === 'lt' ? '<' : e === 'gt' ? '>' : e === 'quot' ? '"' : e === 'apos' ? "'" : e === 'amp' ? '&' :
    e[1] === 'x' ? String.fromCodePoint(parseInt(e.slice(2), 16)) : String.fromCodePoint(parseInt(e.slice(1), 10)));
}
function esc(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}
function attr(tag, name) { const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"')); return m ? unesc(m[1]) : null; }

/* ---------------- XLSX reader ---------------- */
async function openWorkbook(JSZip, data) {
  const zip = await JSZip.loadAsync(data);
  const read = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
  const wbXml = await read('xl/workbook.xml');
  if (!wbXml) throw new Error('Not an Excel (.xlsx) workbook');
  const relsXml = await read('xl/_rels/workbook.xml.rels') || '';
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id'); let t = attr(m[0], 'Target');
    if (t.startsWith('/')) t = t.slice(1); else t = 'xl/' + t.replace(/^\.\//, '');
    rels[id] = t;
  }
  const sheets = [];
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) sheets.push({ name: attr(m[0], 'name'), path: rels[attr(m[0], 'r:id')] });
  // shared strings
  const sst = [];
  const sstXml = await read('xl/sharedStrings.xml');
  if (sstXml) for (const m of sstXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const body = m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let t = ''; for (const tm of body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) t += tm[1];
    sst.push(unesc(t));
  }
  // styles: which xf ids are dates / have a fill
  const stXml = await read('xl/styles.xml') || '';
  const custom = {};
  for (const m of stXml.matchAll(/<numFmt\b[^>]*>/g)) custom[attr(m[0], 'numFmtId')] = attr(m[0], 'formatCode');
  const fills = [];
  const fillsBlock = (stXml.match(/<fills\b[\s\S]*?<\/fills>/) || [''])[0];
  for (const m of fillsBlock.matchAll(/<fill>([\s\S]*?)<\/fill>|<fill\/>/g)) {
    const pt = m[1] ? (m[1].match(/patternType="([^"]*)"/) || [])[1] : null;
    fills.push(!!pt && pt !== 'none' && pt !== 'gray125');
  }
  const xfDate = [], xfFill = [];
  const xfsBlock = (stXml.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [''])[0];
  for (const m of xfsBlock.matchAll(/<xf\b[^>]*>/g)) {
    const id = +attr(m[0], 'numFmtId') || 0, fillId = +attr(m[0], 'fillId') || 0;
    let isD = (id >= 14 && id <= 22) || (id >= 45 && id <= 47);
    if (custom[id]) { const f = custom[id].replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); isD = /[dy]/i.test(f) || (/m/i.test(f) && !/[#0?]/.test(f)); }
    xfDate.push(isD); xfFill.push(!!fills[fillId]);
  }
  return {
    names: sheets.map(s => s.name),
    has: n => sheets.some(s => s.name.toLowerCase() === n.toLowerCase()),
    async sheet(name, opts = {}) {
      const s = typeof name === 'number' ? sheets[name] : sheets.find(x => x.name.toLowerCase() === String(name).toLowerCase());
      if (!s) return null;
      const xml = await read(s.path);
      return parseSheet(xml, sst, xfDate, xfFill, opts.maxCol == null ? 60 : opts.maxCol);
    }
  };
}
function parseSheet(xml, sst, xfDate, xfFill, maxCol) {
  // returns { rows: [ [value...] ], fill: [ [bool...] ] }  (0-based rows/cols)
  const rows = [], fill = [];
  const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = re.exec(xml))) {
    const a = m[1];
    const ref = (a.match(/\br="([A-Z]+)(\d+)"/) || []);
    if (!ref[1]) continue;
    const c = colIndex(ref[1]); if (c > maxCol) continue;
    const r = +ref[2] - 1;
    const sAttr = a.match(/\bs="(\d+)"/); const s = sAttr ? +sAttr[1] : 0;
    const tAttr = a.match(/\bt="(\w+)"/); const t = tAttr ? tAttr[1] : 'n';
    const inner = m[2] || '';
    let v = null;
    if (t === 'inlineStr') { let tx = ''; for (const tm of inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) tx += tm[1]; v = unesc(tx); }
    else {
      const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
      if (vm) {
        const raw = vm[1];
        if (t === 's') v = sst[+raw];
        else if (t === 'str' || t === 'e') v = unesc(raw);
        else if (t === 'b') v = raw === '1';
        else { v = +raw; if (xfDate[s] && isNum(v)) v = serialToDate(v); }
      }
    }
    if (xfFill[s]) { (fill[r] || (fill[r] = []))[c] = true; }
    if (v === null || v === '') continue;
    (rows[r] || (rows[r] = []))[c] = v;
  }
  return { rows, fill };
}

/* ---------------- file detection ---------------- */
async function detect(JSZip, data) {
  const wb = await openWorkbook(JSZip, data);
  if (wb.has('Pending Data')) return { kind: 'report', wb };
  if (wb.has('IGMH') && wb.has('VENDORS')) return { kind: 'billentry', wb };
  const sh = await wb.sheet(0, { maxCol: 40 });
  const hdr = (sh.rows[0] || []).map(cleanStr);
  if (hdr.includes('Document Number') && hdr.includes('Account')) {
    const ci = hdr.indexOf('Clearing Document'), cd = hdr.indexOf('Clearing date');
    let cleared = 0, total = 0;
    for (let r = 1; r < sh.rows.length && total < 400; r++) {
      const row = sh.rows[r]; if (!row || blank(row[hdr.indexOf('Document Number')])) continue;
      total++; if ((ci >= 0 && !blank(row[ci])) || (cd >= 0 && !blank(row[cd]))) cleared++;
    }
    return { kind: total && cleared / total > 0.5 ? 'cleared' : 'open', wb, sheet: sh };
  }
  return { kind: 'unknown', wb };
}

/* ---------------- reading inputs ---------------- */
function readSap(sh) {
  const hdr = (sh.rows[0] || []).map(cleanStr);
  const col = n => { const i = hdr.findIndex(h => h.toLowerCase() === n.toLowerCase()); return i; };
  const need = ['Account', 'Document Number', 'Document Date', 'Posting Date', 'Document currency', 'Amount in doc. curr.', 'Reference'];
  const miss = need.filter(n => col(n) < 0);
  if (miss.length) throw new Error('Missing column(s): ' + miss.join(', '));
  const C = Object.fromEntries(need.map(n => [n, col(n)]));
  const tCol = col('Text'), dtCol = col('Document Type');
  const items = []; let dropped = 0;
  for (let r = 1; r < sh.rows.length; r++) {
    const row = sh.rows[r]; if (!row) continue;
    const coloured = sh.fill[r] && sh.fill[r][0];
    if (coloured || blank(row[C['Document Number']])) { dropped++; continue; }
    const amt = row[C['Amount in doc. curr.']];
    items.push({
      raw: row, rowIdx: r,
      acct: normAcct(row[C.Account]), docNo: cleanStr(row[C['Document Number']]),
      docDate: toDate(row[C['Document Date']]), postDate: toDate(row[C['Posting Date']]),
      cur: cleanStr(row[C['Document currency']]).toUpperCase() || 'MVR',
      amt: isNum(amt) ? Math.abs(amt) : Math.abs(+amt || 0), ref: cleanStr(row[C.Reference]),
      text: tCol >= 0 ? cleanStr(row[tCol]) : '', docType: dtCol >= 0 ? cleanStr(row[dtCol]) : ''
    });
  }
  return { items, dropped, hdr };
}
async function readVendors(wb) {
  const sh = await wb.sheet('VENDORS', { maxCol: 2 });
  const map = {};
  for (let r = 1; r < sh.rows.length; r++) {
    const row = sh.rows[r]; if (!row) continue;
    const a = normAcct(row[0]); const n = cleanStr(row[1]);
    if (a && n && !(a in map)) map[a] = n;
  }
  return map;
}
async function readBillEntry(wb) {
  const sh = await wb.sheet('IGMH', { maxCol: 20 });
  let hr = sh.rows.findIndex(r => r && r.some(v => cleanStr(v).toUpperCase() === 'RECEIVED DATE'));
  if (hr < 0) hr = 1;
  const headers = []; for (let c = 1; c <= 17; c++) headers.push(cleanStr((sh.rows[hr] || [])[c]) || colLetter(c));
  const bills = [];
  for (let r = hr + 1; r < sh.rows.length; r++) {
    const row = sh.rows[r]; if (!row) continue;
    if (cleanStr(row[16]).toUpperCase() !== 'PENDING') continue;
    const real = [1, 2, 5, 6, 7, 9].some(i => !blank(row[i])) || (isNum(row[10]) && row[10] !== 0);
    if (!real) continue;
    const cells = []; for (let c = 1; c <= 17; c++) cells.push(row[c] === undefined ? null : row[c]);
    bills.push({ excelRow: r + 1, entry: row[0], cells });
  }
  return { headers, bills };
}
async function readPrevReport(wb) {
  const out = { pd: null, usd: null, mvr: null, periodEnd: null };
  const pd = await wb.sheet('Pending Data', { maxCol: 14 });
  out.pd = pd.rows;
  for (const key of ['USD', 'MVR']) {
    const sh = await wb.sheet(key, { maxCol: 60 }); if (!sh) continue;
    const info = { labels: null, rows: {}, history: [] };
    const known = new Set(AGE_LABELS.map(x => x.toLowerCase()));
    for (let r = 0; r < sh.rows.length; r++) {
      const row = sh.rows[r]; if (!row) continue;
      const a = cleanStr(row[0]);
      const m = a.match(/TO\s+(\d{1,2})\.(\d{1,2})\.(\d{4})/i);
      if (m && !out.periodEnd) out.periodEnd = utcDay(+m[3], +m[2] - 1, +m[1]);
      const hits = row.filter(v => known.has(cleanStr(v).toLowerCase())).length;
      if (!info.labels && cleanStr(row[0]).toUpperCase() === 'VENDOR') {
        // a filtered pivot shows only some columns: fall back to the standard column order
        info.complete = hits >= 20;
        info.labels = hits >= 20 ? row.map(v => { const t = cleanStr(v); return AGE_LABELS.find(x => x.toLowerCase() === t.toLowerCase()) || t; }) : [''].concat(AGE_LABELS);
        continue;
      }
      if (!info.labels) continue;
      const vals = {}; info.labels.forEach((l, i) => { if (i > 0 && l && isNum(row[i])) vals[l] = row[i]; });
      const al = a.toLowerCase();
      if (al.startsWith('grand total') || al === 'total result') info.rows.grand = vals;
      else if (al.startsWith('previous total')) { info.rows.prev = vals; info.prevLabel = a; }
      else if (al.startsWith('total as at')) info.history.push({ label: a, vals });
    }
    info.rows.pivotGrand = info.rows.grand;
    out[key.toLowerCase()] = info;
  }
  // last week's totals, worked out from last week's Pending Data (not the pivot, which may be filtered or stale)
  const sums = [], hdrs = [];
  (out.pd || []).forEach((row, r) => { if (row && cleanStr(row[1]) === 'Vendor' && cleanStr(row[10]) === 'AGE') hdrs.push(r); });
  hdrs.slice(0, 2).forEach((h, b) => {
    const end = hdrs[b + 1] || out.pd.length, col = b === 0 ? 5 : 6, t = {};
    const asOf = out.periodEnd ? out.periodEnd.getTime() + DAY : null;
    for (let r = h + 1; r < end; r++) { const row = out.pd[r]; if (!row) continue; let a = cleanStr(row[10]); const v = row[col];
      if (!a || !isNum(v) || blank(row[1])) continue;
      const post = toDate(row[4]);
      if (asOf && post && AGE_LABELS.indexOf(a) >= 0 && AGE_LABELS.indexOf(a) < 6) a = bucket6(Math.round((asOf - post.getTime()) / DAY));
      const k = AGE_LABELS.find(x => x.toLowerCase() === a.toLowerCase()) || a; t[k] = r2((t[k] || 0) + v); }
    sums.push(t);
  });
  ['usd', 'mvr'].forEach((k, i) => { const info = out[k]; if (!info) return;
    if (!(info.complete && info.rows.grand && Object.keys(info.rows.grand).length) && sums[i]) { info.rows.grand = sums[i]; info.fromData = true; } });
  return out;
}

/* ---------------- analysis (find what needs the user) ---------------- */
async function analyse(JSZip, files, opt) {
  // files: { open, cleared, billentry, report } each { kind, wb, sheet? }
  const res = { issues: {}, stats: {} };
  const openSh = files.open.sheet || await files.open.wb.sheet(0, { maxCol: 40 });
  const clrSh = files.cleared.sheet || await files.cleared.wb.sheet(0, { maxCol: 40 });
  const open = readSap(openSh), cleared = readSap(clrSh);
  // Weekly Clear: leave out HR payroll postings, salary payments, dividend setoffs (matched on Text)
  const norm = x => String(x).toLowerCase().replace(/\s+/g, ' ').trim();
  const phrases = (opt.clearExclude || []).map(norm).filter(Boolean);
  cleared.excluded = [];
  if (phrases.length) cleared.items = cleared.items.filter(it => {
    const t = norm(it.text);
    const hit = phrases.find(p => t.includes(p));
    if (hit) { cleared.excluded.push({ ...it, phrase: hit }); return false; }
    return true;
  });
  const vendors = await readVendors(files.billentry.wb);
  const be = await readBillEntry(files.billentry.wb);
  const prev = await readPrevReport(files.report.wb);
  res.data = { open, cleared, vendors, be, prev, openSh, clrSh };
  // vendor accounts missing
  const missing = {};
  const chk = (list, src) => { for (const it of list) if (it.acct && !vendors[it.acct]) { const m = missing[it.acct] || (missing[it.acct] = { count: 0, src: new Set() }); m.count++; m.src.add(src); } };
  chk(open.items, 'Pending'); chk(cleared.items, 'Clearing'); chk(cleared.excluded, 'Clearing (removed lines)');
  res.issues.vendors = Object.keys(missing).sort().map(a => ({ acct: a, count: missing[a].count, sources: [...missing[a].src] }));
  // foreign currencies
  const cur = {};
  const addCur = (c, src) => { (cur[c] || (cur[c] = new Set())).add(src); };
  for (const it of open.items) if (!['MVR', 'USD'].includes(it.cur)) addCur(it.cur, 'Pending items');
  for (const it of cleared.items) if (!['MVR', 'USD'].includes(it.cur)) addCur(it.cur, 'Cleared items');
  // bill entry: cutoff + review
  const end = opt.periodEnd;
  const kept = [], review = [], noDate = [];
  for (const b of be.bills) {
    const rec = toDate(b.cells[1]);
    if (!rec) { noDate.push(b); continue; }
    if (rec.getTime() >= end.getTime()) continue;
    b.received = rec; kept.push(b);
  }
  for (const b of kept) {
    const c = cleanStr(b.cells[7]).toUpperCase();
    if (c && c !== 'MVR' && c !== 'USD') addCur(c, 'Bill Entry');
    const vendor = cleanStr(b.cells[4]), sent = cleanStr(b.cells[16]);
    const okStatus = sent && matchStatus(sent);
    if (!vendor || !okStatus) review.push({ bill: b, needVendor: !vendor, suggestVendor: vendor || cleanStr(b.cells[6]),
      needStatus: !okStatus, given: sent });
  }
  res.issues.currencies = Object.keys(cur).sort().map(c => ({ code: c, sources: [...cur[c]] }));
  res.issues.bills = review; res.issues.noDate = noDate;
  res.data.keptBills = kept;
  res.stats = { openItems: open.items.length, openDropped: open.dropped, clearedItems: cleared.items.length,
    clearedDropped: cleared.dropped, clearedExcluded: cleared.excluded.length, pendingBills: be.bills.length, keptBills: kept.length,
    excludedBills: be.bills.length - kept.length - noDate.length };
  if (!prev.usd || !prev.usd.rows.grand || !prev.mvr || !prev.mvr.rows.grand) res.issues.prevWarning = 'Could not find the Grand Total rows in last week\'s USD/MVR sheets.';
  return res;
}
function matchStatus(s) {
  const k = String(s).trim().toLowerCase().replace(/[\s,.]+$/g, '').replace(/\s+/g, ' ');
  return STATUSES.find(x => x.toLowerCase() === k) || null;
}

/* ---------------- build the report model ---------------- */
function buildModel(an, answers, opt) {
  // answers: { vendorNames:{acct:name}, rates:{code:unitsPerUsd}, bills:{excelRow:{vendor,status,received}} }
  const { open, cleared, vendors, be, prev } = an.data;
  const asOf = new Date(opt.periodEnd.getTime() + DAY);   // ageing date: day after the period
  const vname = a => answers.vendorNames[a] || vendors[a] || a;
  const ahmed = new Set(opt.ahmedAccounts.map(normAcct));
  const rate = c => { const r = +answers.rates[c]; if (!(r > 0)) throw new Error('No exchange rate for ' + c); return r; };

  const classify = (it, isPending) => {
    const v = vname(it.acct);
    if (it.cur === 'MVR') return { blk: 'mvr', amt: r2(it.amt), v };
    if (it.cur === 'USD') {
      if (isPending && ahmed.has(it.acct)) return { blk: 'mvr', amt: r2(it.amt * opt.ahmedRate), v };
      return { blk: 'usd', amt: r2(it.amt), v };
    }
    return { blk: 'usd', amt: r2(it.amt / rate(it.cur)), v };
  };
  const blocks = { usd: { lines: [], igmh: [], clear: {} }, mvr: { lines: [], igmh: [], clear: {} } };
  const how = (it, isPending) => it.cur === 'MVR' ? { k: 'abs' } : it.cur === 'USD' ? (isPending && ahmed.has(it.acct) ? { k: 'mul', r: opt.ahmedRate } : { k: 'abs' }) : { k: 'div', r: rate(it.cur) };
  const openRows = [], clearRows = [];
  for (const it of open.items) { const c = classify(it, true); blocks[c.blk].lines.push({ ...it, vendor: c.v, amount: c.amt }); openRows.push({ it, vendor: c.v, blk: c.blk, amount: c.amt, how: how(it, true) }); }
  for (const it of cleared.items) { const c = classify(it, false); const b = blocks[c.blk].clear; b[c.v] = r2((b[c.v] || 0) + c.amt); clearRows.push({ it, vendor: c.v, blk: c.blk, amount: c.amt, how: how(it, false) }); }
  const clearExcluded = (cleared.excluded || []).map(it => ({ it, vendor: vname(it.acct) }));

  // bill entry rows (Other Pendings)
  const other = [];
  for (const b of an.data.keptBills) {
    const fix = answers.bills[b.excelRow] || {};
    const cells = b.cells.slice();
    if (fix.vendor) cells[4] = fix.vendor;
    const status = fix.status || matchStatus(cells[16]) || cleanStr(cells[16]);
    cells[16] = status;
    const days = Math.round((asOf.getTime() - b.received.getTime()) / DAY);
    cells[2] = days; cells[3] = bucket5(days);
    let cur = cleanStr(cells[7]).toUpperCase(), blk, amt;
    if (cur && cur !== 'MVR' && cur !== 'USD') {
      const usd = r2((+cells[8] || 0) / rate(cur));
      cells[7] = 'USD'; cells[8] = usd; blk = 'usd'; amt = usd;
    } else if (isNum(cells[8]) && cells[8] !== 0) { blk = 'usd'; amt = r2(cells[8]); }
    else { blk = 'mvr'; amt = isNum(cells[9]) ? r2(cells[9]) : 0; }
    other.push(cells);
    blocks[blk].igmh.push({ vendor: cleanStr(cells[4]), billDate: cells[0], amount: amt, invoice: cleanStr(cells[5]), status });
  }
  for (const k of ['usd', 'mvr']) {
    const B = blocks[k];
    B.mof = sumBy(B.lines, x => x.vendor, x => x.amount, true);
    B.igmhSum = sumBy(B.igmh, x => x.vendor, x => x.amount, false);        // first-appearance order
    const names = B.igmhSum.map(x => x.vendor);
    for (const x of B.mof) if (!names.includes(x.vendor)) names.push(x.vendor);
    B.grandNames = names;
    B.clearList = Object.keys(B.clear).sort(cmp).map(v => ({ vendor: v, amount: B.clear[v] }));
  }
  return { blocks, other, beHeaders: be.headers, prev, opt, openRows, clearRows, clearExcluded, openSh: an.data.openSh, clrSh: an.data.clrSh };
}
function sumBy(list, kf, vf, sortAlpha) {
  const order = [], m = {};
  for (const x of list) { const k = kf(x); if (!(k in m)) { m[k] = 0; order.push(k); } m[k] += vf(x); }
  if (sortAlpha) order.sort(cmp);
  return order.map(k => ({ vendor: k, amount: r2(m[k]) }));
}

/* ---------------- workbook writer ---------------- */
function Styles() {
  const fonts = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'];
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const numFmts = []; const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const keyMap = { '': 0 }, fontMap = {}, fillMap = {}, fmtMap = {};
  const builtin = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, 'm/d/yyyy': 14 };
  function font(f) {
    const k = JSON.stringify(f); if (k in fontMap) return fontMap[k];
    fonts.push('<font>' + (f.b ? '<b/>' : '') + '<sz val="' + (f.sz || 11) + '"/>' + (f.color ? '<color rgb="' + f.color + '"/>' : '<color theme="1"/>') + '<name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');
    return (fontMap[k] = fonts.length - 1);
  }
  function fill(rgb) {
    if (fillMap[rgb] != null) return fillMap[rgb];
    fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + rgb + '"/><bgColor indexed="64"/></patternFill></fill>');
    return (fillMap[rgb] = fills.length - 1);
  }
  function fmt(code) {
    if (code in builtin) return builtin[code]; if (code in fmtMap) return fmtMap[code];
    const id = 164 + numFmts.length; numFmts.push('<numFmt numFmtId="' + id + '" formatCode="' + esc(code) + '"/>'); return (fmtMap[code] = id);
  }
  return {
    id(st) {
      if (!st) return 0;
      const k = JSON.stringify(st); if (k in keyMap) return keyMap[k];
      const f = st.b || st.sz || st.color ? font({ b: st.b, sz: st.sz, color: st.color }) : 0;
      const fi = st.fill ? fill(st.fill) : 0, nf = st.fmt ? fmt(st.fmt) : 0;
      xfs.push('<xf numFmtId="' + nf + '" fontId="' + f + '" fillId="' + fi + '" borderId="0" xfId="0"' +
        (nf ? ' applyNumberFormat="1"' : '') + (f ? ' applyFont="1"' : '') + (fi ? ' applyFill="1"' : '') +
        (st.wrap || st.center ? ' applyAlignment="1"><alignment' + (st.center ? ' horizontal="center"' : '') + (st.wrap ? ' wrapText="1"' : '') + '/></xf>' : '/>'));
      return (keyMap[k] = xfs.length - 1);
    },
    xml() {
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (numFmts.length ? '<numFmts count="' + numFmts.length + '">' + numFmts.join('') + '</numFmts>' : '') +
        '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts><fills count="' + fills.length + '">' + fills.join('') + '</fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>' +
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
    }
  };
}
function Sheet(name) {
  const rows = []; // rows[r] = { c: { col: cell } }
  return {
    name, rows, cols: {}, freeze: null, autoFilter: null, pivot: null,
    set(r, c, v, st, f) { // r,c 0-based. v: number|string|Date|null ; f: formula (without =)
      if (v === undefined) v = null;
      if (v === null && !f && !st) return;
      (rows[r] || (rows[r] = {}))[c] = { v, st, f };
    },
    width(c, w) { this.cols[c] = w; }
  };
}
function writeSheetXml(sh, S, sst) {
  const sstIdx = s => { if (!(s in sst.map)) { sst.map[s] = sst.list.length; sst.list.push(s); } return sst.map[s]; };
  let maxR = 0, maxC = 0, out = [];
  sh.rows.forEach((row, r) => {
    if (!row) return;
    const cells = Object.keys(row).map(Number).sort((a, b) => a - b);
    if (!cells.length) return;
    maxR = Math.max(maxR, r); maxC = Math.max(maxC, cells[cells.length - 1]);
    let x = '<row r="' + (r + 1) + '">';
    for (const c of cells) {
      const { v, st, f } = row[c]; const ref = colLetter(c) + (r + 1); const s = S.id(st);
      const sa = s ? ' s="' + s + '"' : '';
      const fx = f ? '<f>' + esc(f) + '</f>' : '';
      if (v === null || v === undefined || v === '') x += f ? '<c r="' + ref + '"' + sa + ' t="str">' + fx + '<v></v></c>' : '<c r="' + ref + '"' + sa + '/>';
      else if (v instanceof Date) x += '<c r="' + ref + '"' + sa + '>' + fx + '<v>' + dateToSerial(v) + '</v></c>';
      else if (isNum(v)) x += '<c r="' + ref + '"' + sa + '>' + fx + '<v>' + v + '</v></c>';
      else if (typeof v === 'boolean') x += '<c r="' + ref + '"' + sa + ' t="b">' + fx + '<v>' + (v ? 1 : 0) + '</v></c>';
      else if (f) x += '<c r="' + ref + '"' + sa + ' t="str">' + fx + '<v>' + esc(v) + '</v></c>';
      else x += '<c r="' + ref + '"' + sa + ' t="s"><v>' + sstIdx(String(v)) + '</v></c>';
    }
    out.push(x + '</row>');
  });
  const cols = Object.keys(sh.cols).map(Number).sort((a, b) => a - b);
  const pane = sh.freeze ? '<pane ' + (sh.freeze.c ? 'xSplit="' + sh.freeze.c + '" ' : '') + (sh.freeze.r ? 'ySplit="' + sh.freeze.r + '" ' : '') +
    'topLeftCell="' + colLetter(sh.freeze.c || 0) + ((sh.freeze.r || 0) + 1) + '" activePane="' + (sh.freeze.c && sh.freeze.r ? 'bottomRight' : sh.freeze.r ? 'bottomLeft' : 'topRight') + '" state="frozen"/>' : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<dimension ref="A1:' + colLetter(maxC) + (maxR + 1) + '"/>' +
    '<sheetViews><sheetView workbookViewId="0"' + (sh.tabSelected ? ' tabSelected="1"' : '') + '>' + pane + '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
    (cols.length ? '<cols>' + cols.map(c => '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + sh.cols[c] + '" customWidth="1"/>').join('') + '</cols>' : '') +
    '<sheetData>' + out.join('') + '</sheetData>' +
    (sh.autoFilter ? '<autoFilter ref="' + sh.autoFilter + '"/>' : '') +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
}

/* ---------------- pivot cache + table ---------------- */
function isoDate(d) { return d.toISOString().slice(0, 19); }
function fieldInfo(vals) {
  let s = 0, n = 0, d = 0, b = 0, int = true, mn = Infinity, mx = -Infinity, dmn = null, dmx = null;
  for (const v of vals) {
    if (v === null || v === undefined || v === '') b++;
    else if (v instanceof Date) { d++; if (!dmn || v < dmn) dmn = v; if (!dmx || v > dmx) dmx = v; }
    else if (isNum(v)) { n++; if (!Number.isInteger(v)) int = false; mn = Math.min(mn, v); mx = Math.max(mx, v); }
    else s++;
  }
  const a = [];
  if (!s && !b) a.push('containsSemiMixedTypes="0"');
  if (!s && !n) a.push('containsNonDate="0"');
  if (d) a.push('containsDate="1"');
  if (!s) a.push('containsString="0"');
  if (b) a.push('containsBlank="1"');
  if ((s ? 1 : 0) + (n ? 1 : 0) + (d ? 1 : 0) > 1) a.push('containsMixedTypes="1"');
  if (n) { a.push('containsNumber="1"'); if (int) a.push('containsInteger="1"'); a.push('minValue="' + mn + '"', 'maxValue="' + mx + '"'); }
  if (d) a.push('minDate="' + isoDate(dmn) + '"', 'maxDate="' + isoDate(dmx) + '"');
  return '<sharedItems' + (a.length ? ' ' + a.join(' ') : '') + '/>';
}
function inlineVal(v) {
  if (v === null || v === undefined || v === '') return '<m/>';
  if (v instanceof Date) return '<d v="' + isoDate(v) + '"/>';
  if (isNum(v)) return '<n v="' + v + '"/>';
  return '<s v="' + esc(v) + '"/>';
}
/* records: array of arrays (10 values, cols B..K). Field 0 = Vendor (shared), 9 = AGE (shared) */
function buildPivot(o) {
  // o: { names[10], fmt[10], records, ageItems, vendorItems, dataFld, dataName, sheet, ref, cacheId, rowOrder, colCount, location }
  const recXml = [], fieldVals = o.names.map(() => []);
  const vIdx = new Map(o.vendorItems.map((v, i) => [v, i])), aIdx = new Map(o.ageItems.map((v, i) => [v, i]));
  for (const rec of o.records) {
    let x = '<r>';
    rec.forEach((v, i) => {
      if (i === 0) x += '<x v="' + vIdx.get(v === null || v === '' ? null : v) + '"/>';
      else if (i === 9) x += '<x v="' + aIdx.get(v) + '"/>';
      else { x += inlineVal(v); fieldVals[i].push(v); }
    });
    recXml.push(x + '</r>');
  }
  const cf = o.names.map((n, i) => {
    let si;
    if (i === 0) si = '<sharedItems containsBlank="1" count="' + o.vendorItems.length + '">' + o.vendorItems.map(v => v === null ? '<m/>' : '<s v="' + esc(v) + '"/>').join('') + '</sharedItems>';
    else if (i === 9) si = '<sharedItems count="' + o.ageItems.length + '">' + o.ageItems.map(v => '<s v="' + esc(v) + '"/>').join('') + '</sharedItems>';
    else si = fieldInfo(fieldVals[i]);
    return '<cacheField name="' + esc(n) + '" numFmtId="' + o.fmt[i] + '">' + si + '</cacheField>';
  }).join('');
  const def = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Weekly Report Tool" refreshedDate="' + (dateToSerial(o.today)) + '" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="' + o.records.length + '">' +
    '<cacheSource type="worksheet"><worksheetSource ref="' + o.ref + '" sheet="' + esc(o.sheet) + '"/></cacheSource><cacheFields count="10">' + cf + '</cacheFields></pivotCacheDefinition>';
  const recs = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="' + o.records.length + '">' + recXml.join('') + '</pivotCacheRecords>';
  const vItems = o.vendorItems.map((v, i) => '<item x="' + i + '"/>').join('') + '<item t="default"/>';
  const aItems = o.ageItems.map((v, i) => '<item x="' + i + '"/>').join('') + '<item t="default"/>';
  const pf = o.names.map((n, i) => {
    if (i === 0) return '<pivotField axis="axisRow" showAll="0" sortType="descending"><items count="' + (o.vendorItems.length + 1) + '">' + vItems + '</items>' +
      '<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="2"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference><reference field="9" count="1" selected="0"><x v="' + MOF_IDX + '"/></reference></references></pivotArea></autoSortScope></pivotField>';
    if (i === 9) return '<pivotField axis="axisCol" showAll="0"><items count="' + (o.ageItems.length + 1) + '">' + aItems + '</items></pivotField>';
    if (i === o.dataFld) return '<pivotField dataField="1" numFmtId="4" showAll="0"/>';
    return '<pivotField showAll="0"/>';
  }).join('');
  const xv = v => v ? '<x v="' + v + '"/>' : '<x/>';
  const rowItems = o.rowOrder.map(i => '<i>' + xv(i) + '</i>').join('') + '<i t="grand"><x/></i>';
  const colItems = o.ageItems.map((v, i) => '<i>' + xv(i) + '</i>').join('');
  const tbl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="' + o.tableName + '" cacheId="' + o.cacheId + '" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" colGrandTotals="0" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" rowHeaderCaption="VENDOR" colHeaderCaption="AGE">' +
    '<location ref="' + o.location + '" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>' +
    '<pivotFields count="10">' + pf + '</pivotFields>' +
    '<rowFields count="1"><field x="0"/></rowFields><rowItems count="' + (o.rowOrder.length + 1) + '">' + rowItems + '</rowItems>' +
    '<colFields count="1"><field x="9"/></colFields><colItems count="' + o.ageItems.length + '">' + colItems + '</colItems>' +
    '<dataFields count="1"><dataField name="' + esc(o.dataName) + '" fld="' + o.dataFld + '" baseField="0" baseItem="0" numFmtId="4"/></dataFields>' +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';
  return { def, recs, tbl };
}

/* ---------------- assemble the report ---------------- */
const FILL = { header: 'FFBFBFBF', dummy: 'FF000000', lines: 'FFFBE5D6', mofUsd: null, mofMvr: 'FFFBE5D6',
  igmhUsd: 'FFEDE2F6', igmhMvr: 'FFDEEBF7', igmhSumUsd: 'FFD9D9D9', igmhSumMvr: 'FFE2C5FF', grand: 'FFFFF2CC', clear: 'FFC5E0B4', hist: 'FFDEEBF7' };
const NUM = '#,##0.00', DATE = 'm/d/yyyy';

function buildWorkbookParts(model) {
  const { blocks, other, beHeaders, prev, opt } = model;
  const today = opt.today, end = opt.periodEnd, asOf = new Date(opt.periodEnd.getTime() + DAY),
    asOfF = 'DATE(' + asOf.getUTCFullYear() + ',' + (asOf.getUTCMonth() + 1) + ',' + asOf.getUTCDate() + ')', start = new Date(end.getTime() - 6 * DAY);
  const S = Styles(), sst = { list: [], map: {} };
  const pd = Sheet('Pending Data');
  pd.width(0, 3); pd.width(1, 36.7); pd.width(2, 14); pd.width(3, 12); pd.width(4, 12); pd.width(5, 15); pd.width(6, 15); pd.width(7, 28.4); pd.width(8, 8); pd.width(9, 10); pd.width(10, 18);
  ['Pending at MoF', 'Pending at IGMH', 'Grand Pending', 'Weekly Clear'].forEach((t, i) => pd.set(0, 2 + i, t, { b: true }));
  pd.set(1, 1, 'USD', { b: true }); pd.set(2, 1, 'MVR', { b: true });
  let r = 3; const pivots = {}, summary = {};
  const ageStyleIdx = {};
  for (const key of ['usd', 'mvr']) {
    const B = blocks[key], amtCol = key === 'usd' ? 5 : 6, L = colLetter(amtCol);
    const hdrRow = r;
    const hdr = ['Vendor', 'Document Number', 'Document Date', 'Posting Date', 'Amount in USD', 'Amount In MVR', 'Invoice Number', key === 'usd' ? 'Clearing ' : '.', 'No. of Days Pending', 'AGE'];
    hdr.forEach((h, i) => pd.set(r, 1 + i, h, { b: true, fill: FILL.header }));
    r++;
    const dummyStart = r;
    AGE_LABELS.forEach((a, i) => { for (let c = 1; c <= 10; c++) pd.set(r, c, null, { fill: FILL.dummy }); pd.set(r, 8, i + 1, { fill: FILL.dummy }); pd.set(r, 10, a, { fill: FILL.dummy }); pd.set(r, 11, i + 1); r++; });
    const lookupRange = '$K$' + (dummyStart + 1) + ':$L$' + (dummyStart + 28);
    if (key === 'usd') Object.assign(ageStyleIdx, { range: lookupRange });
    const recs = [];
    const ord = a => AGE_LABELS.indexOf(a) + 1;
    const pushDummyRecs = () => AGE_LABELS.forEach((a, i) => recs.push([null, null, null, null, null, null, null, i + 1, null, a]));
    pushDummyRecs();
    const row = (vals, fillC, ageLabel, amtVal, amtF, extra) => {
      const st = fillC ? { fill: fillC } : null;
      for (let c = 1; c <= 10; c++) pd.set(r, c, null, st);
      pd.set(r, 1, vals.vendor, st);
      if (vals.docNo != null) pd.set(r, 2, vals.docNo, st);
      if (vals.docDate != null) pd.set(r, 3, vals.docDate, vals.docDate instanceof Date ? { fill: fillC || undefined, fmt: DATE } : st);
      if (vals.postDate != null) pd.set(r, 4, vals.postDate, { fill: fillC || undefined, fmt: DATE });
      pd.set(r, amtCol, amtVal, { fill: fillC || undefined, fmt: NUM }, amtF);
      if (vals.invoice != null && vals.invoice !== '') pd.set(r, 7, vals.invoice, st);
      pd.set(r, 8, ord(ageLabel), st, 'VLOOKUP(K' + (r + 1) + ',' + lookupRange + ',2,FALSE)');
      let days = null;
      if (extra && extra.days != null) { days = extra.days; pd.set(r, 9, days, st, asOfF + '-E' + (r + 1)); pd.set(r, 10, ageLabel, st, 'LOOKUP(MAX(0,J' + (r + 1) + ')' + AGE_FORMULA_TAIL); }
      else pd.set(r, 10, ageLabel, st);
      const rec = [vals.vendor, vals.docNo || null, vals.docDate == null ? null : vals.docDate, vals.postDate || null, null, null, vals.invoice || null, ord(ageLabel), days, ageLabel];
      rec[amtCol - 1] = amtVal; recs.push(rec); r++;
    };
    // 1 lines
    const linesStart = r;
    for (const it of B.lines) {
      const days = it.postDate ? Math.round((asOf.getTime() - it.postDate.getTime()) / DAY) : 0;
      row({ vendor: it.vendor, docNo: it.docNo, docDate: it.docDate, postDate: it.postDate, invoice: it.ref }, key === 'usd' ? FILL.lines : null, bucket6(days), it.amount, null, { days });
      if (key === 'mvr') pd.set(r - 1, 10, bucket6(days), { fill: FILL.lines }, 'LOOKUP(MAX(0,J' + r + ')' + AGE_FORMULA_TAIL);
    }
    const linesEnd = r - 1;
    // 2 MoF pivot
    const mofStart = r;
    for (const x of B.mof) row({ vendor: x.vendor }, key === 'usd' ? FILL.mofUsd : FILL.mofMvr, 'Pendings MoF', x.amount);
    const mofEnd = r - 1;
    // 3 IGMH lines
    const igStart = r;
    for (const x of B.igmh) row({ vendor: x.vendor, docDate: toDate(x.billDate) || (blank(x.billDate) ? null : cleanStr(x.billDate)), invoice: x.invoice }, key === 'usd' ? FILL.igmhUsd : FILL.igmhMvr, x.status, x.amount);
    const igEnd = r - 1;
    // 4 IGMH pivot
    const igsStart = r;
    const igRange = c => igEnd >= igStart ? c + '$' + (igStart + 1) + ':' + c + '$' + (igEnd + 1) : null;
    for (const x of B.igmhSum) row({ vendor: x.vendor }, key === 'usd' ? FILL.igmhSumUsd : FILL.igmhSumMvr, 'Pendings IGMH', x.amount,
      'SUMIF(' + igRange('B') + ',B' + (r + 1) + ',' + igRange(L) + ')');
    const igsEnd = r - 1;
    // 5 Grand pending
    const grStart = r;
    const igsMap = Object.fromEntries(B.igmhSum.map(x => [x.vendor, x.amount])), mofMap = Object.fromEntries(B.mof.map(x => [x.vendor, x.amount]));
    for (const v of B.grandNames) {
      const parts = [];
      if (igsEnd >= igsStart) parts.push('SUMIF($B$' + (igsStart + 1) + ':$B$' + (igsEnd + 1) + ',B' + (r + 1) + ',$' + L + '$' + (igsStart + 1) + ':$' + L + '$' + (igsEnd + 1) + ')');
      if (mofEnd >= mofStart) parts.push('SUMIF($B$' + (mofStart + 1) + ':$B$' + (mofEnd + 1) + ',B' + (r + 1) + ',$' + L + '$' + (mofStart + 1) + ':$' + L + '$' + (mofEnd + 1) + ')');
      row({ vendor: v }, FILL.grand, 'Grand Pending', r2((igsMap[v] || 0) + (mofMap[v] || 0)), parts.join('+') || null);
    }
    const grEnd = r - 1;
    // 6 Weekly clear
    const clStart = r;
    for (const x of B.clearList) row({ vendor: x.vendor }, FILL.clear, 'Weekly Clear', x.amount);
    const clEnd = r - 1;
    const lastRow = r - 1;
    // top totals
    const sumF = (a, b) => b >= a ? 'SUM(' + L + (a + 1) + ':' + L + (b + 1) + ')' : null;
    const tot = list => r2(list.reduce((s, x) => s + x.amount, 0));
    const tr = key === 'usd' ? 1 : 2;
    const tv = [tot(B.mof), tot(B.igmhSum), r2(tot(B.mof) + tot(B.igmhSum)), tot(B.clearList)];
    [[mofStart, mofEnd], [igsStart, igsEnd], [grStart, grEnd], [clStart, clEnd]].forEach(([a, b], i) => pd.set(tr, 2 + i, tv[i], { fmt: NUM }, sumF(a, b) || '0'));
    summary[key] = { mof: tv[0], igmh: tv[1], grand: tv[2], clear: tv[3] };
    // pivot description
    const vendorSet = new Set(); recs.forEach(x => { if (x[0]) vendorSet.add(x[0]); });
    const extraAges = []; recs.forEach(x => { if (!AGE_LABELS.includes(x[9]) && !extraAges.includes(x[9])) extraAges.push(x[9]); });
    pivots[key] = { records: recs, hdr, hdrRow, lastRow, vendors: [...vendorSet].sort(cmp), ages: AGE_LABELS.concat(extraAges), amtCol };
    r += 1;
  }
  pd.freeze = { r: 4 };
  pd.autoFilter = 'B4:K' + (pivots.mvr.lastRow + 1);

  /* ---------- pivot sheets ---------- */
  const pivotSheets = {}, parts = [];
  let cacheId = 1;
  for (const key of ['usd', 'mvr']) {
    const P = pivots[key];
    const sh = Sheet(key.toUpperCase());
    const ages = P.ages, nAge = ages.length;
    const vendorItems = [null].concat(P.vendors);
    const vi = new Map(vendorItems.map((v, i) => [v, i]));
    // aggregate
    const agg = {}; // vendor -> age -> sum
    const has = {};
    for (const rec of P.records) {
      const v = rec[0] || null, a = rec[9], amt = rec[P.amtCol - 1];
      if (!agg[v]) { agg[v] = {}; has[v] = {}; }
      if (isNum(amt)) { agg[v][a] = (agg[v][a] || 0) + amt; has[v][a] = true; }
    }
    const mofOf = v => (agg[v] && has[v]['Pendings MoF']) ? agg[v]['Pendings MoF'] : null;
    const rowOrderV = vendorItems.slice().sort((a, b) => {
      const x = mofOf(a), y = mofOf(b);
      if (x === null && y === null) return vi.get(a) - vi.get(b);
      if (x === null) return 1; if (y === null) return -1;
      return y - x || vi.get(a) - vi.get(b);
    });
    const title = key === 'usd' ? 'USD PENDINGS DURING THE PERIOD ' : 'MVR PENDING DURING THE PERIOD ';
    sh.set(1, 0, title + dotDate(start) + ' TO ' + dotDate(end), { b: true, sz: 20 });
    const top = 4; // row 5
    sh.set(top, 0, key === 'usd' ? 'USD Pendings' : 'MVR PENDINGS', { b: true });
    sh.set(top, 1, 'AGE', { b: true });
    sh.set(top + 1, 0, 'VENDOR', { b: true });
    ages.forEach((a, i) => sh.set(top + 1, 1 + i, a, { b: true, wrap: true, center: true }));
    let rr = top + 2;
    const colTot = {};
    for (const v of rowOrderV) {
      sh.set(rr, 0, v === null ? '(blank)' : v);
      ages.forEach((a, i) => { if (agg[v] && has[v][a]) { const val = r2(agg[v][a]); sh.set(rr, 1 + i, val, { fmt: NUM }); colTot[a] = (colTot[a] || 0) + agg[v][a]; } });
      rr++;
    }
    const grandRow = rr;
    sh.set(rr, 0, 'Grand Total', { b: true });
    ages.forEach((a, i) => { if (a in colTot) sh.set(rr, 1 + i, r2(colTot[a]), { b: true, fmt: NUM }); });
    rr++;
    // history rows
    const info = prev[key] || { rows: {}, history: [] };
    const prevEndText = prev.periodEnd ? longDate(prev.periodEnd) : 'last week';
    const oldPrevText = info.prevLabel ? (info.prevLabel.match(/\(\s*(.*?)\s*\)/) || [, info.prevLabel.replace(/^Previous Total\s*/i, '')])[1] : '';
    const hist = [];
    if (info.history.length) hist.push(info.history[info.history.length - 1]);
    if (info.rows.prev) hist.push({ label: 'Total as at ' + oldPrevText, vals: info.rows.prev });
    while (hist.length > 2) hist.shift();
    const HS = { fill: FILL.hist, b: true };
    const prevRow = rr;
    sh.set(rr, 0, 'Previous Total ( ' + prevEndText + ')', HS);
    ages.forEach((a, i) => { const v = info.rows.grand && info.rows.grand[a]; sh.set(rr, 1 + i, isNum(v) ? v : null, { ...HS, fmt: NUM }); });
    rr++;
    const chgRow = rr;
    sh.set(rr, 0, 'Change', HS);
    ages.forEach((a, i) => { const c = colLetter(1 + i); const g = colTot[a] || 0, p = (info.rows.grand && info.rows.grand[a]) || 0; sh.set(rr, 1 + i, r2(g - p), { ...HS, fmt: NUM }, c + (grandRow + 1) + '-' + c + (prevRow + 1)); });
    rr++;
    sh.set(rr, 0, '% change', HS);
    ages.forEach((a, i) => { const c = colLetter(1 + i); const g = colTot[a] || 0, p = (info.rows.grand && info.rows.grand[a]) || 0;
      sh.set(rr, 1 + i, p ? (g - p) / p : '', { ...HS, fmt: '0.00%' }, 'IFERROR(' + c + (chgRow + 1) + '/' + c + (prevRow + 1) + ',"")'); });
    rr++;
    for (const h of hist) {
      sh.set(rr, 0, h.label, HS);
      ages.forEach((a, i) => { const v = h.vals[a]; sh.set(rr, 1 + i, isNum(v) ? v : null, { ...HS, fmt: NUM }); });
      rr++;
    }
    sh.width(0, 43.7); for (let i = 0; i < nAge; i++) sh.width(1 + i, 15.5);
    sh.freeze = { r: top + 2, c: 1 };
    const lastCol = colLetter(nAge);
    const built = buildPivot({
      names: P.hdr, fmt: [0, 0, 14, 14, 4, 4, 0, 0, 0, 0], records: P.records, ageItems: ages, vendorItems,
      dataFld: P.amtCol - 1, dataName: key === 'usd' ? 'USD Pendings' : 'MVR PENDINGS', sheet: 'Pending Data',
      ref: 'B' + (P.hdrRow + 1) + ':K' + (P.lastRow + 1), cacheId: cacheId, tableName: key === 'usd' ? 'USDPivot' : 'MVRPivot',
      rowOrder: rowOrderV.map(v => vi.get(v)), location: 'A' + (top + 1) + ':' + lastCol + (grandRow + 1), today
    });
    sh.pivot = { n: cacheId, ...built };
    cacheId++;
    pivotSheets[key] = sh;
    summary[key].grandTotal = r2(colTot['Grand Pending'] || 0);
    summary[key].prevGrand = info.rows.grand ? info.rows.grand['Grand Pending'] || 0 : null;
    summary[key].vendors = P.vendors.length;
  }
  pivotSheets.usd.tabSelected = true;

  /* ---------- Other Pendings ---------- */
  const op = Sheet('Other Pendings');
  beHeaders.forEach((h, i) => op.set(0, i, h, { b: true, fill: FILL.header }));
  other.forEach((cells, k) => cells.forEach((v, i) => {
    if (v === null || v === undefined || v === '') return;
    const isDateCol = i === 0 || i === 1 || i === 13;
    if (v instanceof Date) op.set(k + 1, i, v, { fmt: DATE });
    else if (isNum(v) && isDateCol && v > 20000) op.set(k + 1, i, serialToDate(v), { fmt: DATE });
    else if (isNum(v) && (i === 8 || i === 9)) op.set(k + 1, i, v, { fmt: NUM });
    else op.set(k + 1, i, v);
  }));
  [10, 11, 6, 7, 34, 18, 22, 7, 8, 12, 13, 14, 14, 14, 11, 12, 18].forEach((w, i) => op.width(i, w));
  op.freeze = { r: 1 }; op.autoFilter = 'A1:Q' + (other.length + 1);

  /* ---------- Working Sheet ---------- */
  const ws = Sheet('Working Sheet');
  ws.set(0, 2, 'USD vendors', { b: true }); ws.set(0, 6, 'MVR vendors', { b: true });
  blocks.usd.grandNames.forEach((v, i) => ws.set(i + 1, 2, v));
  blocks.mvr.grandNames.forEach((v, i) => ws.set(i + 1, 6, v));
  ws.width(2, 40); ws.width(6, 40);

  /* ---------- Previous ---------- */
  const pv = Sheet('Previous');
  (prev.pd || []).forEach((row, r0) => { if (row) row.forEach((v, c) => {
    if (v === null || v === undefined || v === '') return;
    if (v instanceof Date) pv.set(r0, c, v, { fmt: DATE });
    else if (isNum(v) && (c === 5 || c === 6)) pv.set(r0, c, v, { fmt: NUM });
    else pv.set(r0, c, v);
  }); });
  pv.width(1, 36.7); pv.width(7, 28.4); for (const c of [2, 3, 4, 5, 6]) pv.width(c, 14);

  const sheets = [pivotSheets.usd, pivotSheets.mvr, pd, op, ws, pv];
  return { sheets, S, sst, summary };
}

async function writeXlsx(JSZip, model) {
  const { sheets, S, sst, summary } = buildWorkbookParts(model);
  const data = await packWorkbook(JSZip, sheets, S, sst, 'Weekly Pending Report');
  return { data, summary };
}
async function packWorkbook(JSZip, sheets, S, sst, title) {
  const zip = new JSZip(); let pk = 0;
  const ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'];
  const wbRels = [], sheetEls = [], caches = [];
  sheets.forEach((sh, i) => {
    const n = i + 1;
    zip.file('xl/worksheets/sheet' + n + '.xml', writeSheetXml(sh, S, sst));
    ct.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    wbRels.push('<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>');
    sheetEls.push('<sheet name="' + esc(sh.name) + '" sheetId="' + n + '" r:id="rId' + n + '"/>');
    const pivs = sh.pivots || (sh.pivot ? [sh.pivot] : []);
    const sheetRels = [];
    pivs.forEach((p0, j) => {
      const k = ++pk;
      const p = { ...p0, tbl: p0.tbl.replace(/cacheId="\d+"/, 'cacheId="' + k + '"') };
      zip.file('xl/pivotCache/pivotCacheDefinition' + k + '.xml', p.def);
      zip.file('xl/pivotCache/pivotCacheRecords' + k + '.xml', p.recs);
      zip.file('xl/pivotCache/_rels/pivotCacheDefinition' + k + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords' + k + '.xml"/></Relationships>');
      zip.file('xl/pivotTables/pivotTable' + k + '.xml', p.tbl);
      zip.file('xl/pivotTables/_rels/pivotTable' + k + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition' + k + '.xml"/></Relationships>');
      sheetRels.push('<Relationship Id="rId' + (j + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable' + k + '.xml"/>');
      ct.push('<Override PartName="/xl/pivotCache/pivotCacheDefinition' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
        '<Override PartName="/xl/pivotCache/pivotCacheRecords' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>',
        '<Override PartName="/xl/pivotTables/pivotTable' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>');
      const rid = 'rId' + (100 + k);
      wbRels.push('<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition' + k + '.xml"/>');
      caches.push('<pivotCache cacheId="' + k + '" r:id="' + rid + '"/>');
    });
    if (sheetRels.length) zip.file('xl/worksheets/_rels/sheet' + n + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheetRels.join('') + '</Relationships>');
  });
  wbRels.push('<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rId91" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>');
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr defaultThemeVersion="164011"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews><sheets>' + sheetEls.join('') + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/>' + (caches.length ? '<pivotCaches>' + caches.join('') + '</pivotCaches>' : '') + '</workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels.join('') + '</Relationships>');
  zip.file('xl/styles.xml', S.xml());
  zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sst.list.length + '" uniqueCount="' + sst.list.length + '">' + sst.list.map(s => '<si><t xml:space="preserve">' + esc(s) + '</t></si>').join('') + '</sst>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
  const now = new Date().toISOString().slice(0, 19) + 'Z';
  zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + esc(title) + '</dc:title><dc:creator>Weekly Report Tool</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + ct.join('') + '</Types>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}


/* ---------------- Pending & Clearing working workbooks ---------------- */
function sharedField(vals) {        // shared items (with index map) for a pivot axis/page field
  const keyOf = v => v === null || v === undefined || v === '' ? '\u0000m' : v instanceof Date ? 'd' + v.getTime() : (isNum(v) ? 'n' + v : 's' + v);
  const uniq = new Map();
  for (const v of vals) { const k = keyOf(v); if (!uniq.has(k)) uniq.set(k, v === '' || v === undefined ? null : v); }
  const items = [...uniq.values()];
  const str = items.filter(v => typeof v === 'string').sort(cmp), other = items.filter(v => typeof v !== 'string' && v !== null)
    .sort((a, b) => (a instanceof Date ? a.getTime() : a) - (b instanceof Date ? b.getTime() : b));
  const ordered = other.concat(str).concat(items.includes(null) ? [null] : []);
  const idx = new Map(ordered.map((v, i) => [keyOf(v), i]));
  const attrs = fieldInfo(ordered).replace('<sharedItems', '').replace('/>', '');
  const xml = '<sharedItems' + attrs + ' count="' + ordered.length + '">' + ordered.map(inlineVal).join('') + '</sharedItems>';
  return { ordered, xml, at: v => idx.get(keyOf(v)) };
}
function buildVendorPivot(o) {
  // o: { names, rows (arrays), vCol, aCol, pCol, sheet, ref, location, today, tableName }
  const n = o.names.length;
  const vf = sharedField(o.rows.map(r => r[o.vCol])), pf = sharedField(o.rows.map(r => r[o.pCol]));
  const cols = o.names.map(() => []);
  const recs = o.rows.map(r => {
    let x = '<r>';
    for (let i = 0; i < n; i++) {
      const v = r[i] === undefined ? null : r[i];
      if (i === o.vCol) x += '<x v="' + vf.at(v) + '"/>';
      else if (i === o.pCol) x += '<x v="' + pf.at(v) + '"/>';
      else { x += inlineVal(v); cols[i].push(v); }
    }
    return x + '</r>';
  });
  const cf = o.names.map((nm, i) => '<cacheField name="' + esc(nm) + '" numFmtId="' + (i === o.pCol ? 14 : 0) + '">' +
    (i === o.vCol ? vf.xml : i === o.pCol ? pf.xml : fieldInfo(cols[i])) + '</cacheField>').join('');
  const def = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Weekly Report Tool" refreshedDate="' + dateToSerial(o.today) + '" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="' + o.rows.length + '">' +
    '<cacheSource type="worksheet"><worksheetSource ref="' + o.ref + '" sheet="' + esc(o.sheet) + '"/></cacheSource><cacheFields count="' + n + '">' + cf + '</cacheFields></pivotCacheDefinition>';
  const recXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="' + o.rows.length + '">' + recs.join('') + '</pivotCacheRecords>';
  const items = list => '<items count="' + (list.length + 1) + '">' + list.map((v, i) => '<item x="' + i + '"/>').join('') + '<item t="default"/></items>';
  const pfx = o.names.map((nm, i) => i === o.vCol ? '<pivotField axis="axisRow" showAll="0">' + items(vf.ordered) + '</pivotField>' :
    i === o.pCol ? '<pivotField axis="axisPage" numFmtId="14" showAll="0">' + items(pf.ordered) + '</pivotField>' :
    i === o.aCol ? '<pivotField dataField="1" showAll="0"/>' : '<pivotField showAll="0"/>').join('');
  // vendor totals in item order
  const tot = new Map();
  o.rows.forEach(r => { const v = r[o.vCol] ?? null, a = r[o.aCol]; if (!tot.has(vf.at(v))) tot.set(vf.at(v), 0); if (isNum(a)) tot.set(vf.at(v), tot.get(vf.at(v)) + a); });
  const order = [...tot.keys()].sort((a, b) => a - b);
  const tl = parseInt(o.location.match(/\d+/)[0], 10), colA = o.location.match(/^[A-Z]+/)[0];
  const lastRow = tl + order.length + 1;
  const tbl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="' + o.tableName + '" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">' +
    '<location ref="' + colA + tl + ':' + colLetter(colIndex(colA) + 1) + lastRow + '" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"/>' +
    '<pivotFields count="' + n + '">' + pfx + '</pivotFields><rowFields count="1"><field x="' + o.vCol + '"/></rowFields>' +
    '<rowItems count="' + (order.length + 1) + '">' + order.map(i => '<i>' + (i ? '<x v="' + i + '"/>' : '<x/>') + '</i>').join('') + '<i t="grand"><x/></i></rowItems>' +
    '<colItems count="1"><i/></colItems><pageFields count="1"><pageField fld="' + o.pCol + '" hier="-1"/></pageFields>' +
    '<dataFields count="1"><dataField name="Sum of Actual" fld="' + o.aCol + '" baseField="0" baseItem="0" numFmtId="4"/></dataFields>' +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';
  const cells = [];   // [row(1-based), col offset, value, style]
  cells.push([tl - 2, 0, 'Posting Date', { b: true }], [tl - 2, 1, '(All)']);
  cells.push([tl, 0, 'Row Labels', { b: true }], [tl, 1, 'Sum of Actual', { b: true }]);
  let gt = 0;
  order.forEach((i, k) => { const v = vf.ordered[i]; cells.push([tl + 1 + k, 0, v === null ? '(blank)' : v], [tl + 1 + k, 1, r2(tot.get(i)), { fmt: NUM }]); gt += tot.get(i); });
  cells.push([lastRow, 0, 'Grand Total', { b: true }], [lastRow, 1, r2(gt), { b: true, fmt: NUM }]);
  return { pivot: { def, recs: recXml, tbl }, cells, colA: colIndex(colA), lastRow, total: r2(gt) };
}
function placePivot(sh, built) { for (const [r, c, v, st] of built.cells) sh.set(r - 1, built.colA + c, v, st); }

function sapHeaders(raw) {
  const h = (raw.rows[0] || []).map(cleanStr);
  let n = h.length; while (n > 0 && !h[n - 1]) n--;
  const names = [], seen = {};
  for (let i = 0; i < n; i++) { let nm = h[i] || ('Column ' + colLetter(i)); while (seen[nm.toLowerCase()]) nm += ' '; seen[nm.toLowerCase()] = 1; names.push(nm); }
  return names;
}
function rawSheet(name, raw, ncol) {
  const sh = Sheet(name);
  raw.rows.forEach((row, r) => {
    const filled = raw.fill[r] && raw.fill[r][0];
    for (let c = 0; c < ncol; c++) {
      const v = row ? row[c] : null; const st = {};
      if (r === 0) st.b = true;
      if (filled && r > 0) st.fill = 'FFFFFF99';
      if (v instanceof Date) st.fmt = DATE; else if (isNum(v) && Math.abs(v) >= 1000 && !Number.isInteger(v)) st.fmt = NUM;
      if (v !== null && v !== undefined) sh.set(r, c, v, Object.keys(st).length ? st : null); else if (st.fill) sh.set(r, c, null, st);
    }
  });
  sh.freeze = { r: 1 }; sh.autoFilter = 'A1:' + colLetter(ncol - 1) + Math.max(1, raw.rows.length);
  return sh;
}
function workRows(list, names, amtCol) {     // rows for the worked sheets: A = Vendor, N = Actual
  return list.map(x => {
    const row = []; for (let c = 0; c < names.length; c++) row[c] = x.it.raw[c] === undefined ? null : x.it.raw[c];
    row[0] = x.vendor; row[13] = x.amount;
    const ref = colLetter(amtCol);
    row.formula = x.how.k === 'abs' ? 'ABS(' + ref + '#)' : x.how.k === 'mul' ? 'ROUND(ABS(' + ref + '#)*' + x.how.r + ',2)' : 'ROUND(ABS(' + ref + '#)/' + x.how.r + ',2)';
    return row;
  });
}
function dataSheet(name, names, rows) {
  const sh = Sheet(name);
  names.forEach((h, c) => sh.set(0, c, c === 0 ? 'Vendor' : c === 13 ? 'Actual' : h, { b: true, fill: c === 0 || c === 13 ? 'FFFFF2CC' : 'FFD9D9D9' }));
  rows.forEach((row, i) => {
    for (let c = 0; c < names.length; c++) {
      const v = row[c];
      if (c === 13) { sh.set(i + 1, c, v, { fmt: NUM }, row.formula.replace('#', String(i + 2))); continue; }
      if (v === null || v === undefined) continue;
      sh.set(i + 1, c, v, v instanceof Date ? { fmt: DATE } : (isNum(v) && !Number.isInteger(v) ? { fmt: NUM } : null));
    }
  });
  sh.width(0, 36); sh.width(13, 15); sh.freeze = { r: 1 }; sh.autoFilter = 'A1:' + colLetter(names.length - 1) + (rows.length + 1);
  return sh;
}
function pivotRows(rows, names) { return rows.map(r => names.map((_, c) => r[c] === undefined ? null : r[c])); }
function addPivotSheet(sheets, name, srcName, names, rows, today, tableName) {
  const pv = Sheet(name);
  const b = buildVendorPivot({ names: names.map((h, c) => c === 0 ? 'Vendor' : c === 13 ? 'Actual' : h), rows: pivotRows(rows, names), vCol: 0, aCol: 13,
    pCol: names.findIndex(h => h.toLowerCase() === 'posting date'), sheet: srcName, ref: 'A1:' + colLetter(names.length - 1) + (rows.length + 1),
    location: 'A3', today, tableName });
  placePivot(pv, b); pv.pivots = [b.pivot]; pv.width(0, 44); pv.width(1, 18);
  sheets.push(pv); return b.total;
}

async function writeWorkingFiles(JSZip, model) {
  const today = model.opt.today, end = model.opt.periodEnd, asOf = new Date(end.getTime() + DAY);
  const out = {};
  /* Pending items */
  {
    const S = Styles(), sst = { list: [], map: {} }, sheets = [];
    const raw = model.openSh, names = sapHeaders(raw);
    const amtCol = names.findIndex(h => h.toLowerCase() === 'amount in doc. curr.');
    sheets.push(rawSheet('Sheet1', raw, names.length));
    const all = workRows(model.openRows, names, amtCol);
    sheets.push(dataSheet('Sheet2', names, all));
    const tot = {};
    for (const [blk, nm] of [['usd', 'USD'], ['mvr', 'MVR']]) {
      const rows = workRows(model.openRows.filter(x => x.blk === blk), names, amtCol);
      sheets.push(dataSheet(nm, names, rows));
      if (rows.length) tot[blk] = addPivotSheet(sheets, nm + ' PIVOT', nm, names, rows, today, nm + 'Pending');
    }
    // order like the working file: Sheet1, Sheet2, MVR PIVOT, MVR, USD PIVOT, USD
    const by = n => sheets.find(x => x.name === n);
    const ordered = ['Sheet1', 'Sheet2', 'MVR PIVOT', 'MVR', 'USD PIVOT', 'USD'].map(by).filter(Boolean);
    ordered[0].tabSelected = true;
    out.pending = { name: 'Open item till ' + ordinal(asOf.getUTCDate()) + ' ' + MON3[asOf.getUTCMonth()] + '.xlsx',
      data: await packWorkbook(JSZip, ordered, S, sst, 'Pending items'), totals: tot };
  }
  /* Cleared items */
  {
    const S = Styles(), sst = { list: [], map: {} }, sheets = [];
    const raw = model.clrSh, names = sapHeaders(raw);
    const amtCol = names.findIndex(h => h.toLowerCase() === 'amount in doc. curr.');
    sheets.push(rawSheet('Sheet1', raw, names.length));
    const tot = {};
    const mvr = workRows(model.clearRows.filter(x => x.blk === 'mvr'), names, amtCol);
    sheets.push(dataSheet('MVR', names, mvr));
    if (mvr.length) tot.mvr = addPivotSheet(sheets, 'MVR PIVOT', 'MVR', names, mvr, today, 'MVRCleared');
    const usd = workRows(model.clearRows.filter(x => x.blk === 'usd'), names, amtCol);
    const us = dataSheet('USD', names, usd); sheets.push(us);
    if (usd.length && usd.length <= 10) {          // few USD lines: pivot on the same sheet, under the list
      const start = usd.length + 5;
      const b = buildVendorPivot({ names: names.map((h, c) => c === 0 ? 'Vendor' : c === 13 ? 'Actual' : h), rows: pivotRows(usd, names), vCol: 0, aCol: 13,
        pCol: names.findIndex(h => h.toLowerCase() === 'posting date'), sheet: 'USD', ref: 'A1:' + colLetter(names.length - 1) + (usd.length + 1),
        location: 'A' + start, today, tableName: 'USDCleared' });
      placePivot(us, b); us.pivots = [b.pivot]; tot.usd = b.total; us.autoFilter = null;
    } else if (usd.length) tot.usd = addPivotSheet(sheets, 'USD PIVOT', 'USD', names, usd, today, 'USDCleared');
    if (model.clearExcluded.length) {
      const ex = Sheet('Removed lines');
      ex.set(0, 0, 'Left out of Weekly Clear (HR payroll postings, salary payments, dividend setoffs)', { b: true });
      names.forEach((h, c) => ex.set(1, c, c === 0 ? 'Vendor' : h, { b: true, fill: 'FFD9D9D9' }));
      model.clearExcluded.forEach((x, i) => names.forEach((_, c) => { const v = c === 0 ? x.vendor : x.it.raw[c];
        if (v !== null && v !== undefined) ex.set(i + 2, c, v, v instanceof Date ? { fmt: DATE } : (isNum(v) && !Number.isInteger(v) ? { fmt: NUM } : null)); }));
      ex.width(0, 36); sheets.push(ex);
    }
    const by = n => sheets.find(x => x.name === n);
    const ordered = ['Sheet1', 'MVR PIVOT', 'MVR', 'USD PIVOT', 'USD', 'Removed lines'].map(by).filter(Boolean);
    ordered[0].tabSelected = true;
    out.cleared = { name: 'Clearing ' + ordinal(end.getUTCDate()) + ' ' + MON3[end.getUTCMonth()] + '.xlsx',
      data: await packWorkbook(JSZip, ordered, S, sst, 'Cleared items'), totals: tot };
  }
  return out;
}

function defaultPeriodEnd(today) {      // most recent Wednesday before today
  const d = utcDay(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
  let back = (d.getUTCDay() - 3 + 7) % 7; if (back === 0) back = 7;
  return new Date(d.getTime() - back * DAY);
}
function fileName(end) { return 'Final Pending as of ' + ordinal(end.getUTCDate()) + ' ' + MON3[end.getUTCMonth()] + '.xlsx'; }

const CLEAR_EXCLUDE = ['HR Payroll Posting', 'Salary Payment', 'Dividend Setoff'];
const api = { CLEAR_EXCLUDE, AGE_LABELS, STATUSES, detect, analyse, buildModel, writeXlsx, writeWorkingFiles, matchStatus, defaultPeriodEnd, fileName,
  utcDay, toDate, longDate, dotDate, cleanStr, normAcct, DAY };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.WR = api;
})(typeof self !== 'undefined' ? self : this);
