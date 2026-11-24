# mpdf

Split a stack of invoices, delivery notes and POs into **one PDF per invoice**, each named by its invoice number. Everything runs in the browser — no files are uploaded anywhere.

## How it works
1. Add PDFs (any of three arrangements):
   - **Separate files, in order** — Invoice 1, DN 1, PO 1, Invoice 2… (groups every N files)
   - **One file per type** — all invoices, all DNs, all POs (set N = page N of each file)
   - **One combined PDF** — every N pages is a set
2. Names come from the first page of each set:
   - **Auto-detect** — "Invoice No: …", "Invoice #…", table headers, etc.
   - **Point to it** — click the invoice number on the first invoice; mpdf finds the same spot on the rest
   - **Pattern** — custom regular expression
3. Edit any name, then download a ZIP (or individual files).

Scanned PDFs without selectable text can't be read automatically — type the names in the table.

## Deploy (Cloudflare Pages)
Static site, no build step. In Cloudflare: **Workers & Pages → Create → Pages → Connect to Git** → pick `mpdf` →
- Framework preset: **None**
- Build command: *(empty)*
- Build output directory: `/`

## Local
```
python3 -m http.server 8000
```
then open http://localhost:8000

Libraries (bundled in `vendor/`): pdf-lib 1.17.1, pdf.js 6.2, JSZip 3.10.
