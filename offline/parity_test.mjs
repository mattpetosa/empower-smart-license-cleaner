// Runs the JS core over both fixtures (PDF via pdf.js in Node, txt directly)
// and diffs the rows against the Python implementation's output.
import { createRequire } from 'node:module';
import { readFileSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
const require = createRequire(import.meta.url);
const ESLC = require('./core.js');
const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');

async function pdfText(bytes) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), useWorkerFetch: false, isEvalSupported: false }).promise;
  const lines = [];
  for (let p = 1; p <= doc.numPages; p++) lines.push(...ESLC.itemsToLines((await (await doc.getPage(p)).getTextContent()).items));
  return lines.join('\n');
}
const norm = r => ({ installation: r.installation, printed: r.printed, company: r.company, support_id: r.support_id,
  licenses: r.licenses.map(l => [l.category, l.name, l.qty, l.qty_label, l.serial]), removed: r.removed.map(x => x.reason) });

let failed = 0;
for (const f of ['fixtures/sample-empower370.pdf', 'fixtures/sample-checksum.txt']) {
  const bytes = readFileSync(path.join(root, f));
  const text = f.endsWith('.pdf') ? await pdfText(bytes) : bytes.toString('utf8');
  const js = norm(ESLC.parseText(text));
  const py = JSON.parse(execFileSync(path.join(root, 'venv/bin/python'), ['-c', `
import json,sys; sys.path.insert(0,'${root}')
from licenses import parse_upload; r=parse_upload(open('${path.join(root, f)}','rb').read())
print(json.dumps({'installation':r.installation,'printed':r.printed,'company':r.company,'support_id':r.support_id,
 'licenses':[[l.category,l.name,l.qty,l.qty_label,l.serial] for l in r.licenses],'removed':[x.reason for x in r.removed]}))`]).toString());
  const a = JSON.stringify(js), b = JSON.stringify(py);
  if (a === b) console.log(`OK   ${f}: ${js.licenses.length} rows match Python`);
  else { failed++; console.log(`FAIL ${f}`); writeFileSync('/tmp/js.json', JSON.stringify(js, null, 1)); writeFileSync('/tmp/py.json', JSON.stringify(py, null, 1)); }
}
// xlsx sanity: write one and let Python/openpyxl open it
const wb = ESLC.buildWorkbook(ESLC.parseText(readFileSync(path.join(root, 'fixtures/sample-checksum.txt'), 'utf8')), [['Company', 'Acme'], ['Support ID', 'SR1']]);
const out = path.join(root, 'offline/.parity.xlsx'); writeFileSync(out, wb);
console.log(execFileSync(path.join(root, 'venv/bin/python'), ['-c', `
from openpyxl import load_workbook; wb=load_workbook('${out}'); ws=wb['Licenses']
print('xlsx opens:', wb.sheetnames, ws.max_row, 'rows; A1=', ws['A1'].value, '; header row', ws.freeze_panes, ws.auto_filter.ref)`]).toString().trim());
process.exit(failed ? 1 : 0);
