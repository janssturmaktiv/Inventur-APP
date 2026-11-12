// Ende-zu-Ende ohne Netz: App erzeugt E-Mail-Text -> Office Script trägt ihn in die Liste ein.
import { readFileSync, writeFileSync, mkdtempSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { pathToFileURL } from 'url';
import * as XLSX from 'xlsx';
import { parseTargetWorkbook } from '../src/lib/targetList';
import { buildMailBatches, mailtoUrl } from '../src/lib/mailSync';
import type { Scan } from '../src/lib/types';

const buf = readFileSync(new URL('./beispiel-liste.xlsx', import.meta.url));
const { items } = parseTargetWorkbook(buf.toString('base64'), 'x.xlsx');
const scans: Scan[] = [];
for (let i = 0; i < 120; i++) {
  scans.push({ id: 's' + i, batch: `LOT-${i % 7}`, quantity: 1, raw: '', codeType: 'qr', scannedAt: '', udiDi: i % 2 ? '04049999123457' : '04049999000001' });
}
scans.push({ id: 'x', batch: 'NOPE', quantity: 1, raw: '', codeType: 'qr', scannedAt: '', udiDi: '0999' }); // nicht zugeordnet
const mappings = { '4049999123457': '07853|RECHTS|M', '4049999000001': '07780|RECHTS|XS' };
const batches = buildMailBatches(scans, items, mappings, 'Phillip; Test', true);

const src = readFileSync(new URL('../office-script/inventur-scans-eintragen.ts', import.meta.url), 'utf8');
const dir = mkdtempSync(join(tmpdir(), 'osk-'));
writeFileSync(join(dir, 's.ts'), src + '\nexport default main;\n');
const grid: string[][] = XLSX.utils.sheet_to_json<string[]>(XLSX.read(buf, { sheetStubs: true }).Sheets.Bestand, { header: 1, raw: false, defval: '', blankrows: true });
const values: Record<string, unknown> = {};
const width = Math.max(...grid.map((r) => r.length));
grid.forEach((r, i) => r.forEach((v, j) => (values[`${i},${j}`] = v)));
const ws = {
  getName: () => 'Bestand',
  getUsedRange: () => ({ getTexts: () => grid.map((r, i) => Array.from({ length: width }, (_, j) => String(values[`${i},${j}`] ?? ''))), getRowIndex: () => 0, getColumnIndex: () => 0 }),
  getCell: (r: number, c: number) => ({ getValue: () => values[`${r},${c}`] ?? '', setValue: (v: unknown) => (values[`${r},${c}`] = v) }),
};

(async () => {
  const main = (await import(pathToFileURL(join(dir, 's.ts')).href)).default as (wb: unknown, body: string) => string;
  // Outlook liefert HTML: Zeilenumbrüche als <br>, Sonderzeichen kodiert
  for (const b of batches) main({ getWorksheets: () => [ws] }, '<div>' + b.body.replace(/&/g, '&amp;').replace(/\n/g, '<br>') + '</div>');
  const v = (ref: string) => { const c = XLSX.utils.decode_cell(ref); return values[`${c.r},${c.c}`]; };
  const checks: [string, boolean][] = [
    ['mehrere E-Mails', batches.length > 1 && batches.every((b) => b.body.length < 1700)],
    ['Betreff nummeriert', batches[0].subject === `INVENTUR-SCANS (1/${batches.length})`],
    ['alle zugeordneten Scans enthalten', batches.reduce((n, b) => n + b.scanIds.length, 0) === 120],
    ['mailto-Link', mailtoUrl('a@b.de', batches[0]).startsWith('mailto:a%40b.de?subject=INVENTUR-SCANS')],
    ['Ist 07853 rechts M = 60', v('F23') === 60],
    ['Ist 07780 XS = 60', v('F10') === 60],
    ['7 Chargen gesammelt', String(v('I23')).split(', ').length === 7],
    ['Abschluss nur einmal, Rest auf 0', v('F11') === 0],
    ['Prüfer ohne Semikolon', v('G4') === 'Phillip, Test'],
  ];
  let fail = 0;
  for (const [n, ok] of checks) { if (!ok) fail++; console.log(ok ? 'OK  ' : 'FAIL', n); }
  if (fail) console.log(batches.length, v('F23'), v('F10'), v('I23'), v('G4'));
  process.exit(fail ? 1 : 0);
})();
