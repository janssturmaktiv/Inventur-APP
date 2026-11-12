import { readFileSync } from 'fs';
import * as XLSX from 'xlsx';
import { parseTargetWorkbook } from '../src/lib/targetList';
import { writeCells } from '../src/lib/xlsxWriteback';

(async () => {
  const b64 = readFileSync(new URL('./beispiel-liste.xlsx', import.meta.url)).toString('base64');
  const { items, source } = parseTargetWorkbook(b64, 'beispiel-liste.xlsx');
  const checks: [string, boolean][] = [
    ['11 Artikelzeilen', items.length === 11],
    ['erste Zeile 10 = 07780 rechts XS', items[0].row === 10 && items[0].key === '07780|RECHTS|XS'],
    ['Spalten erkannt', source.istCol === 'F' && source.batchCol === 'I' && source.dateCell === 'B4' && source.checkerCell === 'G4'],
  ];
  const out = await writeCells(b64, source.sheetName, [
    { ref: 'F10', value: 1 },
    { ref: 'I10', value: 'LOT-1 & 2' },
    { ref: 'F11', value: 0 },
  ]);
  const wb = XLSX.read(out, { type: 'base64', cellFormula: true, sheetStubs: true });
  const ws = wb.Sheets[source.sheetName];
  checks.push(['Ist geschrieben', ws.F10?.v === 1 && ws.F11?.v === 0]);
  checks.push(['Charge geschrieben', ws.I10?.v === 'LOT-1 & 2']);
  checks.push(['Formel erhalten', String(ws.H10?.f ?? '').includes('nachbestellen')]);
  checks.push(['Titel erhalten', String(ws.A1?.v).startsWith('Konsignationslager')]);
  let fail = 0;
  for (const [name, ok] of checks) {
    if (!ok) fail++;
    console.log(ok ? 'OK  ' : 'FAIL', name);
  }
  process.exit(fail ? 1 : 0);
})();
