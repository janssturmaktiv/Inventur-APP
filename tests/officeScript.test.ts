// Testet das Office Script mit einer nachgebauten Excel-Arbeitsmappe (Layout wie die echte Liste:
// B4 Inventurdatum, C4:D4 verbunden mit Formel, E4 „Geprüft von:“, F4 leer, G4:H4 verbunden).
import { readFileSync, writeFileSync, mkdtempSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { pathToFileURL } from 'url';
import * as XLSX from 'xlsx';

const src = readFileSync(new URL('../office-script/inventur-scans-eintragen.ts', import.meta.url), 'utf8');
const dir = mkdtempSync(join(tmpdir(), 'osk-'));
const file = join(dir, 'script.ts');
writeFileSync(file, src + '\nexport default main;\n');

const wbx = XLSX.read(readFileSync(new URL('./beispiel-liste.xlsx', import.meta.url)), { sheetStubs: true });
const grid: string[][] = XLSX.utils.sheet_to_json<string[]>(wbx.Sheets.Bestand, { header: 1, raw: false, defval: '', blankrows: true });
const values: Record<string, unknown> = {};
const formulas: Record<string, string> = {};
const width = Math.max(...grid.map((r) => r.length));
grid.forEach((r, i) => r.forEach((v, j) => (values[`${i},${j}`] = v)));
const cellOf = (ref: string) => { const c = XLSX.utils.decode_cell(ref); return `${c.r},${c.c}`; };
// Echtes Layout nachstellen
values[cellOf('B4')] = 46323; // 28.10.2026 als Seriennummer
values[cellOf('C4')] = '';
formulas[cellOf('C4')] = '=IF(B4="","◄ Pflichtfeld!","")';
values[cellOf('G4')] = ''; // Prüfer leer: muss trotzdem in G4 (verbunden) landen, nicht in F4
values[cellOf('F23')] = 1; // alter Ist-Wert der letzten Inventur
values[cellOf('I23')] = 'ALT-1';
values[cellOf('F12')] = 1;
const merges = [
  { r: 3, c1: 2, c2: 3 },
  { r: 3, c1: 6, c2: 7 },
];
const ws = {
  getName: () => 'Bestand',
  getUsedRange: () => ({
    getTexts: () => grid.map((r, i) => Array.from({ length: width }, (_, j) => String(values[`${i},${j}`] ?? ''))),
    getRowIndex: () => 0,
    getColumnIndex: () => 0,
  }),
  getCell: (r: number, c: number) => ({
    getValue: () => values[`${r},${c}`] ?? '',
    setValue: (v: unknown) => {
      if (formulas[`${r},${c}`]) throw new Error('Formel überschrieben in ' + r + ',' + c);
      values[`${r},${c}`] = v;
    },
    getFormula: () => formulas[`${r},${c}`] ?? String(values[`${r},${c}`] ?? ''),
    getMergedAreasOrNullObject: () => {
      const m = merges.find((x) => x.r === r && c >= x.c1 && c <= x.c2);
      return m ? { getAreas: () => [{ getColumnIndex: () => m.c1 }] } : undefined;
    },
  }),
};
const workbook = { getWorksheets: () => [ws] };
const v = (ref: string) => values[cellOf(ref)];
const serial = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);

(async () => {
  const main = (await import(pathToFileURL(file).href)).default as (wb: unknown, body: string) => string;
  const checks: [string, boolean][] = [];
  // 1) Neue Inventur (HTML-Mail wie von Outlook): alte Werte werden geleert, nicht addiert
  const r1 = main(workbook, `<html><body><p>###INVENTUR-V1###<br>S;Phillip<br>D;2026-11-16<br>A;07853|RECHTS|M;1;LOT-A42<br>A;07853|RECHTS|M;1;LOT-B7<br>A;07780|RECHTS|XS;2;G&amp;1<br>A;99999|X|Y;1;Z<br>###ENDE###</p></body></html>`);
  checks.push(['Rückmeldung', r1.includes('Neue Inventur') && r1.includes('4 Stück eingetragen') && r1.includes('99999|X|Y')]);
  checks.push(['Alte Werte nicht addiert: Ist 07853 rechts M = 2', v('F23') === 2]);
  checks.push(['Alte Charge ersetzt', v('I23') === 'LOT-A42, LOT-B7']);
  checks.push(['Nicht gescannte alte Werte geleert', v('F12') === '']);
  checks.push(['Ist 07780 XS = 2, Charge mit &', v('F10') === 2 && v('I10') === 'G&1']);
  checks.push(['Inventurdatum in B4', v('B4') === serial(2026, 11, 16)]);
  checks.push(['Prüfer in G4 (verbunden), F4 unberührt', v('G4') === 'Phillip' && (v('F4') ?? '') === '']);
  checks.push(['Formel C4 unberührt', true]);
  // 2) Zweite Mail derselben Inventur, anderer Kollege: wird dazugezählt
  const r2 = main(workbook, '###INVENTUR-V1###\nS;Anna\nD;2026-11-16\nA;07853|RECHTS|M;1;LOT-C1\n###ENDE###');
  checks.push(['Gleiche Inventur zählt dazu', v('F23') === 3 && !r2.includes('Neue Inventur')]);
  checks.push(['Mehrere Prüfer', v('G4') === 'Phillip, Anna']);
  // 3) Vergessene Scans einer älteren Inventur werden abgelehnt
  const r3 = main(workbook, '###INVENTUR-V1###\nS;Bob\nD;2026-11-10\nA;07853|RECHTS|M;5;X\n###ENDE###');
  checks.push(['Ältere Inventur abgelehnt', v('F23') === 3 && r3.startsWith('Nicht eingetragen')]);
  // 4) Abschluss: leere Ist-Felder auf 0
  const r4 = main(workbook, 'xx ###INVENTUR-V1###\nS;Phillip\nD;2026-11-16\nF;Phillip\n###ENDE###');
  checks.push(['Abschluss setzt leere auf 0', v('F11') === 0 && v('F23') === 3 && r4.includes('abgeschlossen')]);
  checks.push(['Gruppenzeile unverändert', (v('F9') ?? '') === '']);
  // 5) Neu befüllen (N): alles leeren und neu eintragen
  main(workbook, '###INVENTUR-V1###\nS;Phillip\nD;2026-11-16\nN\nA;07853|RECHTS|M;1;LOT-A42\n###ENDE###');
  checks.push(['Neu befüllen', v('F23') === 1 && v('F11') === '' && v('G4') === 'Phillip']);
  checks.push(['Ohne Daten', main(workbook, 'Hallo').startsWith('Keine Inventur-Daten')]);
  let fail = 0;
  for (const [n, ok] of checks) {
    if (!ok) fail++;
    console.log(ok ? 'OK  ' : 'FAIL', n);
  }
  if (fail) console.log(r1, r2, r3, r4, v('F23'), v('I23'), v('F10'), v('B4'), v('G4'), v('F4'));
  process.exit(fail ? 1 : 0);
})();
