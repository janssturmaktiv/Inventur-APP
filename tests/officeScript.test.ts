// Testet das Office Script mit einer nachgebauten Excel-Arbeitsmappe.
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
const width = Math.max(...grid.map((r) => r.length));
grid.forEach((r, i) => r.forEach((v, j) => (values[`${i},${j}`] = v)));
const ws = {
  getName: () => 'Bestand',
  getUsedRange: () => ({
    getTexts: () => grid.map((r, i) => Array.from({ length: width }, (_, j) => String(values[`${i},${j}`] ?? ''))),
    getRowIndex: () => 0,
    getColumnIndex: () => 0,
  }),
  getCell: (r: number, c: number) => ({
    getValue: () => values[`${r},${c}`] ?? '',
    setValue: (v: unknown) => (values[`${r},${c}`] = v),
  }),
};
const workbook = { getWorksheets: () => [ws] };

(async () => {
  const main = (await import(pathToFileURL(file).href)).default as (wb: unknown, body: string) => string;
  // HTML-Mail wie von Outlook geliefert
  const body = `<html><body><p>Gesendet aus der Konsilager-App</p><p>###INVENTUR-V1###<br>S;Phillip<br>A;07853|RECHTS|M;1;LOT-A42<br>A;07853|RECHTS|M;1;LOT-B7<br>A;07780|RECHTS|XS;2;G&amp;1<br>A;99999|X|Y;1;Z<br>###ENDE###</p></body></html>`;
  const r1 = main(workbook, body);
  const v = (ref: string) => {
    const c = XLSX.utils.decode_cell(ref);
    return values[`${c.r},${c.c}`];
  };
  // Zeilen: 10 = 07780 XS, 23 = 07853 rechts M (1-basiert)
  const checks: [string, boolean][] = [
    ['Rückmeldung', r1.startsWith('4 Stück eingetragen') && r1.includes('99999|X|Y')],
    ['Ist 07853 rechts M = 2', v('F23') === 2],
    ['Chargen gesammelt', v('I23') === 'LOT-A42, LOT-B7'],
    ['Ist 07780 XS = 2, Charge mit &', v('F10') === 2 && v('I10') === 'G&1'],
  ];
  const r2 = main(workbook, 'xx ###INVENTUR-V1###\nF;Phillip\n###ENDE###');
  checks.push(['Abschluss setzt leere auf 0', v('F11') === 0 && v('F23') === 2 && r2.includes('abgeschlossen')]);
  checks.push(['Gruppenzeile unverändert', (v('F9') ?? '') === '']);
  checks.push(['Datum und Prüfer', Number(v('B4')) > 46000 && v('G4') === 'Phillip']);
  checks.push(['Ohne Daten', main(workbook, 'Hallo').startsWith('Keine Inventur-Daten')]);
  let fail = 0;
  for (const [n, ok] of checks) {
    if (!ok) fail++;
    console.log(ok ? 'OK  ' : 'FAIL', n);
  }
  if (fail) console.log(r1, r2, v('F23'), v('I23'), v('F10'), v('B4'), v('G4'));
  process.exit(fail ? 1 : 0);
})();
