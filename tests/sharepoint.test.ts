// Testet die SharePoint-Anbindung gegen einen nachgebauten Microsoft-Graph-Server (Zellen im Speicher).
import { readFileSync } from 'fs';
import { __setTokenProvider, type FileRef } from '../src/lib/msGraph';
import { addToRow, finishInventory, loadShared, removeFromRow, resetInventory, saveMapping } from '../src/lib/sharepointSync';

const xlsx = readFileSync(new URL('./beispiel-liste.xlsx', import.meta.url));
const cells: Record<string, Record<string, unknown>> = { Bestand: {} };
const sheets = [{ name: 'Bestand', visibility: 'Visible' }];
const calls: string[] = [];

function expand(address: string): string[][] {
  const [a, b = a] = address.split(':');
  const p = (r: string) => {
    const m = /^([A-Z]+)(\d+)$/.exec(r)!;
    return { c: m[1].charCodeAt(0) - 64, r: Number(m[2]) };
  };
  const s = p(a), e = p(b);
  const out: string[][] = [];
  for (let r = s.r; r <= e.r; r++) {
    const row: string[] = [];
    for (let c = s.c; c <= e.c; c++) row.push(String.fromCharCode(64 + c) + r);
    out.push(row);
  }
  return out;
}

__setTokenProvider(async () => 'TEST');
globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
  const url = decodeURIComponent(String(input));
  const method = init?.method ?? 'GET';
  calls.push(`${method} ${url.replace('https://graph.microsoft.com/v1.0', '')}`);
  const json = (o: unknown, status = 200) => new Response(JSON.stringify(o), { status, headers: { 'Content-Type': 'application/json' } });
  if (url.endsWith('/content')) return new Response(xlsx);
  if (/\/worksheets\?/.test(url)) return json({ value: sheets });
  if (/\/worksheets\/add$/.test(url)) {
    const { name } = JSON.parse(String(init!.body));
    sheets.push({ name, visibility: 'Visible' });
    cells[name] = {};
    return json({ name });
  }
  const sm = /worksheets\('(.+?)'\)(.*)$/.exec(url);
  if (sm) {
    const sheet = sm[1].replace(/''/g, "'");
    const rest = sm[2];
    if (rest === '' && method === 'PATCH') {
      sheets.find((s) => s.name === sheet)!.visibility = JSON.parse(String(init!.body)).visibility;
      return json({});
    }
    if (rest.startsWith('/usedRange')) {
      const keys = Object.keys(cells[sheet]);
      if (!keys.length) return json({ values: [] });
      const maxRow = Math.max(...keys.map((k) => Number(k.replace(/[A-Z]+/, ''))));
      return json({ values: expand(`A1:D${maxRow}`).map((r) => r.map((ref) => cells[sheet][ref] ?? '')) });
    }
    const am = /range\(address='([A-Z0-9:]+)'\)/.exec(rest)!;
    const grid = expand(am[1]);
    if (method === 'PATCH') {
      const { values } = JSON.parse(String(init!.body));
      grid.forEach((r, i) => r.forEach((ref, j) => (cells[sheet][ref] = typeof values[i][j] === 'string' ? values[i][j].replace(/^'/, '') : values[i][j])));
      return json({});
    }
    return json({ values: grid.map((r) => r.map((ref) => cells[sheet][ref] ?? '')) });
  }
  return json({ error: { message: 'unbekannt ' + url } }, 404);
}) as typeof fetch;

(async () => {
  const f: FileRef = { driveId: 'D1', itemId: 'I1', name: 'beispiel-liste.xlsx' };
  const st = await loadShared(f);
  const checks: [string, boolean][] = [['11 Zeilen geladen', st.items.length === 11 && st.source.istCol === 'F']];
  const m = st.items.find((i) => i.key === '07853|RECHTS|M')!;
  let r = await addToRow(f, st.source, m.row!, 1, 'LOT-A');
  r = await addToRow(f, st.source, m.row!, 1, 'LOT-B'); // Kollege scannt zweites Stück
  r = await addToRow(f, st.source, m.row!, 1, 'LOT-A');
  checks.push(['Ist hochgezählt', cells.Bestand[`F${m.row}`] === 3 && r.ist === 3]);
  checks.push(['Chargen gesammelt', cells.Bestand[`I${m.row}`] === 'LOT-A, LOT-B']);
  r = await removeFromRow(f, st.source, m.row!, 1, 'LOT-B', false);
  checks.push(['Austragen', cells.Bestand[`F${m.row}`] === 2 && cells.Bestand[`I${m.row}`] === 'LOT-A']);
  await saveMapping(f, '04049999123457', m);
  await saveMapping(f, '04049999123457', m); // zweites Mal: keine Dublette
  checks.push(['Zuordnungsblatt versteckt angelegt', sheets.some((s) => s.name === 'App-Zuordnungen' && s.visibility === 'Hidden')]);
  checks.push(['Zuordnung einmal gespeichert', cells['App-Zuordnungen'].A2 === '04049999123457' && cells['App-Zuordnungen'].A3 === undefined]);
  const st2 = await loadShared(f);
  checks.push(['Zuordnung wird geladen', st2.mappings['4049999123457'] === '07853|RECHTS|M']);
  await finishInventory(f, st.source, st.items, 'Phillip');
  const istVals = st.items.map((i) => cells.Bestand[`F${i.row}`]);
  checks.push(['Abschluss: leere auf 0, gezählte bleiben', istVals.filter((v) => v === 0).length === 10 && cells.Bestand[`F${m.row}`] === 2]);
  checks.push(['Gruppenzeilen unberührt', cells.Bestand.F15 === '' || cells.Bestand.F15 === undefined]);
  checks.push(['Datum und Prüfer', /\d\d\.\d\d\.\d{4}/.test(String(cells.Bestand.B4)) && cells.Bestand.G4 === 'Phillip']);
  await resetInventory(f, st.source, st.items);
  checks.push(['Neue Inventur leert', st.items.every((i) => cells.Bestand[`F${i.row}`] === '' && cells.Bestand[`I${i.row}`] === '')]);
  let fail = 0;
  for (const [n, ok] of checks) {
    if (!ok) fail++;
    console.log(ok ? 'OK  ' : 'FAIL', n);
  }
  if (fail) console.log(calls.join('\n'));
  process.exit(fail ? 1 : 0);
})().catch((e) => {
  console.error(e);
  console.log(calls.join('\n'));
  process.exit(1);
});
