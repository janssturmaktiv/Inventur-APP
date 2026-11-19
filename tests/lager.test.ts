// Lagerinventur: Artikelstamm (acriba-Format), Zusammenfassung und SharePoint-Ablauf gegen einen nachgebauten Graph-Server.
import * as XLSX from 'xlsx';
import { __setTokenProvider } from '../src/lib/msGraph';
import { parseCatalog, productGroup, summarize, summaryValues, type LagerScan } from '../src/lib/lagerData';
import { appendArticle, appendScans, createInventory, deleteScan, loadCatalogRows, readArticles, readScans, resolveFolder, scanFolder, writeSummary } from '../src/lib/lagerGraph';

const checks: [string, boolean][] = [];
const check = (n: string, ok: boolean) => checks.push([n, ok]);

// ---------- Artikelstamm wie aus acriba ----------
const acriba = [
  ['Eancode', 'Groesse', 'Herstellername', 'Name1', 'Name2', 'Artikelnummer', 'Herstellerartikelnummer'],
  ['', '43', 'Hartmann', 'Secosol BASE - F', '', 737, ''],
  ['4046963817251', 'S', 'Sturm aktiv GmbH', 'Einmalhandschuhe Größe S', '', 362, '...'],
  ['4046963817312', 'M', 'Sturm aktiv GmbH', 'Einmalhandschuhe Größe M', '', 361, '...'],
  ['4029925258897', '2', 'Sporlastic GMBH', 'GENU-TEX OA laterale Ausführung, links', 'Softorthese links schwarz 2', 1, '07382-1-9000-2'],
  ['01042503049023191722', '1St', 'UroVision GmbH', 'VISISOFT BALLONK', '', 5, 'KI-1'],
  ['4046963817251', 'S', 'Doppelt', 'Doppelt', '', 9, ''],
];
const cat = parseCatalog(acriba);
check('Artikelstamm: Spalten erkannt', cat.columns.gtin === 'Eancode' && cat.columns.hersteller === 'Herstellername' && cat.columns.size === 'Groesse');
check('Artikelstamm: 4 Artikel mit EAN, Doppelte ignoriert', cat.withGtin === 4 && cat.total === 6);
const h = cat.articles.get('4046963817251')!;
check('Handschuhe: Gruppe ohne Größe, Art.-Nr. acriba', h.produktart === 'Einmalhandschuhe' && h.artNr === '362' && h.hersteller === 'Sturm aktiv GmbH');
const sp = cat.articles.get('4029925258897')!;
check('Sporlastic: Herstellerartikelnummer, Gruppe ohne Seite', sp.artNr === '07382-1-9000-2' && sp.produktart === 'GENU-TEX OA laterale Ausführung');
check('Ganzer UDI-Code in EAN-Spalte', cat.articles.get('4250304902319')?.hersteller === 'UroVision GmbH');
check('productGroup', productGroup('PATELLADYN Kniebandage Gr. 4', '4') === 'PATELLADYN Kniebandage');

// ---------- Zusammenfassung ----------
const sc = (id: string, a: Partial<LagerScan>): LagerScan => ({ id, time: '2026-11-19T12:00:00Z', person: 'P', hersteller: '', produktart: '', artNr: '', name: '', gtin: '', batch: '', expiry: '', qty: 1, ...a });
const scans = [
  sc('1', { hersteller: 'Sturm aktiv GmbH', produktart: 'Einmalhandschuhe', name: 'Einmalhandschuhe Größe S', gtin: '4046963817251', batch: 'A' }),
  sc('2', { hersteller: 'Sturm aktiv GmbH', produktart: 'Einmalhandschuhe', name: 'Einmalhandschuhe Größe S', gtin: '04046963817251', batch: 'A', expiry: '2027-01-01' }),
  sc('3', { hersteller: 'Bort GmbH', produktart: 'Bandage', name: 'X', gtin: '1' }),
  sc('4', { hersteller: 'Sturm aktiv GmbH', produktart: 'Einmalhandschuhe', name: 'Einmalhandschuhe Größe M', gtin: '4046963817312' }),
];
const g = summarize(scans);
check('Gruppen sortiert nach Hersteller', g.map((x) => x.hersteller).join() === 'Bort GmbH,Sturm aktiv GmbH');
check('Gleiche GTIN zusammengezählt, Charge je Stück', g[1].qty === 3 && g[1].items[1].qty === 2 && g[1].items[1].batches.join() === 'A,A' && g[1].items[1].expiry === '2027-01-01');
const sv = summaryValues(g, 'Lagerinventur', 'jetzt');
check('Übersicht: Gesamt 4', sv[sv.length - 1][0] === 'Gesamt' && sv[sv.length - 1][5] === 4);

// ---------- Nachgebauter Graph-Server ----------
const catalogBytes = XLSX.write((() => { const b = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(acriba), 'EAN'); return b; })(), { type: 'array', bookType: 'xlsx' });
const files: Record<string, { name: string; bytes?: ArrayBuffer }> = { cat1: { name: 'EAN Nummern Artikel.xlsx', bytes: catalogBytes } };
const tables: Record<string, { name: string; rows: unknown[][] }> = {};
const sheetsWritten: Record<string, unknown[][]> = {};
let nextId = 1;
__setTokenProvider(async () => 'TEST');
globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
  const url = decodeURIComponent(String(input)).replace('https://graph.microsoft.com/v1.0', '');
  const method = init?.method ?? 'GET';
  const body = init?.body && typeof init.body === 'string' ? JSON.parse(init.body) : undefined;
  const json = (o: unknown, status = 200) => new Response(JSON.stringify(o), { status, headers: { 'Content-Type': 'application/json' } });
  if (url.startsWith('/shares/')) return json({ id: 'F1', name: 'Lagerinventur', folder: {}, parentReference: { driveId: 'D1' } });
  if (/\/items\/F1\/children/.test(url)) return json({ value: Object.entries(files).map(([id, f]) => ({ id, name: f.name, file: {}, lastModifiedDateTime: '2026-11-19' })) });
  const put = /\/items\/F1:\/(.+):\/content/.exec(url);
  if (put && method === 'PUT') {
    const name = put[1];
    if (Object.values(files).some((f) => f.name === name)) return json({ error: { message: 'exists' } }, 409);
    const id = 'X' + nextId++;
    files[id] = { name };
    return json({ id, name, webUrl: 'https://x/' + name });
  }
  const content = /\/items\/(\w+)\/content$/.exec(url);
  if (content) return new Response(files[content[1]].bytes);
  if (/\/tables\/add$/.test(url)) {
    const id = 'T' + nextId++;
    const m = /!A1:[A-Z]+(\d+)$/.exec(body.address)!;
    tables[id] = { name: id, rows: Array.from({ length: Math.max(1, Number(m[1]) - 1) }, () => []) };
    return json({ id });
  }
  const tp = /\/tables\/(T\d+)$/.exec(url);
  if (tp && method === 'PATCH') { tables[tp[1]].name = body.name; return json({}); }
  const tn = /\/tables\('(\w+)'\)\/(.*)$/.exec(url);
  if (tn) {
    const t = Object.values(tables).find((x) => x.name === tn[1])!;
    if (tn[2].startsWith('dataBodyRange')) return json({ values: t.rows.length ? t.rows : [['']] });
    if (tn[2] === 'rows/add') { t.rows.push(...body.values); return json({}); }
    const del = /rows\/itemAt\(index=(\d+)\)/.exec(tn[2]);
    if (del && method === 'DELETE') { t.rows.splice(Number(del[1]), 1); return new Response(null, { status: 204 }); }
  }
  const ws = /worksheets\('(.+?)'\)\/range\(address='([^']+)'\)(.*)$/.exec(url);
  if (ws) {
    if (ws[3] === '/clear') return json({});
    if (ws[3] === '/format/font') return json({});
    if (method === 'PATCH') { sheetsWritten[ws[1]] = body.values; return json({}); }
  }
  return json({ error: { message: 'unbekannt ' + method + ' ' + url } }, 404);
}) as typeof fetch;

(async () => {
  const folder = await resolveFolder('https://sturmaktiv.sharepoint.com/:f:/s/TeamOT/abc');
  let c = await scanFolder(folder);
  check('Ordner: Artikelstamm gefunden, noch keine Inventur', c.catalog?.name === 'EAN Nummern Artikel.xlsx' && c.inventories.length === 0);
  const rows = await loadCatalogRows(c.catalog!);
  check('Artikelstamm aus SharePoint gelesen', parseCatalog(rows).withGtin === 4);
  const f = await createInventory(folder, '2026-11-19', [{ gtin: '999', artNr: 'A', name: 'Eigen', hersteller: 'H', produktart: 'P' }]);
  check('Inventurdatei angelegt', f.name === 'Lagerinventur 2026-11-19.xlsx');
  c = await scanFolder(folder);
  check('Neueste Inventur wird gefunden', c.inventories[0]?.name === 'Lagerinventur 2026-11-19.xlsx');
  // Leere Startzeile der Tabelle wird ignoriert
  check('Leere Datei: keine Scans', (await readScans(f)).length === 0);
  // Artikel aus der letzten Inventur übernommen (Tabelle wurde beim Anlegen mit dieser Zeile erstellt)
  Object.values(tables).find((t) => t.name === 'AppArtikel')!.rows[0] = ["'999", 'A', 'Eigen', 'H', 'P'];
  await appendArticle(f, { gtin: '04012345678901', artNr: '', name: 'Neu', hersteller: 'H2', produktart: 'P2' });
  const arts = await readArticles(f);
  check('Eigene Artikel gespeichert', arts.length === 2 && arts[1].gtin === '04012345678901');
  await appendScans(f, scans);
  const back = await readScans(f);
  check('Scans geschrieben und gelesen (GTIN ohne Apostroph)', back.length === 4 && back[1].gtin === '04046963817251' && back[0].batch === 'A');
  await deleteScan(f, '3');
  const after = await readScans(f);
  check('Scan gelöscht', after.length === 3 && !after.some((s) => s.id === '3'));
  await writeSummary(f, after, 'Lagerinventur 2026-11-19');
  const sum = sheetsWritten['Übersicht'];
  check('Übersicht geschrieben', !!sum && sum[0][0] === 'Lagerinventur 2026-11-19' && sum.some((r) => String(r[0]).includes('Sturm aktiv GmbH  ·  Einmalhandschuhe')));

  let fail = 0;
  for (const [n, ok] of checks) {
    if (!ok) fail++;
    console.log(ok ? 'OK  ' : 'FAIL', n);
  }
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
