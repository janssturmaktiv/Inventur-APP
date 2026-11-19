// Lagerinventur in SharePoint: Ordner, Inventurdatei (Blätter Übersicht / Scans / App-Artikel), Artikelstamm.
import * as XLSX from 'xlsx';
import { graph, downloadBase64, type FileRef } from './msGraph';
import { ARTICLE_HEADERS, SCAN_HEADERS, SUMMARY_HEADERS, articleToRow, colLetter, rowToArticle, rowToScan, scanToRow, summarize, summaryValues, type Article, type LagerScan } from './lagerData';

export type FolderRef = { driveId: string; itemId: string; name: string; webUrl?: string };

type DriveItem = { id: string; name: string; webUrl?: string; folder?: unknown; file?: unknown; lastModifiedDateTime?: string; parentReference?: { driveId: string } };

export const INVENTORY_PREFIX = 'Lagerinventur ';
export const SUMMARY_SHEET = 'Übersicht';
export const SCAN_SHEET = 'Scans';
export const ARTICLE_SHEET = 'App-Artikel';
const SCAN_TABLE = 'Scans';
const ARTICLE_TABLE = 'AppArtikel';

function shareId(url: string) {
  const b64 = btoa(unescape(encodeURIComponent(url))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${b64}`;
}

/** Ordner über einen SharePoint-Link („Link kopieren“ oder Adresse aus dem Browser) finden. */
export async function resolveFolder(url: string): Promise<FolderRef> {
  // redeemSharingLink: Wer nicht Mitglied der Website ist, bekommt über den Freigabelink dauerhaft Zugriff auf den Ordner
  const it = await graph<DriveItem>(`/shares/${shareId(url.trim())}/driveItem?$select=id,name,webUrl,folder,parentReference`, { headers: { Prefer: 'redeemSharingLink' } });
  if (!it.folder || !it.parentReference?.driveId) throw new Error('Der Link zeigt nicht auf einen Ordner.');
  return { driveId: it.parentReference.driveId, itemId: it.id, name: it.name, webUrl: it.webUrl };
}

async function children(folder: FolderRef): Promise<DriveItem[]> {
  const out: DriveItem[] = [];
  let next: string | undefined = `/drives/${folder.driveId}/items/${folder.itemId}/children?$select=id,name,webUrl,file,folder,lastModifiedDateTime&$top=500`;
  while (next) {
    const r: { value: DriveItem[]; '@odata.nextLink'?: string } = await graph(next);
    out.push(...r.value);
    next = r['@odata.nextLink'];
  }
  return out;
}

export type FolderContents = { inventories: FileRef[]; catalog?: FileRef };

/** Inventurdateien („Lagerinventur JJJJ-MM-TT.xlsx“, neueste zuerst) und Artikelstamm im Ordner. */
export async function scanFolder(folder: FolderRef): Promise<FolderContents> {
  const items = await children(folder);
  const ref = (x: DriveItem): FileRef => ({ driveId: folder.driveId, itemId: x.id, name: x.name, webUrl: x.webUrl });
  const inventories = items
    .filter((x) => x.file && x.name.startsWith(INVENTORY_PREFIX) && /\.xlsx$/i.test(x.name))
    .sort((a, b) => b.name.localeCompare(a.name))
    .map(ref);
  // Artikelstamm: Excel/CSV mit „Artikel“, „EAN“ oder „GTIN“ im Namen (keine Inventurdatei)
  const cat = items
    .filter((x) => x.file && !x.name.startsWith(INVENTORY_PREFIX) && /\.(xlsx|xls|csv)$/i.test(x.name) && /artikel|ean|gtin|acriba/i.test(x.name))
    .sort((a, b) => (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? ''))[0];
  return { inventories, catalog: cat ? ref(cat) : undefined };
}

/** Artikelstamm herunterladen und als Zeilen liefern (erstes Blatt). */
export async function loadCatalogRows(f: FileRef): Promise<unknown[][]> {
  const b64 = await downloadBase64(f);
  const wb = XLSX.read(b64, { type: 'base64', dense: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: '' });
}

const wbPath = (f: FileRef) => `/drives/${f.driveId}/items/${f.itemId}/workbook`;
const wsPath = (f: FileRef, sheet: string) => `${wbPath(f)}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;

/** Neue Inventurdatei anlegen. Selbst angelegte Artikel der letzten Inventur werden übernommen. */
export async function createInventory(folder: FolderRef, date: string, carryArticles: Article[]): Promise<FileRef> {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(summaryValues([], `Lagerinventur ${date.split('-').reverse().join('.')}`, '–')), SUMMARY_SHEET);
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([SCAN_HEADERS]), SCAN_SHEET);
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([ARTICLE_HEADERS, ...carryArticles.map((a) => articleToRow(a).map((x) => String(x).replace(/^'/, '')))]), ARTICLE_SHEET);
  book.Workbook = { Sheets: [{}, {}, { Hidden: 1 }] };
  const bytes = XLSX.write(book, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;

  let name = `${INVENTORY_PREFIX}${date}.xlsx`;
  let item: DriveItem | undefined;
  for (let n = 2; !item; n++) {
    try {
      item = await graph<DriveItem>(
        `/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent(name)}:/content?@microsoft.graph.conflictBehavior=fail`,
        { method: 'PUT', body: bytes as unknown as BodyInit, headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } },
      );
    } catch (e) {
      if (String((e as Error).message).startsWith('409') && n < 10) name = `${INVENTORY_PREFIX}${date} (${n}).xlsx`;
      else throw e;
    }
  }
  const f: FileRef = { driveId: folder.driveId, itemId: item.id, name: item.name, webUrl: item.webUrl };
  // Eine frisch hochgeladene Datei ist für Excel Online manchmal erst nach ein paar Sekunden bereit
  const retry = async <T>(fn: () => Promise<T>): Promise<T> => {
    for (let i = 0; ; i++) {
      try {
        return await fn();
      } catch (e) {
        if (i >= 4 || !/^(404|409|423|500|502)/.test(String((e as Error).message))) throw e;
        await new Promise((r) => setTimeout(r, 1500 * (i + 1)));
      }
    }
  };
  // Excel-Tabellen anlegen: Zeilen anhängen ist dann auch bei mehreren Handys gleichzeitig sicher
  const lastScan = colLetter(SCAN_HEADERS.length);
  const t1 = await retry(() => graph<{ id: string }>(`${wbPath(f)}/tables/add`, { method: 'POST', body: JSON.stringify({ address: `${SCAN_SHEET}!A1:${lastScan}1`, hasHeaders: true }) }));
  await graph(`${wbPath(f)}/tables/${t1.id}`, { method: 'PATCH', body: JSON.stringify({ name: SCAN_TABLE }) });
  const t2 = await graph<{ id: string }>(`${wbPath(f)}/tables/add`, {
    method: 'POST',
    body: JSON.stringify({ address: `'${ARTICLE_SHEET}'!A1:${colLetter(ARTICLE_HEADERS.length)}${1 + Math.max(carryArticles.length, 0)}`, hasHeaders: true }),
  });
  await graph(`${wbPath(f)}/tables/${t2.id}`, { method: 'PATCH', body: JSON.stringify({ name: ARTICLE_TABLE }) });
  await graph(`${wsPath(f, SUMMARY_SHEET)}/range(address='A4:${colLetter(SUMMARY_HEADERS.length)}4')/format/font`, { method: 'PATCH', body: JSON.stringify({ bold: true }) }).catch(() => {});
  await graph(`${wsPath(f, SUMMARY_SHEET)}/range(address='A1')/format/font`, { method: 'PATCH', body: JSON.stringify({ bold: true, size: 14 }) }).catch(() => {});
  return f;
}

async function tableValues(f: FileRef, table: string): Promise<unknown[][]> {
  const r = await graph<{ values: unknown[][] }>(`${wbPath(f)}/tables('${table}')/dataBodyRange?$select=values`);
  return r.values ?? [];
}

export async function readScans(f: FileRef): Promise<LagerScan[]> {
  return (await tableValues(f, SCAN_TABLE)).map(rowToScan).filter((x): x is LagerScan => !!x);
}

export async function readArticles(f: FileRef): Promise<Article[]> {
  return (await tableValues(f, ARTICLE_TABLE)).map(rowToArticle).filter((x): x is Article => !!x);
}

export async function appendScans(f: FileRef, scans: LagerScan[]) {
  if (!scans.length) return;
  await graph(`${wbPath(f)}/tables('${SCAN_TABLE}')/rows/add`, { method: 'POST', body: JSON.stringify({ values: scans.map(scanToRow) }) });
}

export async function appendArticle(f: FileRef, a: Article) {
  await graph(`${wbPath(f)}/tables('${ARTICLE_TABLE}')/rows/add`, { method: 'POST', body: JSON.stringify({ values: [articleToRow(a)] }) });
}

/** Scan anhand seiner ID aus der Tabelle löschen. */
export async function deleteScan(f: FileRef, id: string): Promise<boolean> {
  const rows = await tableValues(f, SCAN_TABLE);
  const idx = rows.findIndex((r) => String(r[10] ?? '').trim() === id);
  if (idx < 0) return false;
  await graph(`${wbPath(f)}/tables('${SCAN_TABLE}')/rows/itemAt(index=${idx})`, { method: 'DELETE' });
  return true;
}

/** Blatt „Übersicht“ aus allen Scans neu schreiben. */
export async function writeSummary(f: FileRef, scans: LagerScan[], title: string) {
  const stand = new Date().toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' });
  const values = summaryValues(summarize(scans), title, stand);
  const last = colLetter(SUMMARY_HEADERS.length);
  await graph(`${wsPath(f, SUMMARY_SHEET)}/range(address='A5:${last}5000')/clear`, { method: 'POST', body: JSON.stringify({ applyTo: 'Contents' }) });
  await graph(`${wsPath(f, SUMMARY_SHEET)}/range(address='A1:${last}${values.length}')`, { method: 'PATCH', body: JSON.stringify({ values }) });
}
