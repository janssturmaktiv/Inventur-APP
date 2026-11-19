// Arbeiten direkt in der gemeinsamen Excel-Liste in SharePoint.
// Jede Änderung betrifft nur einzelne Zellen, damit mehrere Personen gleichzeitig scannen können.
import { MAPPING_SHEET } from '../config';
import { addHiddenSheet, downloadBase64, listSheets, readRange, usedRangeValues, writeRange, type FileRef } from './msGraph';
import { parseTargetWorkbook } from './targetList';
import { normId } from './udi';
import { excelToday } from './excelDate';
import { itemLabel, type Mappings, type TargetItem, type TargetSource } from './types';

export type SharedState = { items: TargetItem[]; source: TargetSource; mappings: Mappings };

const splitBatches = (s: unknown) =>
  String(s ?? '')
    .split(/[,;\n]/)
    .map((x) => x.trim())
    .filter(Boolean);

export async function loadShared(f: FileRef): Promise<SharedState> {
  const b64 = await downloadBase64(f);
  const { items, source } = parseTargetWorkbook(b64, f.name);
  if (!source.istCol) throw new Error('In der Liste wurde keine Spalte „Ist“ gefunden.');
  const mappings = await loadMappings(f).catch(() => ({}));
  // Die heruntergeladene Datei kann ein paar Minuten hinter Excel Online zurückliegen.
  // Ist und Chargen deshalb immer live über die Excel-Schnittstelle lesen.
  await readLiveValues(f, source, items);
  return { items, source, mappings };
}

/** Zusammenhängende Blöcke von Artikelzeilen (ohne Gruppen-Überschriften dazwischen). */
function rowBlocks(items: TargetItem[]): [number, number][] {
  const rows = [...new Set(items.map((i) => i.row!).filter(Boolean))].sort((a, b) => a - b);
  const out: [number, number][] = [];
  for (let i = 0; i < rows.length; ) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    out.push([rows[i], rows[j]]);
    i = j + 1;
  }
  return out;
}

async function readLiveValues(f: FileRef, source: TargetSource, items: TargetItem[]) {
  if (!source.istCol) return;
  const byRow = new Map(items.filter((i) => i.row).map((i) => [i.row!, i]));
  const blocks = rowBlocks(items);
  if (!blocks.length) return;
  const first = blocks[0][0];
  const last = blocks[blocks.length - 1][1];
  const ist = await readRange(f, source.sheetName, `${source.istCol}${first}:${source.istCol}${last}`);
  const bat = source.batchCol ? await readRange(f, source.sheetName, `${source.batchCol}${first}:${source.batchCol}${last}`) : [];
  for (let r = first; r <= last; r++) {
    const it = byRow.get(r);
    if (!it) continue;
    const v = ist[r - first]?.[0];
    it.currentIst = v === '' || v === null || v === undefined ? null : Number(v);
    it.currentBatch = String(bat[r - first]?.[0] ?? '');
  }
}

async function loadMappings(f: FileRef): Promise<Mappings> {
  const sheets = await listSheets(f);
  if (!sheets.some((s) => s.name === MAPPING_SHEET)) return {};
  const rows = await usedRangeValues(f, MAPPING_SHEET);
  const m: Mappings = {};
  for (const r of rows.slice(1)) {
    const id = normId(String(r[0] ?? ''));
    const key = String(r[1] ?? '').trim();
    if (id && key) m[id] = key;
  }
  return m;
}

/** Gelernte Zuordnung für alle speichern (verstecktes Blatt in der Datei). */
export async function saveMapping(f: FileRef, udiDi: string, item: TargetItem) {
  const sheets = await listSheets(f);
  if (!sheets.some((s) => s.name === MAPPING_SHEET)) {
    await addHiddenSheet(f, MAPPING_SHEET);
    await writeRange(f, MAPPING_SHEET, 'A1:D1', [['UDI-DI / GTIN', 'Schlüssel (Artikel|Seite|Größe)', 'Artikel', 'Gelernt am']]);
  }
  const rows = await usedRangeValues(f, MAPPING_SHEET);
  const id = normId(udiDi);
  let rowNo = rows.findIndex((r, i) => i > 0 && normId(String(r[0] ?? '')) === id) + 1; // 1-basiert
  if (rowNo <= 0) rowNo = Math.max(rows.length, 1) + 1;
  // als Text speichern (führende Nullen der GTIN bleiben erhalten)
  await writeRange(f, MAPPING_SHEET, `A${rowNo}:D${rowNo}`, [[`'${udiDi}`, item.key, itemLabel(item), new Date().toLocaleString('de-DE')]]);
}

/** Einen Scan in die Zeile eintragen: Ist um die Menge erhöhen, Charge ergänzen. Liefert die neuen Werte. */
export async function addToRow(f: FileRef, source: TargetSource, row: number, qty: number, batch: string) {
  const sheet = source.sheetName;
  const istRef = `${source.istCol}${row}`;
  const [[istVal]] = await readRange(f, sheet, istRef);
  const ist = (Number(istVal) || 0) + qty;
  await writeRange(f, sheet, istRef, [[ist]]);
  let batches: string[] = [];
  if (source.batchCol) {
    const batchRef = `${source.batchCol}${row}`;
    const [[bVal]] = await readRange(f, sheet, batchRef);
    batches = splitBatches(bVal);
    if (batch) {
      // Jedes Stück einzeln: zwei Schienen aus derselben Charge = Chargennummer zweimal
      for (let i = 0; i < Math.max(1, Math.round(qty)); i++) batches.push(batch);
      await writeRange(f, sheet, batchRef, [[batches.join(', ')]]);
    }
  }
  return { ist, batch: batches.join(', ') };
}

/** Einen gelöschten Scan wieder austragen. */
export async function removeFromRow(f: FileRef, source: TargetSource, row: number, qty: number, batch: string, batchStillUsed: boolean) {
  const sheet = source.sheetName;
  const istRef = `${source.istCol}${row}`;
  const [[istVal]] = await readRange(f, sheet, istRef);
  const ist = Math.max(0, (Number(istVal) || 0) - qty);
  await writeRange(f, sheet, istRef, [[ist]]);
  let batches: string[] = [];
  if (source.batchCol) {
    const batchRef = `${source.batchCol}${row}`;
    const [[bVal]] = await readRange(f, sheet, batchRef);
    batches = splitBatches(bVal);
    // je Stück eine Chargennummer wieder entfernen
    let changed = false;
    for (let i = 0; i < Math.max(1, Math.round(qty)); i++) {
      const at = batches.lastIndexOf(batch);
      if (at < 0) break;
      batches.splice(at, 1);
      changed = true;
    }
    void batchStillUsed;
    if (changed) await writeRange(f, sheet, batchRef, [[batches.join(', ')]]);
  }
  return { ist, batch: batches.join(', ') };
}

/** Spalte für Artikelzeilen schreiben. Je zusammenhängendem Zeilenblock ein Aufruf, damit verbundene
 *  Gruppen-Überschriften dazwischen nie berührt werden. Liefert die Zeilen, die sich geändert haben. */
async function writeColumn(f: FileRef, source: TargetSource, col: string, items: TargetItem[], value: (current: unknown, item: TargetItem) => unknown) {
  const byRow = new Map(items.filter((i) => i.row).map((i) => [i.row!, i]));
  const rows = [...byRow.keys()].sort((a, b) => a - b);
  const changed: { row: number; value: unknown }[] = [];
  for (let i = 0; i < rows.length; ) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    const first = rows[i];
    const last = rows[j];
    const addr = `${col}${first}:${col}${last}`;
    const current = await readRange(f, source.sheetName, addr);
    let dirty = false;
    const values = current.map((r, n) => {
      const v = value(r[0], byRow.get(first + n)!);
      if (v !== r[0]) {
        dirty = true;
        changed.push({ row: first + n, value: v });
      }
      return [v];
    });
    if (dirty) await writeRange(f, source.sheetName, addr, values);
    i = j + 1;
  }
  return changed;
}

const isEmpty = (v: unknown) => v === '' || v === null || v === undefined;
const toSerial = (iso: string) => {
  const [y, m, d] = iso.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
};

async function addChecker(f: FileRef, source: TargetSource, name: string) {
  if (!source.checkerCell || !name.trim()) return;
  const [[cur]] = await readRange(f, source.sheetName, source.checkerCell);
  const names = String(cur ?? '').split(',').map((x) => x.trim()).filter(Boolean);
  const add = name.split(',').map((x) => x.trim()).filter((x) => x && !names.includes(x));
  if (add.length) await writeRange(f, source.sheetName, source.checkerCell, [[[...names, ...add].join(', ')]]);
}

/**
 * Vor dem ersten Scan einer Inventur aufrufen. Steht in der Liste ein anderes Inventurdatum als der Tag
 * des ersten Scans, werden Ist und Chargen-Nr. geleert und das Datum gesetzt. Der Name kommt zu „Geprüft von“.
 * Liefert 'reset', 'ok' oder 'older' (Scans gehören zu einer älteren Inventur als die Liste).
 */
export async function ensureInventory(f: FileRef, source: TargetSource, items: TargetItem[], inventoryDate: string, checker: string): Promise<'reset' | 'ok' | 'older'> {
  const want = toSerial(inventoryDate);
  let result: 'reset' | 'ok' = 'ok';
  if (source.dateCell) {
    const [[cur]] = await readRange(f, source.sheetName, source.dateCell);
    const have = Math.floor(Number(cur) || 0);
    if (have && want < have) return 'older';
    if (have !== want) {
      await resetInventory(f, source, items);
      if (source.checkerCell) await writeRange(f, source.sheetName, source.checkerCell, [['']]);
      await writeRange(f, source.sheetName, source.dateCell, [[want]]);
      result = 'reset';
    }
  }
  await addChecker(f, source, checker);
  return result;
}

/** Alle Größen/Seiten eines gescannten Artikels ohne Ist-Wert auf 0 setzen (→ „nachbestellen“). */
export async function zeroFillArticle(f: FileRef, source: TargetSource, items: TargetItem[], articleNo: string) {
  const same = items.filter((i) => normId(i.articleNo) === normId(articleNo));
  const changed: { row: number; value: unknown }[] = [];
  for (const [a, b] of rowBlocks(same)) {
    const cur = await readRange(f, source.sheetName, `${source.istCol}${a}:${source.istCol}${b}`);
    // nur die leeren Zellen einzeln setzen: so wird nie ein Wert überschrieben, den gerade jemand anderes einträgt
    for (let r = a; r <= b; r++) {
      if (!isEmpty(cur[r - a]?.[0])) continue;
      await writeRange(f, source.sheetName, `${source.istCol}${r}`, [[0]]);
      changed.push({ row: r, value: 0 });
    }
  }
  return changed;
}

/** Inventur abschließen: leere Ist-Felder auf 0 (→ nachbestellen), Datum und Prüfer eintragen. */
export async function finishInventory(f: FileRef, source: TargetSource, items: TargetItem[], checker: string) {
  await writeColumn(f, source, source.istCol!, items, (cur) => (isEmpty(cur) ? 0 : cur));
  if (source.dateCell) {
    const [[cur]] = await readRange(f, source.sheetName, source.dateCell);
    if (!Number(cur)) await writeRange(f, source.sheetName, source.dateCell, [[excelToday()]]);
  }
  await addChecker(f, source, checker);
}

/** Neue Inventur: Ist und Chargen-Nr. aller Artikelzeilen leeren. */
export async function resetInventory(f: FileRef, source: TargetSource, items: TargetItem[]) {
  await writeColumn(f, source, source.istCol!, items, () => '');
  if (source.batchCol) await writeColumn(f, source, source.batchCol, items, () => '');
}
