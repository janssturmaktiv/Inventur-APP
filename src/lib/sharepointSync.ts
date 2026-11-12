// Arbeiten direkt in der gemeinsamen Excel-Liste in SharePoint.
// Jede Änderung betrifft nur einzelne Zellen, damit mehrere Personen gleichzeitig scannen können.
import { MAPPING_SHEET } from '../config';
import { addHiddenSheet, downloadBase64, listSheets, readRange, usedRangeValues, writeRange, type FileRef } from './msGraph';
import { parseTargetWorkbook } from './targetList';
import { normId } from './udi';
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
  return { items, source, mappings };
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
    if (batch && !batches.includes(batch)) {
      batches.push(batch);
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
    if (!batchStillUsed && batches.includes(batch)) {
      batches = batches.filter((b) => b !== batch);
      await writeRange(f, sheet, batchRef, [[batches.join(', ')]]);
    }
  }
  return { ist, batch: batches.join(', ') };
}

/** Spalte über alle Artikelzeilen in einem Rutsch schreiben; andere Zeilen bleiben unverändert. */
async function writeColumn(f: FileRef, source: TargetSource, col: string, items: TargetItem[], value: (current: unknown, item: TargetItem) => unknown) {
  const rows = items.map((i) => i.row!).filter(Boolean);
  if (!rows.length) return;
  const first = Math.min(...rows);
  const last = Math.max(...rows);
  const addr = `${col}${first}:${col}${last}`;
  const current = await readRange(f, source.sheetName, addr);
  const byRow = new Map(items.map((i) => [i.row, i]));
  const values = current.map((r, n) => {
    const item = byRow.get(first + n);
    return [item ? value(r[0], item) : r[0]];
  });
  await writeRange(f, source.sheetName, addr, values);
}

/** Inventur abschließen: leere Ist-Felder auf 0 (→ nachbestellen), Datum und Prüfer eintragen. */
export async function finishInventory(f: FileRef, source: TargetSource, items: TargetItem[], checker: string) {
  await writeColumn(f, source, source.istCol!, items, (cur) => (cur === '' || cur === null ? 0 : cur));
  const date = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  if (source.dateCell) await writeRange(f, source.sheetName, source.dateCell, [[date]]);
  if (source.checkerCell && checker) await writeRange(f, source.sheetName, source.checkerCell, [[checker]]);
}

/** Neue Inventur: Ist und Chargen-Nr. aller Artikelzeilen leeren. */
export async function resetInventory(f: FileRef, source: TargetSource, items: TargetItem[]) {
  await writeColumn(f, source, source.istCol!, items, () => '');
  if (source.batchCol) await writeColumn(f, source, source.batchCol, items, () => '');
}
