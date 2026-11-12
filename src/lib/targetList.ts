import * as XLSX from 'xlsx';
import type { TargetItem, TargetSource } from './types';

// Spalten werden anhand der Überschrift erkannt (Groß-/Kleinschreibung egal).
const COLS = {
  articleNo: /^(art(ikel)?[\s.\-_]*(nr|nummer|no)\.?$|artikel$|ref(erenz)?(nr)?\.?$|material(nummer)?$)/i,
  name: /^(artikel)?(bezeichnung|beschreibung|name|produkt)/i,
  side: /^(seite|links\/rechts|l\/r)$/i,
  size: /^(gr(ö|oe)(ß|ss)e|gr\.|size)$/i,
  udiDi: /(gtin|udi|ean|hibc|pzn)/i,
  expectedQty: /^(soll|soll[\s-]*(menge|bestand)|menge|anzahl|bestand)$/i,
  ist: /^(ist|ist[\s-]*(menge|bestand)|gezählt)$/i,
  batchOut: /^charge(n)?[\s.\-]*(nr|nummer)?\.?$|^lot$/i,
  batchExpected: /charge.*soll|soll.*charge/i,
} as const;
type ColKey = keyof typeof COLS;

export type ParsedTargets = { items: TargetItem[]; source: TargetSource };

export function parseTargetWorkbook(base64: string, fileName: string): ParsedTargets {
  const wb = XLSX.read(base64, { type: 'base64' });
  // Erstes Blatt mit erkennbarer Kopfzeile nehmen
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const ref = sheet['!ref'];
    if (!ref) continue;
    const startRow = XLSX.utils.decode_range(ref).s.r; // 0-basiert
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: false, defval: '', blankrows: true });
    const cell = (r: unknown[], i?: number) => (i === undefined ? '' : String(r[i] ?? '').trim());

    const headerIdx = rows.findIndex((r) => r.some((c) => COLS.articleNo.test(String(c).trim())) && r.some((c) => COLS.name.test(String(c).trim())));
    if (headerIdx === -1) continue;
    const header = rows[headerIdx].map((c) => String(c).trim());
    const idx: Partial<Record<ColKey, number>> = {};
    (Object.keys(COLS) as ColKey[]).forEach((k) => {
      const i = header.findIndex((h, j) => h && COLS[k].test(h) && !Object.values(idx).includes(j));
      if (i !== -1) idx[k] = i;
    });

    const items: TargetItem[] = [];
    rows.slice(headerIdx + 1).forEach((r, n) => {
      const articleNo = cell(r, idx.articleNo);
      const name = cell(r, idx.name);
      const qtyText = cell(r, idx.expectedQty);
      // Gruppen-Überschriften („07780 · GENUDYN CI · Rechts“) und Leerzeilen überspringen
      if (!articleNo || !name) return;
      if (idx.expectedQty !== undefined && !/^\d+([.,]\d+)?$/.test(qtyText)) return;
      const side = cell(r, idx.side) || undefined;
      const size = cell(r, idx.size) || undefined;
      const excelRow = startRow + headerIdx + 1 + n + 1; // 1-basiert
      items.push({
        id: `r${excelRow}`,
        key: [articleNo, side, size].map((x) => (x ?? '').toUpperCase()).join('|'),
        articleNo,
        name,
        side,
        size,
        udiDi: cell(r, idx.udiDi) || undefined,
        batch: cell(r, idx.batchExpected) || undefined,
        expectedQty: Number(qtyText.replace(',', '.')) || 1,
        row: excelRow,
      });
    });
    if (!items.length) continue;

    // Felder „Inventurdatum“ und „Geprüft von“ oberhalb der Kopfzeile suchen (Wert = nächste befüllte/gelbe Zelle rechts)
    const findLabelCell = (re: RegExp) => {
      for (let ri = 0; ri < headerIdx; ri++) {
        const ci = rows[ri].findIndex((c) => re.test(String(c).trim()));
        if (ci !== -1) {
          // nächste Spalte rechts, die nicht leer ist (oder direkt daneben)
          let vi = ci + 1;
          for (let k = ci + 1; k < Math.min(rows[ri].length, ci + 4); k++) {
            if (String(rows[ri][k] ?? '').trim()) { vi = k; break; }
          }
          return XLSX.utils.encode_cell({ r: startRow + ri, c: vi });
        }
      }
      return undefined;
    };
    const col = (i?: number) => (i === undefined ? undefined : XLSX.utils.encode_col(i));
    return {
      items,
      source: {
        fileName,
        sheetName,
        istCol: col(idx.ist),
        batchCol: col(idx.batchOut),
        dateCell: findLabelCell(/^inventurdatum/i),
        checkerCell: findLabelCell(/^gepr(ü|ue)ft von/i),
      },
    };
  }
  throw new Error('Keine passende Liste gefunden. Erwartet wird eine Kopfzeile mit „Artikel-Nr.“ und „Artikelbezeichnung“ (optional Seite, Größe, Soll, Ist, Chargen-Nr.).');
}
