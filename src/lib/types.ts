/** Ein Artikel (eine Zeile), der laut Soll-Liste im Konsilager liegen sollte. */
export type TargetItem = {
  id: string;
  key: string; // stabile Kennung: Artikelnr|Seite|Größe (bleibt bei neuem Import gleich)
  articleNo: string;
  name: string;
  side?: string;
  size?: string;
  udiDi?: string; // GTIN bzw. HIBC-Produktkennung, falls die Liste eine solche Spalte hat
  batch?: string; // optional: erwartete Charge
  expectedQty: number;
  row?: number; // Zeilennummer in der Excel-Datei (1-basiert)
  currentIst?: number | null; // aktueller Wert der Spalte „Ist“ (null = leer)
  currentBatch?: string; // aktueller Wert der Spalte „Chargen-Nr.“
};

/** Woher die Soll-Liste stammt und wohin zurückgeschrieben wird. */
export type TargetSource = {
  fileName: string;
  sheetName: string;
  istCol?: string; // z. B. "F"
  batchCol?: string; // z. B. "I"
  dateCell?: string; // z. B. "B4" (Inventurdatum)
  checkerCell?: string; // z. B. "G4" (Geprüft von)
};

/** Ein gescannter (oder manuell erfasster) Artikel. */
export type Scan = {
  id: string;
  udiDi?: string;
  productCode?: string;
  batch: string;
  expiry?: string;
  serial?: string;
  quantity: number;
  raw: string;
  codeType: string;
  scannedAt: string; // ISO
  scannedBy?: string;
  targetKey?: string; // manuell gewählte Soll-Zeile
  sync?: 'ok' | 'pending' | 'unassigned'; // Stand in der SharePoint-Liste
  syncedRow?: number; // in diese Zeile eingetragen
};

/** Gelernte Zuordnung: UDI-DI -> Soll-Zeile (key). */
export type Mappings = Record<string, string>;

export const itemLabel = (t: TargetItem) => [t.articleNo, t.name, t.side, t.size].filter(Boolean).join(' · ');
