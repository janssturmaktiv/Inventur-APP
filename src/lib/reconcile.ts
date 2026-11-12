import { normId } from './udi';
import type { Mappings, Scan, TargetItem } from './types';

export type Status = 'vollständig' | 'teilweise' | 'fehlend';

export type ReconRow = {
  item: TargetItem;
  scans: Scan[];
  foundQty: number;
  status: Status;
};

export type Reconciliation = {
  rows: ReconRow[];
  unknown: Scan[]; // gescannt, aber keiner Soll-Zeile zugeordnet
  counts: Record<Status, number> & { unbekannt: number };
};

type ScanIds = Pick<Scan, 'udiDi' | 'productCode' | 'batch' | 'targetKey'>;

function productMatches(item: TargetItem, scan: ScanIds, mappings: Mappings): boolean {
  if (scan.targetKey) return scan.targetKey === item.key;
  const learned = mappings[normId(scan.udiDi)] ?? mappings[normId(scan.productCode)];
  if (learned) return learned === item.key;
  const target = normId(item.udiDi);
  if (target && (target === normId(scan.udiDi) || target === normId(scan.productCode))) return true;
  // HIBC: Herstellerartikelnummer steht im Code
  const art = normId(item.articleNo);
  return !!art && !item.size && !item.side && (art === normId(scan.productCode) || art === normId(scan.udiDi));
}

/** Passende Soll-Zeile für einen Scan. */
export function findTarget(targets: TargetItem[], scan: ScanIds, mappings: Mappings): TargetItem | undefined {
  const candidates = targets.filter((t) => productMatches(t, scan, mappings));
  return candidates.find((t) => t.batch && normId(t.batch) === normId(scan.batch)) ?? candidates.find((t) => !t.batch) ?? candidates[0];
}

export function reconcile(targets: TargetItem[], scans: Scan[], mappings: Mappings): Reconciliation {
  const rows: ReconRow[] = targets.map((item) => ({ item, scans: [], foundQty: 0, status: 'fehlend' }));
  const unknown: Scan[] = [];
  // Jeder Scan wird genau einer Zeile zugeordnet: zuerst Zeilen mit passender erwarteter Charge,
  // dann Zeilen ohne Chargenvorgabe (bevorzugt solche, die noch nicht voll sind).
  for (const scan of [...scans].reverse()) {
    const matching = rows.filter((r) => productMatches(r.item, scan, mappings));
    const byBatch = matching.filter((r) => r.item.batch && normId(r.item.batch) === normId(scan.batch));
    const open = (list: ReconRow[]) => list.find((r) => r.foundQty < r.item.expectedQty) ?? list[0];
    const target = open(byBatch) ?? open(matching.filter((r) => !r.item.batch));
    if (target) {
      target.scans.push(scan);
      target.foundQty += scan.quantity;
    } else unknown.push(scan);
  }
  for (const row of rows) {
    row.status = row.foundQty >= row.item.expectedQty ? 'vollständig' : row.foundQty > 0 ? 'teilweise' : 'fehlend';
  }
  const counts = { vollständig: 0, teilweise: 0, fehlend: 0, unbekannt: unknown.length };
  rows.forEach((r) => counts[r.status]++);
  return { rows, unknown: unknown.reverse(), counts };
}
