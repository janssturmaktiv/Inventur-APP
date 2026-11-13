import * as XLSX from 'xlsx';
import type { Reconciliation } from './reconcile';
import type { Scan, TargetSource } from './types';
import { writeCells, type CellUpdate } from './xlsxWriteback';
import { excelToday } from './excelDate';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const de = (iso: string) => new Date(iso).toLocaleString('de-DE');
const uniq = (xs: (string | undefined)[]) => [...new Set(xs.filter(Boolean))].join(', ');
const today = () => new Date().toISOString().slice(0, 10);

export type ExportFile = { blob: Blob; name: string };

function b64ToBlob(b64: string, type: string) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type });
}

export function canShareFile(f: ExportFile) {
  try {
    const file = new File([f.blob], f.name, { type: f.blob.type });
    return !!navigator.canShare?.({ files: [file] });
  } catch {
    return false;
  }
}

export async function shareFile(f: ExportFile) {
  const file = new File([f.blob], f.name, { type: f.blob.type });
  await navigator.share({ files: [file], title: f.name });
}

export function downloadFile(f: ExportFile) {
  const url = URL.createObjectURL(f.blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = f.name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** Trägt Ist-Menge, Chargen, Datum und Prüfer direkt in die importierte Original-Liste ein. */
export async function buildFilledOriginal(originalBase64: string, source: TargetSource, r: Reconciliation, checker: string): Promise<ExportFile> {
  const updates: CellUpdate[] = [];
  for (const row of r.rows) {
    if (!row.item.row) continue;
    if (source.istCol) updates.push({ ref: `${source.istCol}${row.item.row}`, value: row.foundQty });
    if (source.batchCol) updates.push({ ref: `${source.batchCol}${row.item.row}`, value: uniq(row.scans.map((s) => s.batch)) || null });
  }
  if (source.dateCell) updates.push({ ref: source.dateCell, value: excelToday() });
  if (source.checkerCell && checker) updates.push({ ref: source.checkerCell, value: checker });
  const b64 = await writeCells(originalBase64, source.sheetName, updates);
  const base = source.fileName.replace(/\.xlsx$/i, '');
  return { blob: b64ToBlob(b64, XLSX_MIME), name: `${base} - Inventur ${today()}.xlsx` };
}

function scanRow(s: Scan) {
  return {
    Charge: s.batch,
    'UDI-DI / GTIN': s.udiDi ?? '',
    Verfall: s.expiry ?? '',
    Seriennummer: s.serial ?? '',
    Menge: s.quantity,
    Zeitpunkt: de(s.scannedAt),
    'Erfasst von': s.scannedBy ?? '',
    Rohdaten: s.raw.replace(/\u001d/g, '<GS>'),
  };
}

/** Eigenständige Auswertung (neue Excel-Datei mit drei Blättern). */
export function buildReport(r: Reconciliation, scans: Scan[]): ExportFile {
  const wb = XLSX.utils.book_new();
  const abgleich = r.rows.map((row) => ({
    'Artikel-Nr.': row.item.articleNo,
    Artikelbezeichnung: row.item.name,
    Seite: row.item.side ?? '',
    Größe: row.item.size ?? '',
    Soll: row.item.expectedQty,
    Ist: row.foundQty,
    Differenz: row.foundQty - row.item.expectedQty,
    Status: row.status === 'fehlend' ? 'FEHLT' : row.status,
    'Chargen-Nr.': uniq(row.scans.map((s) => s.batch)),
    Verfall: uniq(row.scans.map((s) => s.expiry)),
  }));
  const ws1 = XLSX.utils.json_to_sheet(abgleich.length ? abgleich : [{ Hinweis: 'Keine Soll-Liste importiert' }]);
  ws1['!cols'] = [12, 30, 8, 7, 6, 6, 9, 12, 24, 14].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws1, 'Abgleich');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(r.unknown.length ? r.unknown.map(scanRow) : [{ Hinweis: 'Alle Scans zugeordnet' }]), 'Nicht zugeordnet');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(scans.length ? [...scans].reverse().map(scanRow) : [{ Hinweis: 'Keine Scans' }]), 'Alle Scans');
  const b64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
  return { blob: b64ToBlob(b64, XLSX_MIME), name: `Inventur-Auswertung ${today()}.xlsx` };
}

const esc = (s: unknown) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

/** Druckansicht öffnen (dort „Als PDF speichern“ wählen). */
export function printReport(r: Reconciliation, title: string, checker: string) {
  const color = { vollständig: '#1a7f37', teilweise: '#9a6700', fehlend: '#cf222e' } as const;
  const rows = r.rows
    .map(
      (row) => `<tr><td>${esc(row.item.articleNo)}</td><td>${esc(row.item.name)}</td><td>${esc(row.item.side)}</td><td>${esc(row.item.size)}</td>
        <td class="n">${row.item.expectedQty}</td><td class="n">${row.foundQty}</td><td>${esc(uniq(row.scans.map((s) => s.batch)))}</td>
        <td style="color:${color[row.status]};font-weight:600">${row.status === 'fehlend' ? 'FEHLT' : row.status}</td></tr>`,
    )
    .join('');
  const unknown = r.unknown
    .map((s) => `<tr><td>${esc(s.udiDi)}</td><td>${esc(s.batch)}</td><td>${esc(s.expiry)}</td><td class="n">${s.quantity}</td><td>${esc(de(s.scannedAt))}</td></tr>`)
    .join('');
  const html = `<!doctype html><html lang="de"><head><meta charset="utf-8"><title>${esc(title)}</title><style>
    body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;font-size:10pt;margin:24px;color:#111}
    h1{font-size:16pt;margin:0 0 4px} p{margin:0 0 12px;color:#555}
    table{border-collapse:collapse;width:100%;margin-bottom:18px} th,td{border:1px solid #bbb;padding:4px 6px;text-align:left;vertical-align:top}
    th{background:#1f3864;color:#fff} .n{text-align:right}
    @media print{button{display:none}}
  </style></head><body>
    <button onclick="print()" style="float:right;padding:8px 14px;font-size:12pt">Drucken / PDF</button>
    <h1>${esc(title)}</h1>
    <p>Stand ${esc(new Date().toLocaleString('de-DE'))}${checker ? ` · geprüft von ${esc(checker)}` : ''} · vollständig ${r.counts.vollständig} · teilweise ${r.counts.teilweise} · fehlend ${r.counts.fehlend} · nicht zugeordnet ${r.counts.unbekannt}</p>
    <table><tr><th>Artikel-Nr.</th><th>Artikelbezeichnung</th><th>Seite</th><th>Größe</th><th>Soll</th><th>Ist</th><th>Chargen-Nr.</th><th>Status</th></tr>${rows}</table>
    ${unknown ? `<h2 style="font-size:12pt">Nicht zugeordnete Scans</h2><table><tr><th>UDI-DI</th><th>Charge</th><th>Verfall</th><th>Menge</th><th>Zeitpunkt</th></tr>${unknown}</table>` : ''}
  </body></html>`;
  const w = window.open('', '_blank');
  if (!w) throw new Error('Pop-up wurde blockiert. Bitte Pop-ups für diese Seite erlauben.');
  w.document.write(html);
  w.document.close();
  setTimeout(() => w.print(), 400);
}
