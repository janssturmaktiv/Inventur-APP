// Lagerinventur: Artikelstamm (z. B. acriba-Export), Scan-Zeilen und Zusammenfassung nach Hersteller/Produktart.
import { normId } from './udi';
import { GTIN_ZUORDNUNG } from '../data/gtinZuordnung';
import { STANDARD_LISTE } from '../data/standardListe';

export type Article = { gtin: string; artNr: string; name: string; hersteller: string; produktart: string };

export type LagerScan = {
  id: string;
  time: string; // ISO
  person: string;
  hersteller: string;
  produktart: string;
  artNr: string;
  name: string;
  gtin: string;
  batch: string;
  expiry: string; // JJJJ-MM-TT
  qty: number;
};

/** Spalten des Blatts „Scans“ (Reihenfolge = Excel-Spalten A–K). */
export const SCAN_HEADERS = ['Zeitpunkt', 'Person', 'Hersteller', 'Produktart', 'Artikel-Nr.', 'Bezeichnung', 'GTIN', 'Charge', 'Verfall', 'Menge', 'Scan-ID'];
/** Spalten des versteckten Blatts mit selbst angelegten Artikeln. */
export const ARTICLE_HEADERS = ['GTIN', 'Artikel-Nr.', 'Bezeichnung', 'Hersteller', 'Produktart'];

const fmtTime = (iso: string) => {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
};

// GTIN als Text mit Apostroph, damit Excel keine Zahl mit wissenschaftlicher Schreibweise daraus macht
export const scanToRow = (s: LagerScan): (string | number)[] => [
  fmtTime(s.time), s.person, s.hersteller, s.produktart, s.artNr, s.name, s.gtin ? "'" + s.gtin : '', s.batch, s.expiry, s.qty, s.id,
];

export function rowToScan(r: unknown[]): LagerScan | null {
  const t = (i: number) => String(r[i] ?? '').trim();
  if (!t(10)) return null; // leere Tabellenzeile
  return { time: t(0), person: t(1), hersteller: t(2), produktart: t(3), artNr: t(4), name: t(5), gtin: t(6).replace(/^'/, ''), batch: t(7), expiry: t(8), qty: Number(r[9]) || 1, id: t(10) };
}

export const articleToRow = (a: Article) => [a.gtin ? "'" + a.gtin : '', a.artNr, a.name, a.hersteller, a.produktart];
export function rowToArticle(r: unknown[]): Article | null {
  const t = (i: number) => String(r[i] ?? '').trim();
  const a = { gtin: t(0).replace(/^'/, ''), artNr: t(1), name: t(2), hersteller: t(3), produktart: t(4) };
  return a.gtin || a.artNr ? a : null;
}

// ---------- Artikelstamm einlesen ----------

const norm = (s: unknown) => String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');

/** GTIN/EAN aus einer Zelle (Zahl oder Text, auch „4,03E+12“ bei Zahlenzellen bleibt als Zahl erhalten). */
function cellToGtin(v: unknown): string {
  if (typeof v === 'number') return Number.isFinite(v) ? Math.round(v).toString() : '';
  const s = String(v ?? '').trim().replace(/^'/, '').replace(/[\s()]/g, '');
  if (/^\d{8,14}$/.test(s)) return s;
  // Manchmal steht der ganze UDI-Code drin: (01) + 14-stellige GTIN + weitere Angaben
  if (/^01\d{14}/.test(s)) return s.slice(2, 16);
  return '';
}

/** Produktname ohne Größe und Seite (z. B. „Einmalhandschuhe Größe S“ → „Einmalhandschuhe“). */
export function productGroup(name1: string, size: string): string {
  let g = ' ' + name1.replace(/\s+/g, ' ') + ' ';
  g = g.replace(/\b(gr(ö|oe)(ß|ss)e|gr\.)\s*[\w/.,-]+/gi, ' ');
  if (size) g = g.replace(new RegExp(`(^|[\\s,(])${size.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?=$|[\\s,)])`, 'gi'), ' ');
  g = g.replace(/[,\s]+\b(links|rechts|li\.?|re\.?)\b/gi, ' ');
  return g.replace(/\s+/g, ' ').replace(/[\s,;-]+$/, '').trim() || name1.trim();
}

export type CatalogInfo = { articles: Map<string, Article>; total: number; withGtin: number; columns: Record<string, string> };

/** Liest eine Artikelliste (Zeilen wie aus Excel). Erkennt die Spalten an ihren Überschriften. */
export function parseCatalog(rows: unknown[][]): CatalogInfo {
  const find = (h: string[], res: RegExp[], not?: RegExp) => {
    for (const re of res) {
      const i = h.findIndex((x) => re.test(x) && !(not && not.test(x)));
      if (i >= 0) return i;
    }
    return -1;
  };
  let headerRow = -1;
  let col = { gtin: -1, artNr: -1, herstArtNr: -1, name: -1, name2: -1, size: -1, hersteller: -1, produktart: -1 };
  for (let r = 0; r < Math.min(rows.length, 30); r++) {
    const h = (rows[r] ?? []).map(norm);
    const gtin = find(h, [/^(ean|gtin|ean[\s/-]*gtin|gtin[\s/-]*ean)(\s*-?\s*(nr\.?|nummer|code))?$/, /\b(ean|gtin|barcode|strichcode)\b/]);
    if (gtin < 0) continue;
    headerRow = r;
    col = {
      gtin,
      herstArtNr: find(h, [/hersteller.*(art|nr|nummer|bestell)/, /lieferanten.*(art|nr|nummer|bestell)/]),
      artNr: find(h, [/^(artikel|art)[\s.-]*(nr|nummer|no)\.?$/, /artikel.?n(r|ummer)/, /^art\.?\s*nr/], /hersteller|lieferant/),
      name: find(h, [/^(artikel)?bezeichnung ?1?$/, /^name ?1$/, /bezeichnung|artikelname|beschreibung|artikeltext|^name$/]),
      name2: find(h, [/^(artikel)?bezeichnung ?2$/, /^name ?2$/, /^zusatz(text)?$/]),
      size: find(h, [/^(gr(ö|oe)(ß|ss)e|gr\.?|size)$/]),
      hersteller: find(h, [/^hersteller(name)?$/, /^marke$/, /hersteller(?!.*(nr|nummer))/, /^lieferant(enname)?$/]),
      produktart: find(h, [/^(produktart|warengruppe|artikelgruppe|kategorie)$/, /warengruppe|produktart|artikelgruppe|kategorie|gruppe/]),
    };
    break;
  }
  const articles = new Map<string, Article>();
  let total = 0;
  if (headerRow < 0) return { articles, total, withGtin: 0, columns: {} };
  const get = (row: unknown[], i: number) => (i >= 0 ? String(row[i] ?? '').trim() : '');
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] ?? [];
    if (!row.some((x) => String(x ?? '').trim())) continue;
    total++;
    const gtin = cellToGtin(row[col.gtin]);
    if (!gtin) continue;
    const key = normId(gtin);
    if (articles.has(key)) continue;
    const name1 = get(row, col.name);
    const size = get(row, col.size);
    const name = [name1, get(row, col.name2)].filter(Boolean).join(' ').replace(/\s+/g, ' ');
    const sizeText = size && !new RegExp(`\\b${size.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(name) ? ` Gr. ${size}` : '';
    const herstNr = get(row, col.herstArtNr).replace(/^[-.\s]+$/, '');
    articles.set(key, {
      gtin,
      artNr: herstNr || get(row, col.artNr),
      name: name + sizeText,
      hersteller: get(row, col.hersteller),
      // Ohne Warengruppe: das Produkt (Name1) als Gruppe, darunter die Größen/Varianten
      produktart: get(row, col.produktart) || productGroup(name1, size),
    });
  }
  const header = rows[headerRow] ?? [];
  const columns: Record<string, string> = {};
  for (const [k, i] of Object.entries(col)) if (i >= 0) columns[k] = String(header[i]);
  return { articles, total, withGtin: articles.size, columns };
}

/** Eingebaute Sporlastic-Artikel (aus der Konsilager-Liste) als Rückfall, falls nicht im Artikelstamm. */
export function builtinArticles(): Map<string, Article> {
  const byKey = new Map(STANDARD_LISTE.map((t) => [t.key, t]));
  const m = new Map<string, Article>();
  for (const [gtin, key] of Object.entries(GTIN_ZUORDNUNG)) {
    const t = byKey.get(key);
    const [artNr, side, size] = key.split('|');
    m.set(normId(gtin), {
      gtin,
      artNr,
      name: t ? `${t.name} ${t.side ?? side.toLowerCase()} Gr. ${t.size ?? size}` : key,
      hersteller: 'Sporlastic',
      produktart: 'Orthesen',
    });
  }
  return m;
}

// ---------- Zusammenfassung ----------

export type SummaryItem = { artNr: string; name: string; gtin: string; qty: number; batches: string[]; expiry: string };
export type SummaryGroup = { hersteller: string; produktart: string; items: SummaryItem[]; qty: number };

const OHNE = (s: string, fallback: string) => s.trim() || fallback;
const cmp = (a: string, b: string) => a.localeCompare(b, 'de', { numeric: true, sensitivity: 'base' });

/** Gruppiert alle Scans nach Hersteller → Produktart → Artikel (gleiche GTIN bzw. Artikel-Nr.). */
export function summarize(scans: LagerScan[]): SummaryGroup[] {
  const groups = new Map<string, SummaryGroup>();
  for (const s of scans) {
    const h = OHNE(s.hersteller, 'Ohne Hersteller');
    const p = OHNE(s.produktart, 'Ohne Produktart');
    const gk = `${h.toLowerCase()}|${p.toLowerCase()}`;
    let g = groups.get(gk);
    if (!g) groups.set(gk, (g = { hersteller: h, produktart: p, items: [], qty: 0 }));
    const ik = normId(s.gtin) || s.artNr.toLowerCase() || s.name.toLowerCase();
    let it = g.items.find((x) => (normId(x.gtin) || x.artNr.toLowerCase() || x.name.toLowerCase()) === ik);
    if (!it) g.items.push((it = { artNr: s.artNr, name: s.name, gtin: s.gtin, qty: 0, batches: [], expiry: '' }));
    it.qty += s.qty;
    g.qty += s.qty;
    if (s.batch) for (let i = 0; i < Math.max(1, Math.round(s.qty)); i++) it.batches.push(s.batch);
    if (s.expiry && (!it.expiry || s.expiry < it.expiry)) it.expiry = s.expiry;
  }
  const list = [...groups.values()];
  for (const g of list) g.items.sort((a, b) => cmp(a.name || a.artNr, b.name || b.artNr));
  return list.sort((a, b) => cmp(a.hersteller, b.hersteller) || cmp(a.produktart, b.produktart));
}

export const SUMMARY_HEADERS = ['Hersteller', 'Produktart', 'Artikel-Nr.', 'Bezeichnung', 'GTIN', 'Menge', 'Chargen', 'Frühester Verfall'];

/** Werte für das Blatt „Übersicht“: Kopf, dann je Hersteller/Produktart eine Überschrift und die Artikel. */
export function summaryValues(groups: SummaryGroup[], title: string, stand: string): (string | number)[][] {
  const w = SUMMARY_HEADERS.length;
  const pad = (r: (string | number)[]) => [...r, ...Array(w - r.length).fill('')];
  const out: (string | number)[][] = [pad([title]), pad([`Stand: ${stand}`]), pad([]), SUMMARY_HEADERS];
  let total = 0;
  for (const g of groups) {
    out.push(pad([`${g.hersteller}  ·  ${g.produktart}`, '', '', '', '', g.qty]));
    for (const it of g.items) {
      out.push([g.hersteller, g.produktart, it.artNr, it.name, it.gtin ? "'" + it.gtin : '', it.qty, it.batches.join(', '), it.expiry]);
    }
    total += g.qty;
  }
  out.push(pad([]));
  out.push(pad(['Gesamt', '', '', '', '', total]));
  return out;
}

/** Spaltenbuchstabe (1 = A). */
export function colLetter(n: number): string {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
