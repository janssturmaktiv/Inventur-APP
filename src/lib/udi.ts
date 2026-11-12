// Auswertung von UDI-Codes auf Medizinprodukten.
// Unterstützt: GS1 (DataMatrix/QR/128, auch Klartext mit Klammern), GS1 Digital Link (URL) und HIBC.
import { parseGs1 } from './gs1';

export type Udi = {
  issuer: 'GS1' | 'HIBC' | 'unbekannt';
  udiDi?: string; // GS1: GTIN, HIBC: LIC+Artikelcode+Einheit
  productCode?: string; // HIBC: Artikelcode des Herstellers (PCN)
  batch?: string;
  expiry?: string; // JJJJ-MM-TT
  serial?: string;
};

function digitalLink(input: string): Udi | null {
  if (!/^https?:\/\//i.test(input)) return null;
  let url: URL;
  try {
    url = new URL(input);
  } catch {
    return null;
  }
  const parts = url.pathname.split('/').filter(Boolean).map(decodeURIComponent);
  const ais: Record<string, string> = {};
  const start = parts.findIndex((p) => p === '01' || p === 'gtin');
  if (start === -1) return null;
  for (let i = start; i + 1 < parts.length; i += 2) ais[parts[i] === 'gtin' ? '01' : parts[i]] = parts[i + 1];
  url.searchParams.forEach((v, k) => (ais[k] = v));
  const gs1 = parseGs1(Object.entries(ais).map(([k, v]) => `(${k})${v}`).join(''));
  if (!gs1) return null;
  return { issuer: 'GS1', udiDi: gs1.gtin, batch: gs1.batch, expiry: gs1.expiry, serial: gs1.serial };
}

function hibcDate(fmt: string, s: string): { expiry?: string; rest: string } {
  const y = (yy: string) => 2000 + Number(yy);
  const julian = (yy: string, jjj: string) => {
    const d = new Date(Date.UTC(y(yy), 0, Number(jjj)));
    return d.toISOString().slice(0, 10);
  };
  switch (fmt) {
    case '2': return { expiry: `${y(s.slice(4, 6))}-${s.slice(0, 2)}-${s.slice(2, 4)}`, rest: s.slice(6) }; // MMTTJJ
    case '3': return { expiry: `${y(s.slice(0, 2))}-${s.slice(2, 4)}-${s.slice(4, 6)}`, rest: s.slice(6) }; // JJMMTT
    case '4': return { expiry: `${y(s.slice(0, 2))}-${s.slice(2, 4)}-${s.slice(4, 6)}`, rest: s.slice(8) }; // JJMMTTHH
    case '5': return { expiry: julian(s.slice(0, 2), s.slice(2, 5)), rest: s.slice(5) }; // JJTTT
    case '6': return { expiry: julian(s.slice(0, 2), s.slice(2, 5)), rest: s.slice(7) }; // JJTTTHH
    case '7': return { rest: s }; // kein Datum
    default: return { rest: s };
  }
}

function lastDay(year: number, month: number) {
  return String(new Date(year, month, 0).getDate()).padStart(2, '0');
}

// Sekundärdaten ohne Prüfzeichen
function hibcSecondary(sec: string): Pick<Udi, 'batch' | 'expiry' | 'serial'> {
  if (sec.startsWith('$$+')) {
    const r = sec.slice(3);
    if (/^[2-7]/.test(r)) { const d = hibcDate(r[0], r.slice(1)); return { expiry: d.expiry, serial: d.rest || undefined }; }
    const yy = 2000 + Number(r.slice(2, 4));
    return { expiry: `${yy}-${r.slice(0, 2)}-${lastDay(yy, Number(r.slice(0, 2)))}`, serial: r.slice(4) || undefined };
  }
  if (sec.startsWith('$+')) return { serial: sec.slice(2) || undefined };
  if (sec.startsWith('$$')) {
    let r = sec.slice(2);
    if (r[0] === '8') r = r.slice(3); // Menge 2-stellig
    else if (r[0] === '9') r = r.slice(6); // Menge 5-stellig
    if (/^[2-7]/.test(r)) { const d = hibcDate(r[0], r.slice(1)); return { expiry: d.expiry, batch: d.rest || undefined }; }
    const yy = 2000 + Number(r.slice(2, 4)); // MMJJ
    return { expiry: `${yy}-${r.slice(0, 2)}-${lastDay(yy, Number(r.slice(0, 2)))}`, batch: r.slice(4) || undefined };
  }
  if (sec.startsWith('$')) return { batch: sec.slice(1) || undefined };
  if (/^\d{5}/.test(sec)) {
    // JJTTT + Charge
    const d = hibcDate('5', sec);
    return { expiry: d.expiry, batch: d.rest || undefined };
  }
  return {};
}

function hibc(input: string): Udi | null {
  let s = input.trim();
  if (s.startsWith('*') && s.endsWith('*')) s = s.slice(1, -1); // Code-39-Sternchen
  if (!s.startsWith('+') || s.length < 6) return null;
  s = s.slice(1);
  if (/^[A-Z]\d{3}/i.test(s) && !s.startsWith('$')) {
    // Primärdaten (+ ggf. /Sekundärdaten)
    const slash = s.indexOf('/');
    let primary: string;
    let secondary: string | undefined;
    if (slash === -1) {
      primary = s.slice(0, -1); // Prüfzeichen ab
    } else {
      primary = s.slice(0, slash);
      secondary = s.slice(slash + 1, -1); // Prüfzeichen am Ende ab
    }
    const lic = primary.slice(0, 4);
    const pcnU = primary.slice(4);
    const udi: Udi = {
      issuer: 'HIBC',
      udiDi: (lic + pcnU).toUpperCase(),
      productCode: pcnU.slice(0, -1).toUpperCase(),
    };
    return secondary ? { ...udi, ...hibcSecondary(secondary) } : udi;
  }
  // Nur Sekundärcode: +<Verknüpfungszeichen-frei> ... Format: Daten + Link + Prüfzeichen
  const data = s.slice(0, -2);
  return { issuer: 'HIBC', ...hibcSecondary(data) };
}

export function parseUdi(raw: string): Udi {
  const input = raw.trim();
  const dl = digitalLink(input);
  if (dl) return dl;
  const h = hibc(input);
  if (h) return h;
  const gs1 = parseGs1(input);
  if (gs1) return { issuer: 'GS1', udiDi: gs1.gtin, batch: gs1.batch, expiry: gs1.expiry, serial: gs1.serial };
  return { issuer: 'unbekannt' };
}

/** Vergleichbare Form einer Produktkennung (GTIN-13/14 mit führenden Nullen gleich behandeln). */
export function normId(v?: string): string {
  if (!v) return '';
  const s = String(v).trim().toUpperCase().replace(/\s/g, '');
  return /^\d+$/.test(s) ? s.replace(/^0+/, '') : s;
}
