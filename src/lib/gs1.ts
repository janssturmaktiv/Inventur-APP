// Minimaler GS1-Parser für DataMatrix / GS1-128 / GS1-QR.
// Erkennt u. a. (01) GTIN, (10) Charge, (17) Verfallsdatum, (21) Seriennummer.

export type Gs1Result = {
  gtin?: string;
  batch?: string;
  expiry?: string; // JJJJ-MM-TT
  serial?: string;
  all: Record<string, string>;
};

const GS = '\u001d';

// AI -> feste Länge der Daten (ohne AI). null = variabel (bis GS oder Ende).
const FIXED: Record<string, number> = {
  '00': 18, '01': 14, '02': 14,
  '11': 6, '12': 6, '13': 6, '15': 6, '16': 6, '17': 6,
  '20': 2,
};
const VARIABLE_2 = ['10', '21', '22', '30', '37', '90', '91', '92', '93', '94', '95', '96', '97', '98', '99'];
const VARIABLE_3 = ['240', '241', '242', '250', '251', '253', '254', '400', '401', '402', '403', '410', '411', '412', '413', '414', '420', '421', '422', '423', '424', '425', '426'];

function formatDate(yymmdd: string): string {
  const yy = Number(yymmdd.slice(0, 2));
  const mm = yymmdd.slice(2, 4);
  let dd = yymmdd.slice(4, 6);
  // GS1: Jahre 00–49 => 20xx, 51–99 => 19xx (vereinfachte Regel)
  const year = yy <= 49 ? 2000 + yy : 1900 + yy;
  if (dd === '00') {
    // "00" heißt: letzter Tag des Monats
    dd = String(new Date(year, Number(mm), 0).getDate()).padStart(2, '0');
  }
  return `${year}-${mm}-${dd}`;
}

function parseBracketed(input: string): Record<string, string> | null {
  // Klartext-Form "(01)04012345678901(10)ABC123"
  const re = /\((\d{2,4})\)([^(]*)/g;
  const out: Record<string, string> = {};
  let m: RegExpExecArray | null;
  let found = false;
  while ((m = re.exec(input))) {
    out[m[1]] = m[2].trim();
    found = true;
  }
  return found ? out : null;
}

function parseRaw(input: string): Record<string, string> | null {
  let s = input;
  // Symbologie-Kennung entfernen (]d2 DataMatrix, ]C1 GS1-128, ]Q3 QR, ]e0 DataBar)
  s = s.replace(/^\][A-Za-z]\d/, '');
  if (s.startsWith(GS)) s = s.slice(1);
  if (!/^\d{2}/.test(s)) return null;

  const out: Record<string, string> = {};
  let i = 0;
  while (i < s.length) {
    if (s[i] === GS) { i++; continue; }
    const two = s.slice(i, i + 2);
    const three = s.slice(i, i + 3);
    const four = s.slice(i, i + 4);
    let ai: string;
    let len: number | null;
    if (FIXED[two] !== undefined) { ai = two; len = FIXED[two]; }
    else if (/^3[1-6]\d\d$/.test(four)) { ai = four; len = 6; } // Maße/Gewichte
    else if (/^41[0-7]$/.test(three)) { ai = three; len = 13; } // GLNs
    else if (VARIABLE_3.includes(three)) { ai = three; len = null; }
    else if (VARIABLE_2.includes(two)) { ai = two; len = null; }
    else return Object.keys(out).length ? out : null; // unbekannt -> abbrechen
    i += ai.length;
    let value: string;
    if (len !== null) {
      value = s.slice(i, i + len);
      i += len;
    } else {
      const end = s.indexOf(GS, i);
      value = end === -1 ? s.slice(i) : s.slice(i, end);
      i = end === -1 ? s.length : end + 1;
    }
    out[ai] = value;
  }
  return Object.keys(out).length ? out : null;
}

export function parseGs1(input: string): Gs1Result | null {
  const all = input.includes('(') ? parseBracketed(input) : parseRaw(input);
  if (!all) return null;
  // Nur als GS1 werten, wenn typische AIs vorkommen
  if (!all['01'] && !all['10'] && !all['21'] && !all['17']) return null;
  return {
    gtin: all['01'],
    batch: all['10'],
    expiry: all['17'] && /^\d{6}$/.test(all['17']) ? formatDate(all['17']) : undefined,
    serial: all['21'],
    all,
  };
}
