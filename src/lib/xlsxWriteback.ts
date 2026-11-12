// Schreibt Werte direkt in die Original-Excel-Datei (XML-Ebene), damit Formatierung,
// Formeln, Farben und verbundene Zellen unverändert bleiben.
import JSZip from 'jszip';

export type CellUpdate = { ref: string; value: string | number | null };

const xmlEsc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

function colToNum(col: string) {
  return col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}
function splitRef(ref: string) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref.toUpperCase());
  if (!m) throw new Error(`Ungültige Zelle ${ref}`);
  return { col: m[1], row: Number(m[2]) };
}

function cellXml(ref: string, style: string | undefined, value: string | number | null) {
  const s = style ? ` s="${style}"` : '';
  if (value === null || value === '') return `<c r="${ref}"${s}/>`;
  if (typeof value === 'number') return `<c r="${ref}"${s}><v>${value}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(value)}</t></is></c>`;
}

/** Eine Zelle im Blatt-XML setzen (vorhandene Formatierung der Zelle bleibt). */
export function setCellInSheetXml(xml: string, ref: string, value: string | number | null): string {
  const { col, row } = splitRef(ref);
  // vorhandene Zelle? (<c r="F10" .../> oder <c r="F10" ...>...</c>)
  const cellRe = new RegExp(`<c r="${ref}"(\\s[^>]*?)?(/>|>[\\s\\S]*?</c>)`);
  const m = cellRe.exec(xml);
  if (m) {
    const style = /\ss="(\d+)"/.exec(m[1] ?? '')?.[1];
    return xml.slice(0, m.index) + cellXml(ref, style, value) + xml.slice(m.index + m[0].length);
  }
  if (value === null) return xml;
  // Zeile suchen
  const rowRe = new RegExp(`<row r="${row}"(\\s[^>]*?)?(/>|>([\\s\\S]*?)</row>)`);
  const rm = rowRe.exec(xml);
  if (rm) {
    const open = `<row r="${row}"${rm[1] ?? ''}>`;
    const inner = rm[3] ?? '';
    // Zelle an der richtigen Spaltenposition einfügen
    const cells = [...inner.matchAll(/<c r="([A-Z]+)\d+"/g)];
    const target = colToNum(col);
    const after = cells.find((c) => colToNum(c[1]) > target);
    const newCell = cellXml(ref, undefined, value);
    const newInner = after ? inner.slice(0, after.index) + newCell + inner.slice(after.index!) : inner + newCell;
    return xml.slice(0, rm.index) + open + newInner + '</row>' + xml.slice(rm.index + rm[0].length);
  }
  // Zeile fehlt: an der richtigen Stelle in sheetData einfügen
  const rows = [...xml.matchAll(/<row r="(\d+)"/g)];
  const next = rows.find((r) => Number(r[1]) > row);
  const newRow = `<row r="${row}">${cellXml(ref, undefined, value)}</row>`;
  if (next) return xml.slice(0, next.index) + newRow + xml.slice(next.index!);
  return xml.replace('</sheetData>', newRow + '</sheetData>').replace('<sheetData/>', `<sheetData>${newRow}</sheetData>`);
}

async function sheetPath(zip: JSZip, sheetName: string): Promise<string> {
  const wbXml = await zip.file('xl/workbook.xml')!.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  const nameEsc = xmlEsc(sheetName).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const sm = new RegExp(`<sheet\\b[^>]*\\bname="${nameEsc}"[^>]*>`).exec(wbXml);
  if (!sm) throw new Error(`Blatt „${sheetName}“ nicht gefunden.`);
  const rid = /\br:id="([^"]+)"|\bid="([^"]+)"/.exec(sm[0].replace(/sheetId="[^"]*"/, ''));
  const id = rid?.[1] ?? rid?.[2];
  const rel = new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*>`).exec(relsXml)?.[0];
  const target = rel && /Target="([^"]+)"/.exec(rel)?.[1];
  if (!target) throw new Error('Blattverweis nicht gefunden.');
  return target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
}

/** Wendet Zellwerte auf die Excel-Datei an und liefert die neue Datei als Base64. */
export async function writeCells(base64: string, sheetName: string, updates: CellUpdate[]): Promise<string> {
  const zip = await JSZip.loadAsync(base64, { base64: true });
  const path = await sheetPath(zip, sheetName);
  let xml = await zip.file(path)!.async('string');
  for (const u of updates) xml = setCellInSheetXml(xml, u.ref, u.value);
  zip.file(path, xml);

  // Excel soll Formeln (Differenz, Status) beim Öffnen neu berechnen
  let wbXml = await zip.file('xl/workbook.xml')!.async('string');
  if (/<calcPr\b/.test(wbXml)) {
    wbXml = wbXml.replace(/<calcPr\b([^>]*?)\/?>/, (all, attrs: string) => {
      const a = attrs.replace(/\s?fullCalcOnLoad="[^"]*"/, '');
      return `<calcPr${a} fullCalcOnLoad="1"/>`;
    });
  } else {
    wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  }
  zip.file('xl/workbook.xml', wbXml);
  // Zwischengespeicherte Berechnungskette entfernen (wird neu aufgebaut)
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', ct.replace(/<Override[^>]*calcChain[^>]*\/>/, ''));
    const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
    zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship[^>]*calcChain[^>]*\/>/, ''));
  }
  return zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
}
