/**
 * Office Script „Inventur-Scans eintragen“
 *
 * Wird von Power Automate mit dem Text einer E-Mail aus der Konsilager-App aufgerufen.
 * Erhöht in der Bestandsliste die Spalte „Ist“ und ergänzt „Chargen-Nr.“ in der passenden Zeile
 * (Artikel-Nr. + Seite + Größe). Optional: Inventur abschließen (leere Ist-Felder auf 0,
 * Inventurdatum und „Geprüft von“ eintragen).
 */
function main(workbook: ExcelScript.Workbook, mailBody: string): string {
  // --- E-Mail-Text auswerten (HTML entfernen, Daten zwischen den Markierungen) ---
  let text = String(mailBody || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
  const start = text.indexOf('###INVENTUR-V1###');
  const end = text.indexOf('###ENDE###');
  if (start < 0 || end < 0 || end < start) return 'Keine Inventur-Daten in der E-Mail gefunden.';
  const lines = text
    .slice(start + 17, end)
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  // --- Bestandsliste finden: Blatt mit Kopfzeile „Artikel-Nr.“ und „Ist“ ---
  const norm = (s: string) => String(s || '').trim().toLowerCase();
  let sheet: ExcelScript.Worksheet | undefined;
  let texts: string[][] = [];
  let headerRow = -1;
  let rowOffset = 0;
  let colOffset = 0;
  const col: { [k: string]: number } = {};
  for (const ws of workbook.getWorksheets()) {
    const used = ws.getUsedRange();
    if (!used) continue;
    const t = used.getTexts();
    for (let r = 0; r < Math.min(t.length, 40); r++) {
      const h = t[r].map(norm);
      const art = h.findIndex((x) => /^art(ikel)?[\s.\-_]*(nr|nummer|no)\.?$/.test(x));
      const ist = h.findIndex((x) => /^(ist|ist[\s-]*(menge|bestand)|gezählt)$/.test(x));
      if (art >= 0 && ist >= 0) {
        sheet = ws;
        texts = t;
        headerRow = r;
        rowOffset = used.getRowIndex();
        colOffset = used.getColumnIndex();
        col.art = art;
        col.ist = ist;
        col.side = h.findIndex((x) => /^(seite|links\/rechts|l\/r)$/.test(x));
        col.size = h.findIndex((x) => /^(gr(ö|oe)(ß|ss)e|gr\.|size)$/.test(x));
        col.batch = h.findIndex((x) => /^charge(n)?[\s.\-]*(nr|nummer)?\.?$|^lot$/.test(x));
        col.gtin = h.findIndex((x) => /(gtin|udi|ean)/.test(x));
        break;
      }
    }
    if (sheet) break;
  }
  if (!sheet) return 'Keine Bestandsliste mit den Spalten „Artikel-Nr.“ und „Ist“ gefunden.';
  const ws = sheet;

  // Zeilen-Index: "ARTIKEL|SEITE|GRÖSSE" -> Zeilennummer (0-basiert im Blatt)
  const cell = (r: number, c: number) => (c >= 0 ? String(texts[r][c] ?? '').trim() : '');
  const rows: { [key: string]: number } = {};
  const itemRows: number[] = [];
  for (let r = headerRow + 1; r < texts.length; r++) {
    const art = cell(r, col.art);
    if (!art || art.indexOf('·') >= 0) continue; // Gruppen-Überschriften überspringen
    const key = [art, cell(r, col.side), cell(r, col.size)].map((x) => x.toUpperCase()).join('|');
    if (rows[key] === undefined) rows[key] = rowOffset + r;
    itemRows.push(rowOffset + r);
  }

  const istCol = colOffset + col.ist;
  const batchCol = col.batch >= 0 ? colOffset + col.batch : -1;
  const gtinCol = col.gtin >= 0 ? colOffset + col.gtin : -1;
  let learned = 0;
  let added = 0;
  const missing: string[] = [];
  let finishedBy = '';

  for (const line of lines) {
    const p = line.split(';');
    if (p[0] === 'A' && p.length >= 3) {
      const key = p[1].toUpperCase();
      const qty = Number(String(p[2]).replace(',', '.')) || 1;
      const batch = (p[3] || '').trim();
      const r = rows[key];
      if (r === undefined) {
        missing.push(p[1]);
        continue;
      }
      const istCell = ws.getCell(r, istCol);
      istCell.setValue((Number(istCell.getValue()) || 0) + qty);
      if (batchCol >= 0 && batch) {
        const bCell = ws.getCell(r, batchCol);
        const list = String(bCell.getValue() || '')
          .split(/[,;\n]/)
          .map((x) => x.trim())
          .filter((x) => x.length > 0);
        if (list.indexOf(batch) < 0) {
          list.push(batch);
          bCell.setValue(list.join(', '));
        }
      }
      // GTIN in der Liste vermerken (falls Spalte vorhanden und noch leer): dann erkennen alle Handys den Code
      const gtin = (p[4] || '').trim();
      if (gtinCol >= 0 && gtin) {
        const gCell = ws.getCell(r, gtinCol);
        if (String(gCell.getValue()).trim() === '') {
          gCell.setValue("'" + gtin);
          learned++;
        }
      }
      added += qty;
    } else if (p[0] === 'F') {
      finishedBy = (p[1] || '').trim() || '–';
    }
  }

  if (finishedBy) {
    for (const r of itemRows) {
      const c = ws.getCell(r, istCol);
      if (String(c.getValue()) === '') c.setValue(0);
    }
    // Inventurdatum und „Geprüft von“ oberhalb der Kopfzeile: Wert steht rechts neben der Beschriftung
    const now = new Date();
    // Excel-Seriennummer, damit die Zelle ein echtes Datum bleibt
    const date = Math.round((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
    for (let r = 0; r < headerRow; r++) {
      for (let c = 0; c < texts[r].length; c++) {
        const label = norm(texts[r][c]);
        const setRight = (value: string | number) => {
          let target = c + 1;
          for (let k = c + 1; k < Math.min(texts[r].length, c + 4); k++) {
            if (String(texts[r][k]).trim()) {
              target = k;
              break;
            }
          }
          ws.getCell(rowOffset + r, colOffset + target).setValue(value);
        };
        if (/^inventurdatum/.test(label)) setRight(date);
        if (/^gepr(ü|ue)ft von/.test(label) && finishedBy !== '–') setRight(finishedBy);
      }
    }
  }

  let msg = `${added} Stück eingetragen.`;
  if (learned) msg += ` ${learned} GTIN(s) in der Liste ergänzt.`;
  if (missing.length) msg += ` Nicht gefunden: ${missing.join(', ')}.`;
  if (finishedBy) msg += ' Inventur abgeschlossen.';
  return msg;
}
