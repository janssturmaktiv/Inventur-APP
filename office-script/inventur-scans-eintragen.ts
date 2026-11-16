/**
 * Office Script „Inventur-Scans eintragen“
 *
 * Wird von Power Automate mit dem Text einer E-Mail aus der Konsilager-App aufgerufen.
 * Erhöht in der Bestandsliste die Spalte „Ist“ und ergänzt „Chargen-Nr.“ in der passenden Zeile
 * (Artikel-Nr. + Seite + Größe). Bei einer neuen Inventur (anderes Datum als in der Liste) werden
 * vorher Ist und Chargen-Nr. geleert. Inventurdatum und „Geprüft von“ werden immer eingetragen.
 * Optional: Inventur abschließen (leere Ist-Felder auf 0).
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

  // --- Felder „Inventurdatum“ und „Geprüft von“ oberhalb der Kopfzeile finden ---
  // Wertzelle = direkt rechts neben der Beschriftung. Ist die leer und nicht verbunden, wird die
  // nächste verbundene Zelle (z. B. G4:H4) bzw. die nächste Zelle mit festem Wert genommen.
  // Zellen mit Formeln werden nie überschrieben.
  const valueCellRight = (r: number, c: number): ExcelScript.Range | undefined => {
    let fallback: ExcelScript.Range | undefined;
    for (let k = c + 1; k <= c + 3; k++) {
      const x = ws.getCell(rowOffset + r, colOffset + k);
      const isFormula = String(x.getFormula() || '').startsWith('=');
      if (isFormula) continue;
      // Verbundene Zellen erkennen (falls die Excel-Version das nicht kann: ohne diese Info weiter)
      let isMergeStart = false;
      let isMergeInner = false;
      try {
        const merged = x.getMergedAreas();
        const areas = merged ? merged.getAreas() : [];
        if (areas.length > 0) {
          isMergeStart = areas[0].getColumnIndex() === colOffset + k;
          isMergeInner = !isMergeStart;
        }
      } catch (e) {
        // ignorieren
      }
      if (isMergeInner) continue;
      if (k === c + 1 && (isMergeStart || String(texts[r][k] ?? '').trim() !== '')) return x;
      if (k === c + 1) {
        fallback = x;
        continue;
      }
      if (isMergeStart || String(texts[r][k] ?? '').trim() !== '') return x;
    }
    return fallback;
  };
  let dateCell: ExcelScript.Range | undefined;
  let checkerCell: ExcelScript.Range | undefined;
  for (let r = 0; r < headerRow; r++) {
    for (let c = 0; c < texts[r].length; c++) {
      const label = norm(texts[r][c]);
      if (!dateCell && /^inventurdatum/.test(label)) dateCell = valueCellRight(r, c);
      if (!checkerCell && /^gepr(ü|ue)ft von/.test(label)) checkerCell = valueCellRight(r, c);
    }
  }
  const serial = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const now = new Date();
  const today = serial(now.getFullYear(), now.getMonth() + 1, now.getDate());

  // --- Kopfdaten der E-Mail ---
  let sender = '';
  let inventoryDate = 0; // Excel-Seriennummer des Inventurtags laut App
  let forceReset = false;
  for (const line of lines) {
    const p = line.split(';');
    if (p[0] === 'S') sender = (p[1] || '').trim();
    if (p[0] === 'N') forceReset = true;
    if (p[0] === 'D') {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((p[1] || '').trim());
      if (m) inventoryDate = serial(Number(m[1]), Number(m[2]), Number(m[3]));
    }
  }

  // --- Neue Inventur? Dann zuerst alte Ist-Werte und Chargen leeren ---
  // Neu ist die Inventur, wenn das Inventurdatum der App nicht dem Datum in der Liste entspricht
  // (oder die App ausdrücklich „neu befüllen“ verlangt). So werden alte Zahlen nie dazugezählt.
  const listDate = dateCell ? Math.floor(Number(dateCell.getValue()) || 0) : 0;
  if (!forceReset && inventoryDate && listDate && inventoryDate < listDate) {
    // Scans einer älteren Inventur (z. B. vergessen auf einem Handy) nicht in die aktuelle mischen
    return `Nicht eingetragen: Die Scans gehören zu einer älteren Inventur als die in der Liste.`;
  }
  let reset = false;
  if (forceReset || (inventoryDate && listDate !== inventoryDate)) {
    for (const r of itemRows) {
      ws.getCell(r, istCol).setValue('');
      if (batchCol >= 0) ws.getCell(r, batchCol).setValue('');
    }
    if (checkerCell) checkerCell.setValue('');
    reset = true;
  }
  if (dateCell && (reset || !listDate)) dateCell.setValue(inventoryDate || today);
  // Prüfer: alle Namen, die zu dieser Inventur gescannt haben
  const addChecker = (name: string) => {
    if (!checkerCell || !name || name === 'unbekannt' || name === '–') return;
    const names = String(checkerCell.getValue() || '')
      .split(',')
      .map((x) => x.trim())
      .filter((x) => x.length > 0);
    let changed = false;
    for (const part of name.split(',').map((x) => x.trim()).filter((x) => x.length > 0)) {
      if (names.indexOf(part) < 0) {
        names.push(part);
        changed = true;
      }
    }
    if (changed) checkerCell.setValue(names.join(', '));
  };
  addChecker(sender);

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
    // Nicht gefundene Artikel auf 0 setzen (→ „nachbestellen“)
    for (const r of itemRows) {
      const c = ws.getCell(r, istCol);
      if (String(c.getValue()) === '') c.setValue(0);
    }
    if (dateCell && !Number(dateCell.getValue())) dateCell.setValue(inventoryDate || today);
    addChecker(finishedBy);
  }

  let msg = `${added} Stück eingetragen.`;
  if (reset) msg = 'Neue Inventur: alte Ist-Werte und Chargen geleert. ' + msg;
  if (learned) msg += ` ${learned} GTIN(s) in der Liste ergänzt.`;
  if (missing.length) msg += ` Nicht gefunden: ${missing.join(', ')}.`;
  if (finishedBy) msg += ' Inventur abgeschlossen.';
  return msg;
}
