import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { AccountInfo } from '@azure/msal-browser';
import { MS_CONFIG } from './config';
import { getFileInfo, initAuth, listDrives, listFolder, login, logout, me, resolveFile, siteFromUrl, type BrowseEntry, type FileRef } from './lib/msGraph';
import { addToRow, finishInventory, loadShared, removeFromRow, resetInventory, saveMapping } from './lib/sharepointSync';
import { Scanner } from './Scanner';
import { parseUdi, normId } from './lib/udi';
import { parseTargetWorkbook } from './lib/targetList';
import { reconcile, findTarget, type ReconRow, type Reconciliation, type Status } from './lib/reconcile';
import { buildFilledOriginal, buildReport, canShareFile, downloadFile, printReport, shareFile, type ExportFile } from './lib/exportWeb';
import { load, usePersisted } from './lib/storage';
import { itemLabel, type Mappings, type Scan, type TargetItem, type TargetSource } from './lib/types';

const KEYS = {
  scans: 'konsi/scans/v1',
  targets: 'konsi/targets/v2',
  source: 'konsi/source/v1',
  original: 'konsi/original/v1',
  mappings: 'konsi/mappings/v1',
  settings: 'konsi/settings/v1',
  spFile: 'konsi/spfile/v1',
};

type Draft = Omit<Scan, 'id' | 'scannedAt' | 'quantity'> & { quantity: string };
type Settings = { userName: string; listName: string };
type Filter = 'alle' | Status | 'unbekannt';

const emptyDraft = (): Draft => ({ batch: '', udiDi: '', expiry: '', quantity: '1', raw: '', codeType: 'manuell' });
const STATUS_LABEL: Record<Status | 'unbekannt', string> = {
  vollständig: 'vollständig',
  teilweise: 'teilweise',
  fehlend: 'fehlt',
  unbekannt: 'nicht zugeordnet',
};

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(',')[1] ?? '');
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

export default function App() {
  const [scans, setScans] = usePersisted<Scan[]>(KEYS.scans, []);
  const [targets, setTargets] = usePersisted<TargetItem[]>(KEYS.targets, []);
  const [source, setSource] = usePersisted<TargetSource | null>(KEYS.source, null);
  const [original, setOriginal] = usePersisted<string>(KEYS.original, '');
  const [mappings, setMappings] = usePersisted<Mappings>(KEYS.mappings, {});
  const [settings, setSettings] = usePersisted<Settings>(KEYS.settings, { userName: '', listName: '' });
  const [tab, setTab] = useState<'abgleich' | 'scans'>('abgleich');
  const [filter, setFilter] = useState<Filter>('alle');
  const [scanning, setScanning] = useState(false);
  const [draft, setDraft] = useState<Draft | null>(null);
  const [picking, setPicking] = useState(false);
  const [exportMenu, setExportMenu] = useState(false);
  const [ready, setReady] = useState<ExportFile | null>(null);
  const [busy, setBusy] = useState('');
  const [showSettings, setShowSettings] = useState(false);
  const [toast, setToast] = useState('');
  const fileInput = useRef<HTMLInputElement>(null);
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [spFile, setSpFile] = usePersisted<FileRef | null>(KEYS.spFile, null);
  const [lastSync, setLastSync] = useState<Date | null>(null);
  const [spError, setSpError] = useState('');
  const [browsing, setBrowsing] = useState(false);
  const shared = !!(account && spFile);

  const localRecon = useMemo(() => reconcile(targets, scans, mappings), [targets, scans, mappings]);
  // Im SharePoint-Modus zählt der Stand der gemeinsamen Datei (inkl. Scans der Kollegen)
  const recon: Reconciliation = useMemo(() => {
    if (!shared) return localRecon;
    const rows: ReconRow[] = localRecon.rows.map((r) => {
      const ist = r.item.currentIst ?? 0;
      return { ...r, foundQty: ist, status: ist >= r.item.expectedQty ? 'vollständig' : ist > 0 ? 'teilweise' : 'fehlend' };
    });
    const unknown = scans.filter((x) => x.sync === 'unassigned');
    const counts = { vollständig: 0, teilweise: 0, fehlend: 0, unbekannt: unknown.length };
    rows.forEach((r) => counts[r.status]++);
    return { rows, unknown, counts };
  }, [shared, localRecon, scans]);

  // --- Microsoft-Anmeldung und SharePoint ---
  useEffect(() => {
    initAuth()
      .then(async (acc) => {
        setAccount(acc);
        if (acc && !load('konsi/settings/v1', { userName: '' }).userName) {
          const m = await me().catch(() => null);
          if (m?.displayName) setSettings((st) => ({ ...st, userName: m.displayName }));
        }
      })
      .catch((e) => setSpError(`Anmeldung fehlgeschlagen: ${e?.message ?? e}`));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const refreshShared = useCallback(
    async (file: FileRef | null = spFile, quiet = false) => {
      if (!file) return;
      if (!quiet) setBusy('Liste wird aus SharePoint geladen …');
      try {
        const st = await loadShared(file);
        setTargets(st.items);
        setSource(st.source);
        setOriginal('');
        setMappings((m) => ({ ...m, ...st.mappings }));
        setSettings((x) => ({ ...x, listName: file.name.replace(/\.xlsx$/i, '') }));
        setLastSync(new Date());
        setSpError('');
      } catch (e) {
        setSpError(`SharePoint: ${e instanceof Error ? e.message : e}`);
      } finally {
        if (!quiet) setBusy('');
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [spFile],
  );

  // Nach der Anmeldung: Datei über den hinterlegten Link finden
  useEffect(() => {
    if (!account) return;
    if (spFile) {
      refreshShared(spFile);
      return;
    }
    (async () => {
      setBusy('Gemeinsame Liste wird gesucht …');
      try {
        const f = await resolveFile(MS_CONFIG.defaultFileUrl);
        setSpFile(f);
        await refreshShared(f);
      } catch {
        setBrowsing(true);
      } finally {
        setBusy('');
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [account]);

  // Alle 60 Sekunden den Stand der Kollegen holen
  useEffect(() => {
    if (!shared) return;
    const t = setInterval(() => {
      if (document.visibilityState === 'visible' && !scanning && !draft) refreshShared(spFile, true);
    }, 60_000);
    return () => clearInterval(t);
  }, [shared, scanning, draft, spFile, refreshShared]);

  async function pushScan(scan: Scan, target: TargetItem) {
    if (!spFile || !source || !target.row) return;
    try {
      const res = await addToRow(spFile, source, target.row, scan.quantity, scan.batch);
      setScans((p) => p.map((x) => (x.id === scan.id ? { ...x, sync: 'ok', syncedRow: target.row } : x)));
      setTargets((p) => p.map((t) => (t.row === target.row ? { ...t, currentIst: res.ist, currentBatch: res.batch } : t)));
      setLastSync(new Date());
    } catch (e) {
      flash('Nicht gesendet, wird später erneut versucht');
      setSpError(`Eintragen fehlgeschlagen: ${e instanceof Error ? e.message : e}`);
    }
  }

  async function retryPending() {
    const pending = scans.filter((x) => x.sync === 'pending');
    for (const sc of [...pending].reverse()) {
      const t = targets.find((x) => x.row === sc.syncedRow) ?? findTarget(targets, sc, mappings);
      if (t) await pushScan(sc, t);
    }
  }

  const title = `Konsilager-Inventur${settings.listName ? ` – ${settings.listName}` : ''}`;

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 2500);
  };

  const onScanned = useCallback((text: string, format: string) => {
    const udi = parseUdi(text);
    setScanning(false);
    setDraft({
      udiDi: udi.udiDi ?? '',
      productCode: udi.productCode,
      batch: udi.batch ?? (udi.issuer === 'unbekannt' ? text : ''),
      expiry: udi.expiry ?? '',
      serial: udi.serial,
      quantity: '1',
      raw: text,
      codeType: format,
    });
  }, []);

  function saveDraft(): boolean {
    if (!draft) return false;
    if (!draft.batch.trim()) {
      alert('Bitte eine Chargennummer eingeben.');
      return false;
    }
    const scan: Scan = {
      ...draft,
      batch: draft.batch.trim(),
      udiDi: draft.udiDi?.trim() || undefined,
      expiry: draft.expiry?.trim() || undefined,
      quantity: Number(draft.quantity.replace(',', '.')) || 1,
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      scannedAt: new Date().toISOString(),
      scannedBy: settings.userName || undefined,
    };
    // Zuordnung merken: dieselbe UDI-DI wird künftig automatisch erkannt
    const idKey = normId(scan.udiDi) || normId(scan.productCode);
    const newMappings = scan.targetKey && idKey ? { ...mappings, [idKey]: scan.targetKey } : mappings;
    if (newMappings !== mappings) setMappings(newMappings);
    if (shared) {
      const target = findTarget(targets, scan, newMappings);
      if (target?.row) {
        scan.sync = 'pending';
        scan.syncedRow = target.row;
        pushScan(scan, target);
        if (newMappings !== mappings && spFile && (scan.udiDi || scan.productCode)) {
          saveMapping(spFile, (scan.udiDi || scan.productCode)!, target).catch((e) => console.warn('Zuordnung nicht geteilt', e));
        }
      } else scan.sync = 'unassigned';
    }
    setScans((prev) => [scan, ...prev]);
    setDraft(null);
    flash('Gespeichert');
    return true;
  }

  async function onFileChosen(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const b64 = await fileToBase64(file);
      const parsed = parseTargetWorkbook(b64, file.name);
      const isXlsx = /\.xlsx$/i.test(file.name);
      const s = parsed.source;
      const msg =
        `${parsed.items.length} Artikelzeilen aus Blatt „${s.sheetName}“ gelesen.` +
        (isXlsx && (s.istCol || s.batchCol)
          ? `\n\nBeim Export werden eingetragen: ${[s.istCol && `Ist (Spalte ${s.istCol})`, s.batchCol && `Chargen-Nr. (Spalte ${s.batchCol})`, s.dateCell && `Inventurdatum (${s.dateCell})`, s.checkerCell && `Geprüft von (${s.checkerCell})`].filter(Boolean).join(', ')}.`
          : '') +
        (scans.length ? `\n\nDie ${scans.length} bisherigen Scans bleiben erhalten.` : '') +
        '\n\nListe übernehmen?';
      if (!confirm(msg)) return;
      setOriginal(isXlsx ? b64 : '');
      setTargets(parsed.items);
      setSource(isXlsx ? s : { ...s, istCol: undefined, batchCol: undefined });
      setSettings((st) => ({ ...st, listName: file.name.replace(/\.(xlsx|xls|csv)$/i, '') }));
      setTab('abgleich');
      setFilter('alle');
    } catch (err) {
      alert(`Import fehlgeschlagen: ${err instanceof Error ? err.message : err}`);
    }
  }

  async function runExport(kind: 'original' | 'report' | 'pdf') {
    setExportMenu(false);
    try {
      if (kind === 'pdf') return printReport(recon, title, settings.userName);
      setBusy('Datei wird erstellt …');
      const f = kind === 'original' && source && original ? await buildFilledOriginal(original, source, recon, settings.userName) : buildReport(recon, scans);
      setReady(f);
    } catch (err) {
      alert(`Export fehlgeschlagen: ${err instanceof Error ? err.message : err}`);
    } finally {
      setBusy('');
    }
  }

  async function removeScan(id: string) {
    const sc = scans.find((x) => x.id === id);
    if (!sc) return;
    const inShared = shared && sc.sync === 'ok' && sc.syncedRow && source;
    if (!confirm(inShared ? 'Diesen Scan löschen und in der SharePoint-Liste wieder austragen?' : 'Diesen Scan löschen?')) return;
    if (inShared) {
      try {
        const stillUsed = scans.some((x) => x.id !== id && x.sync === 'ok' && x.syncedRow === sc.syncedRow && x.batch === sc.batch);
        const res = await removeFromRow(spFile!, source!, sc.syncedRow!, sc.quantity, sc.batch, stillUsed);
        setTargets((p) => p.map((t) => (t.row === sc.syncedRow ? { ...t, currentIst: res.ist, currentBatch: res.batch } : t)));
      } catch (e) {
        alert(`Austragen in SharePoint fehlgeschlagen: ${e instanceof Error ? e.message : e}`);
        return;
      }
    }
    setScans((p) => p.filter((e) => e.id !== id));
  }

  async function finishShared() {
    setExportMenu(false);
    if (!spFile || !source) return;
    const open = targets.filter((t) => t.currentIst === null || t.currentIst === undefined).length;
    if (!confirm(`Inventur abschließen?\n\n${open} Artikel ohne Ist-Wert werden auf 0 gesetzt (→ nachbestellen). Inventurdatum und „Geprüft von: ${settings.userName || '–'}“ werden eingetragen.`)) return;
    setBusy('Wird in SharePoint eingetragen …');
    try {
      await finishInventory(spFile, source, targets, settings.userName);
      await refreshShared(spFile, true);
      flash('Inventur abgeschlossen');
    } catch (e) {
      alert(`Fehlgeschlagen: ${e instanceof Error ? e.message : e}`);
    } finally {
      setBusy('');
    }
  }

  async function resetShared() {
    if (!spFile || !source) return;
    if (!confirm('Neue Inventur starten?\n\nIst und Chargen-Nr. werden in der gemeinsamen SharePoint-Liste für ALLE Artikel geleert. Die Scans auf diesem Gerät werden ebenfalls gelöscht.')) return;
    setBusy('Liste wird zurückgesetzt …');
    try {
      await resetInventory(spFile, source, targets);
      setScans([]);
      await refreshShared(spFile, true);
      setShowSettings(false);
      flash('Neue Inventur gestartet');
    } catch (e) {
      alert(`Fehlgeschlagen: ${e instanceof Error ? e.message : e}`);
    } finally {
      setBusy('');
    }
  }

  function clear(what: 'scans' | 'targets' | 'mappings') {
    const q = {
      scans: 'Alle Scans löschen? Exportiere vorher, falls du die Daten noch brauchst.',
      targets: 'Soll-Liste entfernen? Die Scans bleiben erhalten.',
      mappings: 'Gelernte Zuordnungen löschen? Danach muss jeder Artikel beim ersten Scan wieder einer Zeile zugeordnet werden.',
    }[what];
    if (!confirm(q)) return;
    if (what === 'scans') setScans([]);
    if (what === 'mappings') setMappings({});
    if (what === 'targets') {
      setTargets([]);
      setSource(null);
      setOriginal('');
      setSettings((s) => ({ ...s, listName: '' }));
    }
  }

  const draftTarget = draft && targets.length ? findTarget(targets, draft, mappings) : undefined;
  const duplicate = draft?.serial ? scans.find((s) => s.serial === draft.serial && normId(s.udiDi) === normId(draft.udiDi)) : undefined;
  const rows = recon.rows.filter((r) => filter === 'alle' || r.status === filter);
  const canWriteOriginal = !!(source && original && (source.istCol || source.batchCol));

  return (
    <div className="app">
      <header className="header">
        <div className="grow">
          <h1>Konsilager</h1>
          <div className="sub">
            {shared ? '☁︎ SharePoint: ' : settings.listName ? 'Liste: ' : ''}
            {settings.listName || 'Keine Soll-Liste importiert'}
            {settings.userName ? ` · ${settings.userName}` : ''}
          </div>
          {shared && lastSync && <div className="sub small">Stand {lastSync.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })}</div>}
        </div>
        {shared && (
          <button className="icon" onClick={() => refreshShared()} aria-label="Aktualisieren" title="Aktualisieren">
            ⟳
          </button>
        )}
        <button className="icon" onClick={() => setShowSettings(true)} aria-label="Einstellungen">
          ⚙︎
        </button>
      </header>

      {!account && (
        <div className="signin">
          <div className="grow">
            <b>Gemeinsame Liste in SharePoint</b>
            <span>Mit dem Firmenkonto anmelden, dann landet jeder Scan sofort in der geteilten Excel-Liste.</span>
          </div>
          <button className="btn primary" onClick={() => login().catch((e) => alert(e?.message ?? e))}>
            Anmelden
          </button>
        </div>
      )}
      {spError && (
        <div className="errorbar" onClick={() => setSpError('')}>
          {spError}
        </div>
      )}
      {shared && scans.some((x) => x.sync === 'pending') && (
        <div className="errorbar warnbar">
          {scans.filter((x) => x.sync === 'pending').length} Scan(s) noch nicht in SharePoint eingetragen.
          <button className="textbtn small" onClick={retryPending}>
            Erneut senden
          </button>
        </div>
      )}

      {targets.length > 0 && (
        <div className="summary">
          {(['vollständig', 'teilweise', 'fehlend', 'unbekannt'] as const).map((k) => (
            <button
              key={k}
              className={`chip st-${k} ${filter === k ? 'active' : ''}`}
              onClick={() => {
                setTab('abgleich');
                setFilter(filter === k ? 'alle' : k);
              }}
            >
              <b>{recon.counts[k]}</b>
              <span>{STATUS_LABEL[k]}</span>
            </button>
          ))}
        </div>
      )}

      <div className="actions">
        <button className="btn primary" onClick={() => setScanning(true)}>
          Scannen
        </button>
        <button className="btn" onClick={() => setDraft(emptyDraft())}>
          Manuell
        </button>
        {shared ? (
          <button className="btn" onClick={() => refreshShared()}>
            Aktualisieren
          </button>
        ) : (
          <button className="btn" onClick={() => fileInput.current?.click()}>
            Soll-Liste
          </button>
        )}
        <button className="btn" onClick={() => (scans.length || targets.length ? setExportMenu(true) : alert('Noch keine Soll-Liste und keine Scans.'))}>
          Export
        </button>
        <input ref={fileInput} type="file" accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden onChange={onFileChosen} />
      </div>

      <nav className="tabs">
        <button className={tab === 'abgleich' ? 'active' : ''} onClick={() => setTab('abgleich')}>
          Abgleich ({targets.length})
        </button>
        <button className={tab === 'scans' ? 'active' : ''} onClick={() => setTab('scans')}>
          Scans ({scans.length})
        </button>
      </nav>

      <main className="list">
        {tab === 'abgleich' && filter === 'unbekannt' && <ScanList scans={recon.unknown} onRemove={removeScan} empty="Alle Scans sind einer Zeile zugeordnet." />}
        {tab === 'abgleich' && filter !== 'unbekannt' && (
          <>
            {rows.length === 0 && (
              <p className="empty">
                {targets.length
                  ? 'Keine Einträge für diesen Filter.'
                  : 'Importiere zuerst die Soll-Liste: Excel-Datei mit Artikel-Nr., Artikelbezeichnung, Seite, Größe, Soll, Ist und Chargen-Nr.'}
              </p>
            )}
            {rows.map((r) => (
              <ReconItem key={r.item.id} row={r} shared={shared} />
            ))}
          </>
        )}
        {tab === 'scans' && (
          <ScanList
            scans={scans}
            onRemove={removeScan}
            empty="Noch nichts gescannt. Tippe auf „Scannen“."
            resolve={(s) => (targets.length ? findTarget(targets, s, mappings) : undefined)}
          />
        )}
      </main>

      {scanning && <Scanner onResult={onScanned} onClose={() => setScanning(false)} />}

      {draft && !picking && (
        <div className="backdrop">
          <div className="sheet">
            <h2>Eintrag prüfen</h2>
            {targets.length > 0 &&
              (draftTarget ? (
                <button className="banner ok" onClick={() => setPicking(true)}>
                  <b>{itemLabel(draftTarget)}</b>
                  {draftTarget.batch && normId(draftTarget.batch) !== normId(draft.batch) && <span className="warn">Achtung: erwartete Charge {draftTarget.batch}</span>}
                  <span className="link">Andere Zeile wählen</span>
                </button>
              ) : (
                <button className="banner unknown" onClick={() => setPicking(true)}>
                  <b>Artikel noch nicht zugeordnet</b>
                  <span>Hier tippen und die passende Zeile (Artikel, Seite, Größe) wählen. Die App merkt sich das für diesen Code.</span>
                </button>
              ))}
            {duplicate && <div className="banner dup">Diese Seriennummer wurde bereits gescannt ({new Date(duplicate.scannedAt).toLocaleString('de-DE')}).</div>}
            <Field label="Chargennummer" value={draft.batch} onChange={(v) => setDraft({ ...draft, batch: v })} autoFocus={!draft.batch} />
            <Field label="UDI-DI / GTIN" value={draft.udiDi ?? ''} onChange={(v) => setDraft({ ...draft, udiDi: v })} />
            <div className="two">
              <Field label="Verfall (JJJJ-MM-TT)" value={draft.expiry ?? ''} onChange={(v) => setDraft({ ...draft, expiry: v })} />
              <Field label="Menge" value={draft.quantity} onChange={(v) => setDraft({ ...draft, quantity: v })} inputMode="decimal" />
            </div>
            {draft.serial && <p className="meta">Seriennummer: {draft.serial}</p>}
            {draft.raw && <p className="meta raw">Rohdaten: {draft.raw.replace(/\u001d/g, '⟨GS⟩')}</p>}
            <div className="actions flush">
              <button className="btn" onClick={() => setDraft(null)}>
                Verwerfen
              </button>
              <button className="btn" onClick={saveDraft}>
                Speichern
              </button>
              <button className="btn primary" onClick={() => saveDraft() && setScanning(true)}>
                + Nächster
              </button>
            </div>
          </div>
        </div>
      )}

      {picking && (
        <TargetPicker
          targets={targets}
          onClose={() => setPicking(false)}
          onPick={(t) => {
            setDraft((d) => (d ? { ...d, targetKey: t.key } : d));
            setPicking(false);
          }}
        />
      )}

      {exportMenu && (
        <div className="backdrop" onClick={() => setExportMenu(false)}>
          <div className="sheet" onClick={(e) => e.stopPropagation()}>
            <h2>{shared ? 'Abschließen & Exportieren' : 'Exportieren'}</h2>
            {shared && (
              <button className="menu-item" onClick={finishShared}>
                <b>Inventur abschließen</b>
                <span>In der SharePoint-Liste: nicht gefundene Artikel auf 0 (→ nachbestellen), Datum und Prüfer eintragen</span>
              </button>
            )}
            {!shared && canWriteOriginal && (
              <button className="menu-item" onClick={() => runExport('original')}>
                <b>In Original-Liste eintragen</b>
                <span>Deine Excel-Datei wie importiert, mit Ist, Chargen-Nr., Datum und Prüfer ausgefüllt</span>
              </button>
            )}
            <button className="menu-item" onClick={() => runExport('report')}>
              <b>Auswertung als Excel</b>
              <span>Abgleich, nicht zugeordnete Scans, alle Scans</span>
            </button>
            <button className="menu-item" onClick={() => runExport('pdf')}>
              <b>Auswertung als PDF</b>
              <span>Öffnet die Druckansicht, dort „Als PDF speichern“ wählen</span>
            </button>
            <div className="actions flush">
              <button className="btn" onClick={() => setExportMenu(false)}>
                Abbrechen
              </button>
            </div>
          </div>
        </div>
      )}

      {ready && (
        <div className="backdrop" onClick={() => setReady(null)}>
          <div className="sheet" onClick={(e) => e.stopPropagation()}>
            <h2>Datei fertig</h2>
            <p className="meta">{ready.name}</p>
            <div className="actions flush">
              <button
                className="btn primary"
                onClick={() => {
                  downloadFile(ready);
                  setReady(null);
                }}
              >
                Herunterladen
              </button>
              {canShareFile(ready) && (
                <button
                  className="btn"
                  onClick={() =>
                    shareFile(ready)
                      .catch(() => {})
                      .finally(() => setReady(null))
                  }
                >
                  Teilen …
                </button>
              )}
            </div>
            <p className="meta">Mit „Teilen“ kannst du die Datei direkt in OneDrive/SharePoint, Teams oder per Mail ablegen.</p>
          </div>
        </div>
      )}

      {showSettings && (
        <div className="backdrop">
          <div className="sheet">
            <h2>Einstellungen</h2>
            <Field label="Dein Name (für „Geprüft von“ und bei jedem Scan)" value={settings.userName} onChange={(v) => setSettings((s) => ({ ...s, userName: v }))} caps="words" />
            <Field label="Name der Inventur" value={settings.listName} onChange={(v) => setSettings((s) => ({ ...s, listName: v }))} caps="sentences" />
            {account && (
              <div className="settings-block">
                <p className="meta">
                  Angemeldet als {account.name ?? account.username}
                  {spFile ? (
                    <>
                      <br />
                      Liste:{' '}
                      {spFile.webUrl ? (
                        <a href={spFile.webUrl} target="_blank" rel="noreferrer">
                          {spFile.name}
                        </a>
                      ) : (
                        spFile.name
                      )}
                    </>
                  ) : null}
                </p>
                <button className="textbtn" onClick={() => setBrowsing(true)}>
                  Andere SharePoint-Datei wählen
                </button>
                {shared && (
                  <button className="danger" onClick={resetShared}>
                    Neue Inventur starten (Ist & Chargen in SharePoint leeren)
                  </button>
                )}
                <button className="textbtn" onClick={() => logout()}>
                  Abmelden
                </button>
              </div>
            )}
            <button className="danger" onClick={() => clear('scans')}>
              {shared ? 'Scan-Verlauf auf diesem Gerät löschen' : 'Alle Scans löschen'}
            </button>
            {!shared && (
              <button className="danger" onClick={() => clear('targets')}>
                Soll-Liste entfernen
              </button>
            )}
            <button className="danger" onClick={() => clear('mappings')}>
              Gelernte Zuordnungen auf diesem Gerät löschen ({Object.keys(mappings).length})
            </button>
            <p className="meta">
              {shared
                ? 'Ist und Chargen-Nr. stehen direkt in der SharePoint-Liste. Gelernte Zuordnungen werden im versteckten Blatt „App-Zuordnungen“ der Datei für alle Kollegen gespeichert.'
                : 'Alle Daten bleiben nur in diesem Browser auf diesem Gerät gespeichert.'}
            </p>
            <div className="actions flush">
              <button className="btn primary" onClick={() => setShowSettings(false)}>
                Fertig
              </button>
            </div>
          </div>
        </div>
      )}

      {browsing && (
        <SpBrowser
          onClose={() => setBrowsing(false)}
          onPick={async (e) => {
            setBrowsing(false);
            try {
              const f = await getFileInfo({ driveId: e.driveId, itemId: e.id });
              setSpFile(f);
              await refreshShared(f);
            } catch (err) {
              setSpError(`Datei konnte nicht geöffnet werden: ${err instanceof Error ? err.message : err}`);
            }
          }}
        />
      )}

      {busy && <div className="toast">{busy}</div>}
      {toast && !busy && <div className="toast">{toast}</div>}
    </div>
  );
}

function TargetPicker({ targets, onClose, onPick }: { targets: TargetItem[]; onClose: () => void; onPick: (t: TargetItem) => void }) {
  const [q, setQ] = useState('');
  const words = q.toLowerCase().split(/\s+/).filter(Boolean);
  const list = targets.filter((t) => words.every((w) => itemLabel(t).toLowerCase().includes(w)));
  return (
    <div className="fullscreen">
      <header className="header">
        <h1 className="grow" style={{ fontSize: 22 }}>
          Zeile auswählen
        </h1>
        <button className="textbtn" onClick={onClose}>
          Abbrechen
        </button>
      </header>
      <div className="pad">
        <input className="input" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Suchen, z. B. „neurodyn rechts m“" autoFocus />
      </div>
      <div className="list">
        {list.map((t) => (
          <button key={t.id} className="row pick" onClick={() => onPick(t)}>
            <div className="grow">
              <div className="row-title">{t.name}</div>
              <div className="meta">
                Art.-Nr. {t.articleNo}
                {t.side ? ` · ${t.side}` : ''}
              </div>
            </div>
            {t.size && <span className="size">{t.size}</span>}
          </button>
        ))}
      </div>
    </div>
  );
}

function ReconItem({ row, shared }: { row: ReconRow; shared?: boolean }) {
  const batches = shared ? (row.item.currentBatch ?? '') : [...new Set(row.scans.map((s) => s.batch))].join(', ');
  const open = shared && (row.item.currentIst === null || row.item.currentIst === undefined);
  return (
    <div className={`row st-${row.status}`}>
      <div className="grow">
        <div className="row-title">{row.item.name || row.item.articleNo}</div>
        <div className="meta">
          Art.-Nr. {row.item.articleNo || '–'}
          {row.item.side ? ` · ${row.item.side}` : ''}
          {row.item.size ? ` · Gr. ${row.item.size}` : ''}
          {row.item.batch ? ` · Charge soll ${row.item.batch}` : ''}
        </div>
        {batches && <div className="meta">Charge: {batches}</div>}
      </div>
      <div className="right">
        <div className="qty">
          {row.foundQty}/{row.item.expectedQty}
        </div>
        <div className="status">{open ? 'offen' : row.status === 'fehlend' ? 'FEHLT' : row.status}</div>
      </div>
    </div>
  );
}

function ScanList({ scans, onRemove, empty, resolve }: { scans: Scan[]; onRemove: (id: string) => void; empty: string; resolve?: (s: Scan) => TargetItem | undefined }) {
  if (!scans.length) return <p className="empty">{empty}</p>;
  return (
    <>
      {scans.map((s) => {
        const t = resolve?.(s);
        return (
          <div key={s.id} className="row">
            <div className="grow">
              <div className="row-title">Charge {s.batch}</div>
              {t && <div className="meta">{itemLabel(t)}</div>}
              <div className="meta">{[s.udiDi && `UDI-DI ${s.udiDi}`, s.expiry && `Verfall ${s.expiry}`].filter(Boolean).join(' · ') || s.codeType}</div>
              <div className="meta">
                {new Date(s.scannedAt).toLocaleString('de-DE')}
                {s.scannedBy ? ` · ${s.scannedBy}` : ''}
              </div>
            </div>
            <div className="right">
              <div className="qty">×{s.quantity}</div>
              <button className="textbtn small" onClick={() => onRemove(s.id)} aria-label="Scan löschen">
                Löschen
              </button>
            </div>
          </div>
        );
      })}
    </>
  );
}

function Field(props: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  autoFocus?: boolean;
  inputMode?: 'text' | 'decimal';
  caps?: 'characters' | 'words' | 'sentences';
}) {
  return (
    <label className="field">
      <span>{props.label}</span>
      <input
        className="input"
        value={props.value}
        onChange={(e) => props.onChange(e.target.value)}
        autoFocus={props.autoFocus}
        inputMode={props.inputMode ?? 'text'}
        autoCapitalize={props.caps ?? 'characters'}
        autoCorrect="off"
        spellCheck={false}
      />
    </label>
  );
}

function SpBrowser({ onClose, onPick }: { onClose: () => void; onPick: (e: BrowseEntry) => void }) {
  const [stack, setStack] = useState<{ title: string; entries: BrowseEntry[] }[]>([]);
  const [err, setErr] = useState('');
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    (async () => {
      try {
        const site = await siteFromUrl(MS_CONFIG.defaultFileUrl);
        if (!site) throw new Error('Website nicht gefunden.');
        setStack([{ title: 'Dokumentbibliotheken', entries: await listDrives(site) }]);
      } catch (e) {
        setErr(e instanceof Error ? e.message : String(e));
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  async function open(e: BrowseEntry) {
    if (e.kind === 'file') return onPick(e);
    setLoading(true);
    try {
      setStack((s) => [...s]);
      const entries = await listFolder(e.driveId, e.id);
      setStack((s) => [...s, { title: e.name, entries }]);
    } catch (x) {
      setErr(x instanceof Error ? x.message : String(x));
    } finally {
      setLoading(false);
    }
  }

  const cur = stack[stack.length - 1];
  return (
    <div className="fullscreen">
      <header className="header">
        {stack.length > 1 && (
          <button className="textbtn" onClick={() => setStack((s) => s.slice(0, -1))}>
            ‹ Zurück
          </button>
        )}
        <h1 className="grow" style={{ fontSize: 20 }}>
          {cur?.title ?? 'SharePoint'}
        </h1>
        <button className="textbtn" onClick={onClose}>
          Abbrechen
        </button>
      </header>
      <p className="meta pad">Bitte die Excel-Bestandsliste auswählen.</p>
      {err && <div className="errorbar">{err}</div>}
      {loading && <p className="empty">Wird geladen …</p>}
      <div className="list">
        {!loading &&
          cur?.entries.map((e) => (
            <button key={e.driveId + e.id} className="row pick" onClick={() => open(e)}>
              <span className="size">{e.kind === 'file' ? '📄' : '📁'}</span>
              <div className="grow row-title">{e.name}</div>
            </button>
          ))}
        {!loading && cur && cur.entries.length === 0 && <p className="empty">Keine Ordner oder Excel-Dateien.</p>}
      </div>
    </div>
  );
}
