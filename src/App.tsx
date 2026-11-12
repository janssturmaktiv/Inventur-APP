import { useCallback, useMemo, useRef, useState } from 'react';
import { Scanner } from './Scanner';
import { parseUdi, normId } from './lib/udi';
import { parseTargetWorkbook } from './lib/targetList';
import { reconcile, findTarget, type ReconRow, type Status } from './lib/reconcile';
import { buildFilledOriginal, buildReport, canShareFile, downloadFile, printReport, shareFile, type ExportFile } from './lib/exportWeb';
import { usePersisted } from './lib/storage';
import { itemLabel, type Mappings, type Scan, type TargetItem, type TargetSource } from './lib/types';

const KEYS = {
  scans: 'konsi/scans/v1',
  targets: 'konsi/targets/v2',
  source: 'konsi/source/v1',
  original: 'konsi/original/v1',
  mappings: 'konsi/mappings/v1',
  settings: 'konsi/settings/v1',
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

  const recon = useMemo(() => reconcile(targets, scans, mappings), [targets, scans, mappings]);
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
    if (scan.targetKey && idKey) setMappings((m) => ({ ...m, [idKey]: scan.targetKey! }));
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

  function removeScan(id: string) {
    if (confirm('Diesen Scan löschen?')) setScans((p) => p.filter((e) => e.id !== id));
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
            {settings.listName ? `Liste: ${settings.listName}` : 'Keine Soll-Liste importiert'}
            {settings.userName ? ` · ${settings.userName}` : ''}
          </div>
        </div>
        <button className="icon" onClick={() => setShowSettings(true)} aria-label="Einstellungen">
          ⚙︎
        </button>
      </header>

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
        <button className="btn" onClick={() => fileInput.current?.click()}>
          Soll-Liste
        </button>
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
              <ReconItem key={r.item.id} row={r} />
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
            <h2>Exportieren</h2>
            {canWriteOriginal && (
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
            <button className="danger" onClick={() => clear('scans')}>
              Alle Scans löschen
            </button>
            <button className="danger" onClick={() => clear('targets')}>
              Soll-Liste entfernen
            </button>
            <button className="danger" onClick={() => clear('mappings')}>
              Gelernte Zuordnungen löschen ({Object.keys(mappings).length})
            </button>
            <p className="meta">Alle Daten bleiben nur in diesem Browser auf diesem Gerät gespeichert.</p>
            <div className="actions flush">
              <button className="btn primary" onClick={() => setShowSettings(false)}>
                Fertig
              </button>
            </div>
          </div>
        </div>
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

function ReconItem({ row }: { row: ReconRow }) {
  const batches = [...new Set(row.scans.map((s) => s.batch))].join(', ');
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
        <div className="status">{row.status === 'fehlend' ? 'FEHLT' : row.status}</div>
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
