// Seite „Lagerinventur“: alles scannen, Artikel über den Artikelstamm erkennen, gemeinsame Excel-Datei in SharePoint.
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { AccountInfo } from '@azure/msal-browser';
import { Scanner } from './Scanner';
import { __setTokenProvider, initAuth, login, logout, me, type FileRef } from './lib/msGraph';
import { parseUdi, normId } from './lib/udi';
import { load, save, usePersisted } from './lib/storage';
import { LAGER_FOLDER_URL } from './config';
import { builtinArticles, parseCatalog, summarize, type Article, type LagerScan } from './lib/lagerData';
import {
  appendArticle, appendScans, createInventory, deleteScan, loadCatalogRows, readArticles, readScans, resolveFolder, scanFolder, writeSummary, type FolderRef,
} from './lib/lagerGraph';

const K = { folder: 'lager.folder', file: 'lager.file', pending: 'lager.pending', mine: 'lager.mine', quick: 'lager.quick' };
const today = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};
const newId = () => `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
const errText = (e: unknown) => (e instanceof Error ? e.message : String(e));

type Unknown = { gtin: string; batch: string; expiry: string; raw: string };

export default function Lager() {
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [authReady, setAuthReady] = useState(false);
  const [person, setPerson] = useState('');
  const [folder, setFolder] = usePersisted<FolderRef | null>(K.folder, null);
  const [file, setFile] = usePersisted<FileRef | null>(K.file, null);
  const [pending, setPending] = usePersisted<LagerScan[]>(K.pending, []); // noch nicht in der Datei
  const [mine, setMine] = usePersisted<string[]>(K.mine, []); // IDs eigener Scans (zum Löschen)
  const [shared, setShared] = useState<LagerScan[]>([]);
  const [learned, setLearned] = useState<Article[]>([]);
  const [catalog, setCatalog] = useState<Map<string, Article> | null>(null);
  const [catalogInfo, setCatalogInfo] = useState('');
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [scanning, setScanning] = useState(false);
  const [feedback, setFeedback] = useState('');
  const [unknown, setUnknown] = useState<Unknown | null>(null);
  const [tab, setTab] = useState<'uebersicht' | 'meine'>('uebersicht');
  const [folderUrl, setFolderUrl] = useState(LAGER_FOLDER_URL);
  const [menu, setMenu] = useState(false);
  const [noInventory, setNoInventory] = useState(false);
  const [lastSync, setLastSync] = useState<Date | null>(null);
  const builtin = useMemo(builtinArticles, []);

  // ---------- Anmeldung ----------
  useEffect(() => {
    // Nur Entwicklung: Oberfläche ohne echtes Firmenkonto testen
    if (import.meta.env.DEV && localStorage.getItem('lager.test')) {
      __setTokenProvider(async () => 'TEST');
      setAccount({ username: 'test@example.com', name: 'Test' } as AccountInfo);
      setPerson('Test Person');
      setAuthReady(true);
      return;
    }
    initAuth()
      .then(async (acc) => {
        setAccount(acc);
        if (acc) setPerson((await me().catch(() => null))?.displayName || acc.name || acc.username);
      })
      .catch((e) => setError('Anmeldung: ' + errText(e)))
      .finally(() => setAuthReady(true));
  }, []);

  // Ordner aus fest eingestelltem Link übernehmen
  useEffect(() => {
    if (account && !folder && LAGER_FOLDER_URL) resolveFolder(LAGER_FOLDER_URL).then(setFolder).catch(() => {});
  }, [account, folder, setFolder]);

  // ---------- Ordner, Artikelstamm und aktuelle Inventur laden ----------
  const refresh = useCallback(
    async (f: FileRef | null = file) => {
      if (!f) return;
      try {
        const [s, a] = await Promise.all([readScans(f), readArticles(f)]);
        setShared(s);
        setLearned(a);
        setLastSync(new Date());
        setError('');
      } catch (e) {
        setError('Inventurdatei konnte nicht gelesen werden: ' + errText(e));
      }
    },
    [file],
  );

  useEffect(() => {
    if (!account || !folder) return;
    let cancelled = false;
    (async () => {
      try {
        setBusy('Ordner wird gelesen …');
        const c = await scanFolder(folder);
        if (cancelled) return;
        const current = c.inventories[0] ?? null;
        setNoInventory(!current);
        setFile(current);
        if (current) await refresh(current);
        if (c.catalog) {
          setBusy('Artikelliste wird geladen …');
          const rows = await loadCatalogRows(c.catalog);
          const info = parseCatalog(rows);
          if (cancelled) return;
          setCatalog(info.articles);
          setCatalogInfo(`${c.catalog.name}: ${info.withGtin.toLocaleString('de-DE')} Artikel mit EAN`);
        } else {
          setCatalog(new Map());
          setCatalogInfo('Keine Artikelliste im Ordner gefunden (Datei mit „Artikel“ oder „EAN“ im Namen).');
        }
      } catch (e) {
        setError(errText(e));
      } finally {
        if (!cancelled) setBusy('');
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [account, folder?.itemId]);

  // Beim Zurückkehren in die App den Stand der Kollegen holen
  useEffect(() => {
    const onVis = () => document.visibilityState === 'visible' && !scanning && refresh();
    document.addEventListener('visibilitychange', onVis);
    return () => document.removeEventListener('visibilitychange', onVis);
  }, [refresh, scanning]);

  // ---------- Senden (Warteschlange) und Übersicht schreiben ----------
  const flushing = useRef(false);
  const summaryTimer = useRef(0);
  const flush = useCallback(async () => {
    if (flushing.current || !file) return;
    const batch = load<LagerScan[]>(K.pending, []);
    if (!batch.length) return;
    flushing.current = true;
    try {
      await appendScans(file, batch);
      const ids = new Set(batch.map((s) => s.id));
      setPending((p) => {
        const rest = p.filter((s) => !ids.has(s.id));
        save(K.pending, rest);
        return rest;
      });
      setShared((s) => [...s, ...batch]);
      setError('');
      // Übersicht nach kurzer Pause neu schreiben (mit den Scans aller Kollegen)
      clearTimeout(summaryTimer.current);
      summaryTimer.current = window.setTimeout(async () => {
        try {
          const all = await readScans(file);
          setShared(all);
          setLastSync(new Date());
          await writeSummary(file, all, file.name.replace(/\.xlsx$/i, ''));
        } catch (e) {
          console.warn('Übersicht nicht geschrieben', e);
        }
      }, 4000);
    } catch (e) {
      setError(`${batch.length} Scan(s) noch nicht in der Excel-Datei (${errText(e)}). Wird automatisch erneut versucht.`);
    } finally {
      flushing.current = false;
    }
  }, [file, setPending]);

  useEffect(() => {
    if (!pending.length || !file || !account) return;
    const t = window.setTimeout(flush, 300);
    return () => clearTimeout(t);
  }, [pending, file, account, flush]);
  useEffect(() => {
    const t = window.setInterval(() => load<LagerScan[]>(K.pending, []).length && flush(), 20000);
    return () => clearInterval(t);
  }, [flush]);

  // ---------- Artikel erkennen ----------
  const learnedMap = useMemo(() => new Map(learned.map((a) => [normId(a.gtin) || a.artNr.toLowerCase(), a])), [learned]);
  const lookup = useCallback(
    (gtin: string): Article | undefined => {
      const k = normId(gtin);
      if (!k) return undefined;
      return learnedMap.get(k) ?? catalog?.get(k) ?? builtin.get(k);
    },
    [learnedMap, catalog, builtin],
  );

  const addScan = useCallback(
    (a: Article, u: { gtin: string; batch: string; expiry: string }) => {
      const s: LagerScan = {
        id: newId(), time: new Date().toISOString(), person, hersteller: a.hersteller, produktart: a.produktart, artNr: a.artNr, name: a.name,
        gtin: u.gtin || a.gtin, batch: u.batch, expiry: u.expiry, qty: 1,
      };
      setPending((p) => {
        const next = [...p, s];
        save(K.pending, next);
        return next;
      });
      setMine((m) => [s.id, ...m].slice(0, 2000));
      return s;
    },
    [person, setPending, setMine],
  );

  const onScanned = useCallback(
    (text: string): boolean => {
      const udi = parseUdi(text);
      const gtin = udi.udiDi ?? (/^\d{8,14}$/.test(text.trim()) ? text.trim() : '');
      const u = { gtin, batch: udi.batch ?? '', expiry: udi.expiry ?? '' };
      const a = lookup(gtin);
      if (a) {
        addScan(a, u);
        setFeedback(`✓ ${a.name}${u.batch ? ` · Charge ${u.batch}` : ''}`);
        return true;
      }
      setScanning(false);
      setFeedback('');
      setUnknown({ ...u, raw: text });
      return false;
    },
    [lookup, addScan],
  );

  async function saveUnknown(a: Article) {
    if (!unknown) return;
    setUnknown(null);
    setLearned((l) => [...l, a]);
    addScan(a, unknown);
    if (file) appendArticle(file, a).catch((e) => setError('Artikel nicht gespeichert: ' + errText(e)));
  }

  async function removeScan(s: LagerScan) {
    if (!confirm(`Scan löschen?\n\n${s.name}${s.batch ? `, Charge ${s.batch}` : ''}`)) return;
    if (pending.some((p) => p.id === s.id)) {
      setPending((p) => {
        const next = p.filter((x) => x.id !== s.id);
        save(K.pending, next);
        return next;
      });
      return;
    }
    if (!file) return;
    try {
      setBusy('Wird gelöscht …');
      await deleteScan(file, s.id);
      const all = await readScans(file);
      setShared(all);
      await writeSummary(file, all, file.name.replace(/\.xlsx$/i, ''));
    } catch (e) {
      setError('Löschen fehlgeschlagen: ' + errText(e));
    } finally {
      setBusy('');
    }
  }

  async function startInventory() {
    if (!folder) return;
    setMenu(false);
    if (file && !confirm('Neue Inventur starten?\n\nEs wird eine neue Excel-Datei angelegt. Die bisherige Datei bleibt unverändert im Ordner. Alle Handys arbeiten ab dann mit der neuen Datei.')) return;
    try {
      setBusy('Neue Inventurdatei wird angelegt …');
      const f = await createInventory(folder, today(), learned);
      setFile(f);
      setNoInventory(false);
      setShared([]);
      setPending([]);
      save(K.pending, []);
      setMine([]);
      await refresh(f);
    } catch (e) {
      setError('Inventurdatei konnte nicht angelegt werden: ' + errText(e));
    } finally {
      setBusy('');
    }
  }

  async function chooseFolder() {
    try {
      setBusy('Ordner wird gesucht …');
      setFolder(await resolveFolder(folderUrl));
      setFile(null);
    } catch (e) {
      setError('Ordner nicht gefunden: ' + errText(e) + ' Bitte in SharePoint beim Ordner „⋯ → Link kopieren“ verwenden.');
    } finally {
      setBusy('');
    }
  }

  // ---------- Anzeige ----------
  const allScans = useMemo(() => {
    const ids = new Set(shared.map((s) => s.id));
    return [...shared, ...pending.filter((p) => !ids.has(p.id))];
  }, [shared, pending]);
  const groups = useMemo(() => summarize(allScans), [allScans]);
  const myScans = useMemo(() => {
    const set = new Set(mine);
    return allScans.filter((s) => set.has(s.id)).reverse();
  }, [allScans, mine]);
  const total = allScans.reduce((n, s) => n + s.qty, 0);
  const pendingIds = new Set(pending.map((p) => p.id));
  const herstellerList = useMemo(() => {
    const set = new Set<string>();
    catalog?.forEach((a) => a.hersteller && set.add(a.hersteller));
    learned.forEach((a) => a.hersteller && set.add(a.hersteller));
    return [...set].sort((a, b) => a.localeCompare(b, 'de'));
  }, [catalog, learned]);

  const ready = account && folder && file && catalog;

  return (
    <div className="app">
      <header className="header">
        <div className="grow">
          <h1>Lagerinventur</h1>
          <div className="sub">
            {file ? file.name.replace(/\.xlsx$/i, '') : 'Gemeinsame Liste in SharePoint'}
            {person ? ` · ${person}` : ''}
          </div>
          {lastSync && <div className="sub small">Stand {lastSync.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })}</div>}
        </div>
        {file && (
          <button className="icon" onClick={() => refresh()} aria-label="Aktualisieren" title="Aktualisieren">
            ⟳
          </button>
        )}
        {account && (
          <button className="icon" onClick={() => setMenu(true)} aria-label="Menü">
            ⚙︎
          </button>
        )}
      </header>

      {error && (
        <div className="errorbar" onClick={() => setError('')}>
          {error}
        </div>
      )}

      {authReady && !account && (
        <div className="signin">
          <div className="grow">
            <b>Mit dem Firmenkonto anmelden</b>
            <span>Jeder Scan landet sofort in der gemeinsamen Excel-Datei in SharePoint.</span>
          </div>
          <button className="btn primary" onClick={() => login().catch((e) => setError(errText(e)))}>
            Anmelden
          </button>
        </div>
      )}

      {account && !folder && (
        <div className="sheet inline">
          <h2>SharePoint-Ordner wählen</h2>
          <p className="meta">In SharePoint beim Ordner auf „⋯ → Link kopieren“ tippen und den Link hier einfügen. Dort liegen die Artikelliste und die Inventurdateien.</p>
          <label className="field">
            <span>Link zum Ordner</span>
            <input className="input" value={folderUrl} onChange={(e) => setFolderUrl(e.target.value)} placeholder="https://…sharepoint.com/…" autoCapitalize="none" autoCorrect="off" />
          </label>
          <button className="btn primary wide" disabled={!folderUrl.trim()} onClick={chooseFolder}>
            Ordner übernehmen
          </button>
        </div>
      )}

      {account && folder && noInventory && !busy && (
        <div className="signin">
          <div className="grow">
            <b>Noch keine Inventur in diesem Ordner</b>
            <span>Legt die Datei „Lagerinventur {today()}.xlsx“ an.</span>
          </div>
          <button className="btn primary" onClick={startInventory}>
            Starten
          </button>
        </div>
      )}

      {account && folder && file && (
        <>
          <div className="summary two">
            <div className="chip">
              <b>{total}</b>
              <span>Stück gezählt</span>
            </div>
            <div className="chip">
              <b>{groups.length}</b>
              <span>Produktgruppen</span>
            </div>
          </div>
          <div className="actions">
            <button className="btn primary" disabled={!ready} onClick={() => setScanning(true)}>
              {catalog ? 'Scannen' : 'Artikelliste lädt …'}
            </button>
            {file.webUrl && (
              <a className="btn linkbtn" href={file.webUrl} target="_blank" rel="noreferrer">
                In Excel öffnen
              </a>
            )}
          </div>
          {pending.length > 0 && (
            <div className="errorbar warnbar">
              {pending.length} Scan(s) werden noch übertragen …
              <button className="textbtn small" onClick={flush}>
                Jetzt senden
              </button>
            </div>
          )}

          <div className="tabs">
            <button className={tab === 'uebersicht' ? 'active' : ''} onClick={() => setTab('uebersicht')}>
              Übersicht
            </button>
            <button className={tab === 'meine' ? 'active' : ''} onClick={() => setTab('meine')}>
              Meine Scans ({myScans.length})
            </button>
          </div>

          {tab === 'uebersicht' &&
            (groups.length ? (
              groups.map((g, i) => (
                <div key={g.hersteller + g.produktart}>
                  {(i === 0 || groups[i - 1].hersteller !== g.hersteller) && <h3 className="group maker">{g.hersteller}</h3>}
                  <div className="lager-group">
                    <div className="lager-group-head">
                      <span>{g.produktart}</span>
                      <b>{g.qty}</b>
                    </div>
                    {g.items.map((it) => (
                      <div key={(it.gtin || it.artNr) + it.name} className="lager-item">
                        <div className="grow">
                          <div>{it.name || it.artNr}</div>
                          <div className="meta">
                            {it.artNr && `Art.-Nr. ${it.artNr}`}
                            {it.batches.length > 0 && ` · Charge ${[...new Set(it.batches)].join(', ')}`}
                          </div>
                        </div>
                        <b className="qty">×{it.qty}</b>
                      </div>
                    ))}
                  </div>
                </div>
              ))
            ) : (
              <p className="empty">Noch nichts gescannt.</p>
            ))}

          {tab === 'meine' &&
            (myScans.length ? (
              myScans.map((s) => (
                <div key={s.id} className="row">
                  <div className="grow">
                    <div className="row-title">{s.name || s.artNr || s.gtin}</div>
                    <div className="meta">
                      {s.hersteller}
                      {s.batch && ` · Charge ${s.batch}`}
                      {s.expiry && ` · Verfall ${s.expiry}`}
                    </div>
                    <div className="meta">
                      {new Date(s.time).toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' })}
                      {pendingIds.has(s.id) ? ' · wird übertragen' : ' · in Excel'}
                    </div>
                  </div>
                  <button className="textbtn small" onClick={() => removeScan(s)}>
                    Löschen
                  </button>
                </div>
              ))
            ) : (
              <p className="empty">Du hast in dieser Inventur noch nichts gescannt.</p>
            ))}
          {catalogInfo && <p className="meta footnote">{catalogInfo}</p>}
        </>
      )}

      {scanning && (
        <Scanner
          continuous
          feedback={feedback}
          onResult={onScanned}
          onClose={() => {
            setScanning(false);
            setFeedback('');
          }}
        />
      )}

      {unknown && (
        <UnknownDialog
          data={unknown}
          herstellerList={herstellerList}
          produktarten={(h) => {
            const set = new Set<string>();
            const add = (a: Article) => a.hersteller === h && a.produktart && set.add(a.produktart);
            learned.forEach(add);
            if (set.size < 300) catalog?.forEach(add);
            return [...set].sort((a, b) => a.localeCompare(b, 'de')).slice(0, 300);
          }}
          onCancel={() => setUnknown(null)}
          onSave={saveUnknown}
        />
      )}

      {menu && (
        <div className="backdrop" onClick={() => setMenu(false)}>
          <div className="sheet" onClick={(e) => e.stopPropagation()}>
            <h2>Lagerinventur</h2>
            <button className="menu-item" onClick={startInventory}>
              <b>Neue Inventur starten</b>
              <span>Legt eine neue Excel-Datei an, die bisherige bleibt erhalten</span>
            </button>
            <button
              className="menu-item"
              onClick={() => {
                setMenu(false);
                setFolder(null);
                setFile(null);
              }}
            >
              <b>Anderen Ordner wählen</b>
              <span>{folder?.name ?? '–'}</span>
            </button>
            <button className="menu-item" onClick={() => logout()}>
              <b>Abmelden</b>
              <span>{account?.username}</span>
            </button>
            <a className="menu-item" href="./">
              <b>Zur Konsilager-Inventur</b>
              <span>Abgleich mit der Sporlastic-Liste</span>
            </a>
            <button className="btn wide" onClick={() => setMenu(false)}>
              Schließen
            </button>
          </div>
        </div>
      )}

      {busy && <div className="toast">{busy}</div>}
    </div>
  );
}

function UnknownDialog(props: {
  data: Unknown;
  herstellerList: string[];
  produktarten: (hersteller: string) => string[];
  onCancel: () => void;
  onSave: (a: Article) => void;
}) {
  const [name, setName] = useState('');
  const [hersteller, setHersteller] = useState('');
  const [produktart, setProduktart] = useState('');
  const [artNr, setArtNr] = useState('');
  const arten = useMemo(() => (hersteller ? props.produktarten(hersteller) : []), [hersteller, props]);
  const ok = name.trim() && hersteller.trim();
  return (
    <div className="backdrop">
      <div className="sheet">
        <h2>Unbekannter Artikel</h2>
        <div className="banner unknown">
          Dieser Code steht nicht in der Artikelliste. Einmal ausfüllen, danach erkennen ihn alle Handys.
        </div>
        <p className="meta">
          {props.data.gtin ? `GTIN ${props.data.gtin}` : `Code: ${props.data.raw.slice(0, 60)}`}
          {props.data.batch && ` · Charge ${props.data.batch}`}
        </p>
        <label className="field">
          <span>Bezeichnung (mit Größe)</span>
          <input className="input" value={name} onChange={(e) => setName(e.target.value)} autoFocus autoCapitalize="sentences" />
        </label>
        <label className="field">
          <span>Hersteller</span>
          <input className="input" list="lager-hersteller" value={hersteller} onChange={(e) => setHersteller(e.target.value)} autoCapitalize="words" />
          <datalist id="lager-hersteller">
            {props.herstellerList.map((h) => (
              <option key={h} value={h} />
            ))}
          </datalist>
        </label>
        <label className="field">
          <span>Produktart (Gruppe in der Liste)</span>
          <input className="input" list="lager-arten" value={produktart} onChange={(e) => setProduktart(e.target.value)} autoCapitalize="sentences" />
          <datalist id="lager-arten">
            {arten.map((h) => (
              <option key={h} value={h} />
            ))}
          </datalist>
        </label>
        <label className="field">
          <span>Artikelnummer (optional)</span>
          <input className="input" value={artNr} onChange={(e) => setArtNr(e.target.value)} autoCapitalize="characters" />
        </label>
        <div className="row-buttons">
          <button className="btn" onClick={props.onCancel}>
            Verwerfen
          </button>
          <button
            className="btn primary"
            disabled={!ok}
            onClick={() =>
              props.onSave({
                gtin: props.data.gtin,
                artNr: artNr.trim(),
                name: name.trim(),
                hersteller: hersteller.trim(),
                produktart: produktart.trim() || name.trim(),
              })
            }
          >
            Speichern
          </button>
        </div>
      </div>
    </div>
  );
}
