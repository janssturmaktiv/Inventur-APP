// Anmeldung mit dem Firmenkonto und Zugriff auf die Excel-Datei in SharePoint (Microsoft Graph).
import { createStandardPublicClientApplication, InteractionRequiredAuthError, type IPublicClientApplication, type AccountInfo } from '@azure/msal-browser';
import { MS_CONFIG } from '../config';

const SCOPES = ['User.Read', 'Files.ReadWrite.All'];
const GRAPH = 'https://graph.microsoft.com/v1.0';

let pca: IPublicClientApplication | null = null;

function redirectUri() {
  return window.location.origin + window.location.pathname.replace(/index\.html$/, '');
}

/** Initialisiert die Anmeldung und verarbeitet eine Rückkehr von der Microsoft-Anmeldeseite. */
export async function initAuth(): Promise<AccountInfo | null> {
  pca ??= await createStandardPublicClientApplication({
    auth: {
      clientId: MS_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${MS_CONFIG.tenantId}`,
      redirectUri: redirectUri(),
    },
    cache: { cacheLocation: 'localStorage' },
  });
  const res = await pca.handleRedirectPromise();
  if (res?.account) pca.setActiveAccount(res.account);
  const acc = pca.getActiveAccount() ?? pca.getAllAccounts()[0] ?? null;
  if (acc) pca.setActiveAccount(acc);
  return acc;
}

export async function login() {
  await pca!.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
}

export async function logout() {
  await pca!.logoutRedirect({ postLogoutRedirectUri: redirectUri() });
}

// Für Tests: eigene Token-Quelle setzen
let tokenOverride: (() => Promise<string>) | null = null;
export function __setTokenProvider(fn: (() => Promise<string>) | null) {
  tokenOverride = fn;
}

async function token(): Promise<string> {
  if (tokenOverride) return tokenOverride();
  const account = pca?.getActiveAccount();
  if (!pca || !account) throw new Error('Nicht angemeldet.');
  try {
    return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      await pca.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw e;
  }
}

export class GraphError extends Error {
  constructor(public status: number, message: string) {
    super(message);
  }
}

async function graph<T = unknown>(path: string, init: RequestInit & { raw?: boolean } = {}): Promise<T> {
  const t = await token();
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
      ...init,
      headers: { Authorization: `Bearer ${t}`, ...(init.body ? { 'Content-Type': 'application/json' } : {}), ...init.headers },
    });
    if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 4) {
      const wait = Number(res.headers.get('Retry-After')) || 2 ** attempt;
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (!res.ok) {
      let msg = res.statusText;
      try {
        const j = await res.json();
        msg = j?.error?.message ?? msg;
      } catch {
        /* ignore */
      }
      throw new GraphError(res.status, `${res.status}: ${msg}`);
    }
    if (init.raw) return res as unknown as T;
    if (res.status === 204) return undefined as T;
    return (await res.json()) as T;
  }
}

export type FileRef = { driveId: string; itemId: string; name: string; webUrl?: string };

function shareId(url: string) {
  const b64 = btoa(unescape(encodeURIComponent(url))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${b64}`;
}

type DriveItem = { id: string; name: string; webUrl?: string; folder?: unknown; file?: unknown; parentReference?: { driveId: string } };

/** Datei anhand eines SharePoint-Links finden. */
export async function resolveFile(url: string): Promise<FileRef> {
  // 1. Freigabe-API (funktioniert mit „Link kopieren“-Links)
  try {
    const it = await graph<DriveItem>(`/shares/${shareId(url)}/driveItem?$select=id,name,webUrl,parentReference`);
    if (it?.parentReference?.driveId) return { driveId: it.parentReference.driveId, itemId: it.id, name: it.name, webUrl: it.webUrl };
  } catch {
    /* weiter mit Plan 2 */
  }
  // 2. Dokument-ID (sourcedoc={GUID}) in der Website suchen
  const u = new URL(url);
  const guid = (u.searchParams.get('sourcedoc') ?? '').replace(/[{}]/g, '').toLowerCase();
  const site = await siteFromUrl(url);
  if (guid && site) {
    const drives = await graph<{ value: { id: string }[] }>(`/sites/${site}/drives?$select=id`);
    for (const d of drives.value) {
      try {
        // SharePoint erlaubt den Zugriff über die eindeutige Listen-ID per sharepointIds-Suche
        const res = await graph<{ value: (DriveItem & { sharepointIds?: { listItemUniqueId?: string } })[] }>(
          `/drives/${d.id}/root/search(q='.xlsx')?$select=id,name,webUrl,parentReference,sharepointIds&$top=200`,
        );
        const hit = res.value.find((x) => x.sharepointIds?.listItemUniqueId?.toLowerCase() === guid);
        if (hit) return { driveId: d.id, itemId: hit.id, name: hit.name, webUrl: hit.webUrl };
      } catch {
        /* nächstes Laufwerk */
      }
    }
  }
  throw new Error('Die Datei wurde über den Link nicht gefunden. Bitte die Datei in der Ordneransicht auswählen.');
}

/** Website-ID aus einer SharePoint-Adresse (…/sites/<Name>/…). */
export async function siteFromUrl(url: string): Promise<string | null> {
  const u = new URL(url);
  const m = /\/(sites|teams)\/([^/]+)/.exec(u.pathname);
  if (!m) return null;
  const s = await graph<{ id: string }>(`/sites/${u.hostname}:/${m[1]}/${m[2]}?$select=id`);
  return s.id;
}

export type BrowseEntry = { id: string; name: string; kind: 'drive' | 'folder' | 'file'; driveId: string };

export async function listDrives(siteId: string): Promise<BrowseEntry[]> {
  const r = await graph<{ value: { id: string; name: string }[] }>(`/sites/${siteId}/drives?$select=id,name`);
  return r.value.map((d) => ({ id: 'root', name: d.name, kind: 'drive', driveId: d.id }));
}

export async function listFolder(driveId: string, itemId: string): Promise<BrowseEntry[]> {
  const path = itemId === 'root' ? `/drives/${driveId}/root/children` : `/drives/${driveId}/items/${itemId}/children`;
  const r = await graph<{ value: DriveItem[] }>(`${path}?$select=id,name,folder,file&$top=500`);
  return r.value
    .filter((x) => x.folder || /\.xlsx$/i.test(x.name))
    .map((x) => ({ id: x.id, name: x.name, kind: x.folder ? 'folder' : 'file', driveId }) as BrowseEntry)
    .sort((a, b) => (a.kind === b.kind ? a.name.localeCompare(b.name) : a.kind === 'folder' ? -1 : 1));
}

export async function getFileInfo(f: Pick<FileRef, 'driveId' | 'itemId'>): Promise<FileRef> {
  const it = await graph<DriveItem>(`/drives/${f.driveId}/items/${f.itemId}?$select=id,name,webUrl`);
  return { driveId: f.driveId, itemId: it.id, name: it.name, webUrl: it.webUrl };
}

/** Aktuellen Stand der Datei herunterladen (als Base64). */
export async function downloadBase64(f: FileRef): Promise<string> {
  const res = await graph<Response>(`/drives/${f.driveId}/items/${f.itemId}/content`, { raw: true });
  const buf = new Uint8Array(await res.arrayBuffer());
  let bin = '';
  for (let i = 0; i < buf.length; i += 0x8000) bin += String.fromCharCode(...buf.subarray(i, i + 0x8000));
  return btoa(bin);
}

const wb = (f: FileRef) => `/drives/${f.driveId}/items/${f.itemId}/workbook`;
const ws = (f: FileRef, sheet: string) => `${wb(f)}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;

/** Werte eines Bereichs lesen, z. B. "F10:I10". */
export async function readRange(f: FileRef, sheet: string, address: string): Promise<unknown[][]> {
  const r = await graph<{ values: unknown[][] }>(`${ws(f, sheet)}/range(address='${address}')?$select=values`);
  return r.values;
}

/** Werte in einen Bereich schreiben (Formatierung bleibt erhalten). */
export async function writeRange(f: FileRef, sheet: string, address: string, values: unknown[][]) {
  await graph(`${ws(f, sheet)}/range(address='${address}')`, { method: 'PATCH', body: JSON.stringify({ values }) });
}

export async function listSheets(f: FileRef): Promise<{ name: string; visibility: string }[]> {
  const r = await graph<{ value: { name: string; visibility: string }[] }>(`${wb(f)}/worksheets?$select=name,visibility`);
  return r.value;
}

export async function addHiddenSheet(f: FileRef, name: string) {
  await graph(`${wb(f)}/worksheets/add`, { method: 'POST', body: JSON.stringify({ name }) });
  await graph(ws(f, name), { method: 'PATCH', body: JSON.stringify({ visibility: 'Hidden' }) });
}

export async function usedRangeValues(f: FileRef, sheet: string): Promise<unknown[][]> {
  const r = await graph<{ values: unknown[][] }>(`${ws(f, sheet)}/usedRange(valuesOnly=true)?$select=values`);
  return r.values ?? [];
}

export async function me(): Promise<{ displayName: string }> {
  return graph('/me?$select=displayName');
}
