// Scans per E-Mail an die gemeinsame Liste senden (Power Automate + Office Script tragen sie ein).
import { findTarget } from './reconcile';
import type { Mappings, Scan, TargetItem } from './types';
import { GTIN_ZUORDNUNG } from '../data/gtinZuordnung';
import { normId } from './udi';

export const MAIL_SUBJECT = 'INVENTUR-SCANS';
// Max. Länge des URL-kodierten Textes je E-Mail (sehr lange mailto-Links kürzen manche Mail-Apps).
// Reicht für ca. 100 Scans in einer Mail.
const MAX_ENCODED = 4000;
const enc = (s: string) => encodeURIComponent(s).length;

export type MailBatch = { subject: string; body: string; scanIds: string[] };

const clean = (s: string) => s.replace(/[;\r\n]/g, ',').trim();

export function buildMailBatches(scans: Scan[], targets: TargetItem[], mappings: Mappings, sender: string, finish = false): MailBatch[] {
  const lines: { line: string; id: string }[] = [];
  for (const s of [...scans].reverse()) {
    const t = findTarget(targets, s, mappings);
    if (!t) continue;
    // GTIN nur mitschicken, wenn sie nicht ohnehin in der eingebauten Tabelle steht (hält die Mail kurz)
    const gtin = s.udiDi && !GTIN_ZUORDNUNG[normId(s.udiDi)] ? `;${clean(s.udiDi)}` : '';
    lines.push({ line: `A;${clean(t.key)};${s.quantity};${clean(s.batch)}${gtin}`, id: s.id });
  }
  const head = `Bitte nicht ändern.\n###INVENTUR-V1###\nS;${clean(sender || 'unbekannt')}\n`;
  const tail = '###ENDE###\n';
  const batches: MailBatch[] = [];
  let cur: { line: string; id: string }[] = [];
  const flush = (withFinish: boolean) => {
    const body = head + cur.map((l) => l.line + '\n').join('') + (withFinish ? `F;${clean(sender || '–')}\n` : '') + tail;
    batches.push({ subject: MAIL_SUBJECT, body, scanIds: cur.map((l) => l.id) });
    cur = [];
  };
  for (const l of lines) {
    const size = enc(head) + enc(tail) + cur.reduce((n, x) => n + enc(x.line + '\n'), 0) + enc(l.line + '\n') + 20;
    if (cur.length && size > MAX_ENCODED) flush(false);
    cur.push(l);
  }
  if (cur.length || finish) flush(finish);
  const n = batches.length;
  if (n > 1) batches.forEach((b, i) => (b.subject = `${MAIL_SUBJECT} (${i + 1}/${n})`));
  return batches;
}

export function mailtoUrl(to: string, b: MailBatch) {
  return `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(b.subject)}&body=${encodeURIComponent(b.body)}`;
}
