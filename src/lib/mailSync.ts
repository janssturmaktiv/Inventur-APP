// Scans per E-Mail an die gemeinsame Liste senden (Power Automate + Office Script tragen sie ein).
import { findTarget } from './reconcile';
import type { Mappings, Scan, TargetItem } from './types';

export const MAIL_SUBJECT = 'INVENTUR-SCANS';
const MAX_BODY = 1500; // Zeichen je E-Mail (manche Mail-Apps kürzen sehr lange Links)

export type MailBatch = { subject: string; body: string; scanIds: string[] };

const clean = (s: string) => s.replace(/[;\r\n]/g, ',').trim();

export function buildMailBatches(scans: Scan[], targets: TargetItem[], mappings: Mappings, sender: string, finish = false): MailBatch[] {
  const lines: { line: string; id: string }[] = [];
  for (const s of [...scans].reverse()) {
    const t = findTarget(targets, s, mappings);
    if (!t) continue;
    lines.push({ line: `A;${clean(t.key)};${s.quantity};${clean(s.batch)};${clean(s.udiDi ?? '')}`, id: s.id });
  }
  const head = `Bitte nicht ändern – wird automatisch in die Bestandsliste eingetragen.\n\n###INVENTUR-V1###\nS;${clean(sender || 'unbekannt')}\n`;
  const tail = '###ENDE###\n';
  const batches: MailBatch[] = [];
  let cur: { line: string; id: string }[] = [];
  const flush = (withFinish: boolean) => {
    const body = head + cur.map((l) => l.line + '\n').join('') + (withFinish ? `F;${clean(sender || '–')}\n` : '') + tail;
    batches.push({ subject: MAIL_SUBJECT, body, scanIds: cur.map((l) => l.id) });
    cur = [];
  };
  for (const l of lines) {
    const size = head.length + tail.length + cur.reduce((n, x) => n + x.line.length + 1, 0) + l.line.length;
    if (cur.length && size > MAX_BODY) flush(false);
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
