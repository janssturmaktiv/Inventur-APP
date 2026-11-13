import { GTIN_ZUORDNUNG } from '../src/data/gtinZuordnung';

let fail = 0;
const ok = (c: boolean, m: string) => { if (!c) { fail++; console.error('FAIL', m); } };
const keys = new Set<string>();
for (const [gtin, key] of Object.entries(GTIN_ZUORDNUNG)) {
  const d = gtin.padStart(14, '0').split('').map(Number);
  const sum = d.slice(0, 13).reduce((s, n, i) => s + n * (i % 2 === 0 ? 3 : 1), 0);
  ok((10 - (sum % 10)) % 10 === d[13], `Prüfziffer ${gtin}`);
  ok(!gtin.startsWith('0'), `ohne führende Null ${gtin}`);
  ok(/^\d{5}\|(LINKS|RECHTS)\|[0-9A-Z]+$/.test(key), `Format ${key}`);
  ok(!keys.has(key), `doppelt ${key}`);
  keys.add(key);
}
console.log(fail ? `gtinTable: ${fail} Fehler` : `gtinTable: ${keys.size} Einträge ok`);
if (fail) process.exit(1);

// Eingebaute Tabelle schlägt alte Handzuordnung und falsch gemerkte Zuordnung
import { findTarget } from '../src/lib/reconcile';
import { STANDARD_LISTE } from '../src/data/standardListe';
const hit = findTarget(STANDARD_LISTE, { udiDi: '04029925258897', batch: 'X', targetKey: '07382|LINKS|3' }, { '4029925258897': '07382|LINKS|3' });
if (hit?.key !== '07382|LINKS|2') { console.error('FAIL Vorrang eingebaute Tabelle', hit?.key); process.exit(1); }
// Jede GTIN der Tabelle landet in genau ihrer Zeile der Standardliste
for (const [g, key] of Object.entries(GTIN_ZUORDNUNG)) {
  const t = findTarget(STANDARD_LISTE, { udiDi: '0' + g, batch: 'X' }, {});
  if (STANDARD_LISTE.some((x) => x.key === key) && t?.key !== key) { console.error('FAIL', g, key, t?.key); process.exit(1); }
}
console.log('gtinTable: Vorrang + Zeilenzuordnung ok');
