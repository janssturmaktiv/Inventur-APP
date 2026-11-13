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
