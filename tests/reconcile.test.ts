import { reconcile } from '../src/lib/reconcile';
import type { Scan, TargetItem } from '../src/lib/types';
const t = (key: string, extra: Partial<TargetItem> = {}): TargetItem => {
  const [articleNo, side, size] = key.split('|');
  return { id: key, key, articleNo, name: 'X', side, size, expectedQty: 1, ...extra };
};
const s = (id: string, extra: Partial<Scan>): Scan => ({ id, batch: 'L' + id, quantity: 1, raw: '', codeType: 'qr', scannedAt: '', ...extra });
const targets = [t('07780|RECHTS|S'), t('07780|RECHTS|M'), t('07853|LINKS|S', { udiDi: '4049999000017' })];
const scans = [
  s('1', { udiDi: '04260000000011' }), // gelernt -> 07780 S
  s('2', { udiDi: '04260000000011' }), // gleiche GTIN, zweites Stück -> Übermenge auf S
  s('3', { udiDi: '04049999000017' }), // GTIN aus Liste (mit führender 0)
  s('4', { udiDi: '09999' }), // unbekannt
  s('5', { targetKey: '07780|RECHTS|M' }), // manuell gewählt
];
const r = reconcile(targets, scans, { '4260000000011': '07780|RECHTS|S' });
const got = r.rows.map((x) => `${x.item.key}:${x.foundQty}:${x.status}`).join(' ') + ' unknown=' + r.unknown.map((x) => x.id);
const exp = '07780|RECHTS|S:2:vollständig 07780|RECHTS|M:1:vollständig 07853|LINKS|S:1:vollständig unknown=4';
console.log(got === exp ? 'OK' : 'FAIL', got);
process.exit(got === exp ? 0 : 1);
