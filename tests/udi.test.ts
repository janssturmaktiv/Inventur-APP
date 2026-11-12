import { parseUdi } from '../src/lib/udi';
const GS = '\u001d';
const cases: [string, Record<string, string | undefined>][] = [
  [']Q30104012345678901172612311' + '0ABC123' + GS + '21XYZ', { issuer: 'GS1', udiDi: '04012345678901', batch: 'ABC123', expiry: '2026-12-31', serial: 'XYZ' }],
  ['(01)04012345678901(17)270600(10)L-4711', { issuer: 'GS1', udiDi: '04012345678901', batch: 'L-4711', expiry: '2027-06-30' }],
  ['https://id.gs1.org/01/04012345678901/10/LOT77/21/S1?17=280131', { issuer: 'GS1', udiDi: '04012345678901', batch: 'LOT77', serial: 'S1', expiry: '2028-01-31' }],
  ['+A99912345/$$52001510X3', { issuer: 'HIBC', udiDi: 'A99912345', productCode: '1234', batch: '10X', expiry: '2020-01-15' }],
  ['+H123ABC01234567890/$$3271231LOT55F', { issuer: 'HIBC', udiDi: 'H123ABC01234567890', productCode: 'ABC0123456789', batch: 'LOT55', expiry: '2027-12-31' }],
  ['+H123ABC1/$LOT9Z', { issuer: 'HIBC', udiDi: 'H123ABC1', productCode: 'ABC', batch: 'LOT9' }],
  ['+H123ABC1Q', { issuer: 'HIBC', udiDi: 'H123ABC1', productCode: 'ABC' }],
  ['HELLO', { issuer: 'unbekannt' }],
];
let fail = 0;
const norm = (o: any) => JSON.stringify(Object.entries(o).filter(([, v]) => v !== undefined).sort());
for (const [inp, exp] of cases) {
  const got = parseUdi(inp);
  const ok = norm(got) === norm(exp);
  if (!ok) fail++;
  console.log(ok ? 'OK  ' : 'FAIL', JSON.stringify(inp), norm(got));
}
process.exit(fail ? 1 : 0);
