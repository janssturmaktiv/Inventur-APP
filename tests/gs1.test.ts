import { parseGs1 } from '../src/lib/gs1';
const GS = '\u001d';
const cases: [string, any][] = [
  [']d201040123456789011726123110ABC123' + GS + '21XYZ', { gtin: '04012345678901', expiry: '2026-12-31', batch: 'ABC123', serial: 'XYZ' }],
  ['(01)04012345678901(10)L-4711(17)270600', { gtin: '04012345678901', batch: 'L-4711', expiry: '2027-06-30' }],
  ['0104012345678901' + '10CH99', { gtin: '04012345678901', batch: 'CH99' }],
  ['4012345678901', null],
  ['ABC-123', null],
];
let fail = 0;
for (const [inp, exp] of cases) {
  const r = parseGs1(inp);
  const got = r && { gtin: r.gtin, batch: r.batch, expiry: r.expiry, serial: r.serial };
  const norm = (o: any) => o && JSON.stringify(Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined).sort()));
  const ok = norm(got) === norm(exp);
  if (!ok) fail++;
  console.log(ok ? 'OK  ' : 'FAIL', JSON.stringify(inp), norm(got));
}
process.exit(fail ? 1 : 0);
