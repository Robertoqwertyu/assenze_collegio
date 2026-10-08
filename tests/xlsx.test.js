const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), cp = require('node:child_process');
const E = require('../engine.js'), X = require('../xlsx.js');
const tpl = new Uint8Array(fs.readFileSync(__dirname + '/../modello.xlsx'));
const NOW = 1790000000000;
function scenario() {
  let s = E.defaults(); s.settings.oggiOverride = '2026-11-15';
  const m = { '2026-10-10': 'A', '2026-10-12': 'G', '2026-11-01': 'A', '2026-12-24': 'A', '2027-03-28': 'A', '2027-07-20': 'G', '2027-02-14': 'A', '2027-09-10': 'G' };
  Object.keys(m).forEach((d, i) => { s = E.setMark(s, d, m[d], 1000 + i); });
  return s;
}
test('zip: scrittura e rilettura identiche, CRC corretto, file valido per unzip', async () => {
  const files = { 'a.txt': new TextEncoder().encode('ciao '.repeat(500)), 'cartella/b.xml': new TextEncoder().encode('<x>è</x>'), 'vuoto': new Uint8Array(0) };
  const z = await X.zip(files, NOW), back = await X.unzip(z);
  assert.deepEqual(Object.keys(back), Object.keys(files));
  for (const k of Object.keys(files)) assert.deepEqual(Buffer.from(back[k]), Buffer.from(files[k]));
  fs.writeFileSync('/tmp/t.zip', z); const out = cp.execSync('unzip -t /tmp/t.zip').toString(); assert.match(out, /No errors detected/);
  assert.ok(z.length < 2500);                                                  // la compressione funziona
});
test('modello: si legge, ha i 6 fogli e la mappa del calendario coincide con il foglio', async () => {
  const f = await X.unzip(tpl); assert.ok(f['xl/workbook.xml']);
  assert.equal(X.statusRef('2026-10-01', '2026-10-01'), 'D17'); assert.equal(X.statusRef('2026-10-10', '2026-10-01'), 'F19'); assert.equal(X.statusRef('2027-09-30', '2026-10-01'), 'D190');                  // giovedì 30 settembre 2027, ultimo blocco
  assert.equal(X.statusRef('2026-11-01', '2026-10-01'), 'G' + (17 + 15));                       // domenica 1 novembre, prima riga del mese
  assert.equal(X.statusRef('2026-09-30', '2026-10-01'), null);
});
test('esporta -> importa: stesse assenze e stesse impostazioni', async () => {
  const s = scenario(); s.chiusure[0].desc = 'Natale & <chiusura> "x"'; s.festivita.push({ data: '2027-05-06', desc: 'San Nicola' });
  const res = E.compute(s, '2026-11-15'), out = await X.build(tpl, s, res, NOW), p = await X.parse(out.bytes);
  const want = {}; Object.keys(s.marks).forEach(d => { if (s.marks[d].s) want[d] = s.marks[d].s; });
  assert.deepEqual(p.marks, want);
  assert.deepEqual(p.settings, { ...s.settings });
  assert.deepEqual(p.chiusure, s.chiusure); assert.deepEqual(p.festivita, s.festivita);
});
test('importa da file dell\'Excel originale vuoto', async () => {
  const p = await X.parse(tpl); assert.deepEqual(p.marks, {}); assert.equal(p.settings.inizio, '2026-10-01'); assert.equal(p.chiusure.length, 2); assert.equal(p.festivita.length, 12);
});
test('periodo diverso e righe oltre il limite del foglio', async () => {
  const s = E.defaults(); s.settings.inizio = '2026-10-20'; s.settings.fine = '2027-10-19'; s.settings.baseFrazione = 'calendario'; s.settings.soglia = 0.6; s.settings.allerta = 7;
  for (let i = 0; i < 31; i++) s.festivita.push({ data: E.fromDays(20700 + i), desc: 'f' + i });
  const res = E.compute(s, '2026-10-03'), out = await X.build(tpl, s, res, NOW);
  assert.equal(out.warnings.length, 1); const p = await X.parse(out.bytes);
  assert.equal(p.settings.fine, '2027-10-19'); assert.equal(p.settings.baseFrazione, 'calendario'); assert.equal(p.settings.soglia, 0.6); assert.equal(p.festivita.length, 30);
});
test('file non Excel o non riconosciuto: errori chiari', async () => {
  await assert.rejects(X.parse(new Uint8Array([1, 2, 3, 4, 5])), /valido/);
  const z = await X.zip({ 'xl/workbook.xml': new TextEncoder().encode('<workbook/>'), 'xl/_rels/workbook.xml.rels': new TextEncoder().encode('<Relationships/>') });
  await assert.rejects(X.parse(z), /non riconosciuto/);
});
test('LibreOffice ricalcola il file esportato e ottiene gli STESSI numeri dell\'app', async () => {
  const s = scenario(), res = E.compute(s, '2026-11-15'), out = await X.build(tpl, s, res, NOW);
  fs.writeFileSync('/tmp/export_lo.xlsx', out.bytes);
  const r = JSON.parse(cp.execSync('python3 /mnt/skills/public/xlsx/scripts/recalc.py /tmp/export_lo.xlsx 120').toString());
  assert.equal(r.status, 'success'); assert.equal(r.total_errors, 0);
  const py = `
from openpyxl import load_workbook
import json
w=load_workbook('/tmp/export_lo.xlsx',data_only=True); r=w['Riepilogo']; c=w['Calendario']; m=w['Mesi']; e=w['Elenco']
print(json.dumps({'validi':r['B9'].value,'min':r['B11'].value,'max':r['B12'].value,'usate':r['B15'].value,'giust':r['B16'].value,'rim':r['B17'].value,'gvrim':r['B18'].value,
 'pres':r['B21'].value,'pctfine':round(r['B24'].value,6),'mesi':round(r['B28'].value,6),'mensili':r['B29'].value,'kpi':[c['A4'].value,c['C4'].value],'stato':c['A6'].value,
 'mesiTot':[m['B17'].value,m['C17'].value,m['D17'].value],'elenco':e['A3'].value,'ott':[m['B5'].value,m['C5'].value]}))`;
  const lo = JSON.parse(cp.execSync('python3 -c "' + py.replace(/"/g, '\\"') + '"').toString());
  assert.equal(lo.validi, res.gv); assert.equal(lo.min, res.minPres); assert.equal(lo.max, res.maxAss);
  assert.equal(lo.usate, res.usate); assert.equal(lo.giust, res.giust); assert.equal(lo.rim, res.rimanenti); assert.equal(lo.gvrim, res.gvRim);
  assert.equal(lo.pres, res.presFinora); assert.equal(lo.pctfine, Math.round(res.pctFine * 1e6) / 1e6); assert.equal(lo.mesi, Math.round(res.mesiRim * 1e6) / 1e6);
  assert.equal(lo.mensili, res.mensili); assert.deepEqual(lo.kpi, [res.rimanenti, res.mensili]);
  assert.deepEqual(lo.mesiTot, [res.gv, res.usate, res.giust]);
  assert.match(lo.elenco, /Segnate: 8/);
});
