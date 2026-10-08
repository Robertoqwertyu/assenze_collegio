const test = require('node:test'), assert = require('node:assert/strict'), E = require('../engine.js');

function scenario(base) {
  let s = E.defaults(); s.settings.oggiOverride = '2026-11-15'; s.settings.baseFrazione = base;
  const m = { '2026-10-10': 'A', '2026-10-12': 'G', '2026-10-20': '', '2026-11-01': 'A', '2026-12-24': 'A',
              '2027-03-28': 'A', '2027-07-20': 'G', '2027-02-14': 'A', '2027-09-10': 'G' };
  Object.keys(m).forEach((d, i) => { s = E.setMark(s, d, m[d], 1000 + i); });
  return E.compute(s, '2026-11-15');
}

test('base pulita = Excel (312 / 175 / 137)', () => {
  const r = E.compute(E.defaults(), '2026-10-03');
  assert.equal(r.gv, 312); assert.equal(r.minPres, 175); assert.equal(r.maxAss, 137);
  assert.equal(r.giorniCalendario, 365); assert.equal(r.giorniEsclusi, 53); assert.equal(r.rimanenti, 137);
  assert.equal(r.mensili, 12.528); assert.equal(r.status, 'regola'); assert.equal(r.check, 'OK');
});
test('scenario con A/G/P, festivi e chiusure = Excel (base giorni validi)', () => {
  const r = scenario('validi');
  assert.equal(r.usate, 2); assert.equal(r.giust, 3); assert.equal(r.rimanenti, 135);
  assert.equal(r.gvRim, 268); assert.equal(r.presFinora, 43); assert.equal(r.mensili, 14.1336);
  assert.equal(r.elenco.length, 8);
  assert.deepEqual(r.elenco.map(e => [e.date, e.tipo, e.conta, e.nota]), [
    ['2026-10-10', 'Assenza', true, ''], ['2026-10-12', 'Giustificata', true, 'Conta come presenza'],
    ['2026-11-01', 'Assenza', false, 'Festività: non conta'], ['2026-12-24', 'Assenza', false, 'Chiusura: non conta'],
    ['2027-02-14', 'Assenza', true, ''], ['2027-03-28', 'Assenza', false, 'Festività: non conta'],
    ['2027-07-20', 'Giustificata', true, 'Conta come presenza'], ['2027-09-10', 'Giustificata', true, 'Conta come presenza']]);
});
test('scenario = Excel (base giorni di calendario)', () => {
  const r = scenario('calendario');
  assert.equal(r.rimanenti, 135); assert.equal(r.mensili, 14.1608);
});
test('P annulla un\'assenza', () => {
  let s = E.defaults(); s = E.setMark(s, '2026-10-10', 'A', 1); s = E.setMark(s, '2026-10-10', 'P', 2);
  assert.equal(E.compute(s, '2026-10-03').usate, 0);
});
test('sforamento soglia e avviso', () => {
  let s = E.defaults(); s.settings.allerta = 10;
  let r = E.compute(s, '2026-10-03'); assert.equal(r.status, 'regola');
  const dates = []; for (let d = '2026-10-01', n = 0; dates.length < 128; d = E.fromDays(E.toDays(d) + 1)) if (E.compute(s,'x').byDate[d].valido) dates.push(d);
  dates.forEach((d, i) => { s = E.setMark(s, d, 'A', i); });
  r = E.compute(s, '2026-10-03'); assert.equal(r.rimanenti, 9); assert.equal(r.status, 'attenzione');
  s = E.setMark(s, E.compute(s,'x').days.find(x => x.valido && !x.stato).date, 'A', 999); s = E.setMark(s, E.compute(s,'x').days.find(x => x.valido && !x.stato).date, 'A', 1000);
  r = E.compute(s, '2026-10-03'); assert.equal(r.rimanenti, 7);
});
test('controlli impostazioni (stessi messaggi dell\'Excel)', () => {
  const t = f => { const s = E.defaults(); f(s); return E.check(s); };
  assert.equal(t(s => { s.settings.fine = '2026-09-01'; }), "La fine del periodo deve essere dopo l'inizio");
  assert.equal(t(s => { s.settings.soglia = 1.5; }), 'La soglia deve essere tra 0% e 100%');
  assert.equal(t(s => { s.chiusure.push({ dal: '2027-03-01', al: '', desc: '' }); }), 'Chiusure: controlla le date Dal/Al');
  assert.equal(t(s => { s.chiusure.push({ dal: '2027-03-10', al: '2027-03-01', desc: '' }); }), 'Chiusure: controlla le date Dal/Al');
  assert.equal(t(s => { s.settings.inizio = '2026-09-01'; }), 'Il periodo supera i 12 mesi mostrati: sposta inizio o fine');
});
test('giorno della settimana e mesi', () => {
  assert.equal(E.dow('2026-10-01'), 3); assert.equal(E.dow('2027-03-28'), 6); assert.equal(E.dow('2026-11-01'), 6);
  assert.equal(E.addMonths('2027-09-01', -11), '2026-10-01'); assert.equal(E.fmtLong('2026-10-10'), 'sabato 10 ottobre 2026');
});
test('data di ingresso diversa: giorni validi cambiano', () => {
  const s = E.defaults(); s.settings.inizio = '2026-10-15';
  const r = E.compute(s, '2026-10-03'); assert.equal(r.gv, 312 - 14); assert.equal(r.minPres, Math.ceil(0.56 * 298));
});
test('paintRange: solo giorni validi', () => {
  const s = E.defaults(), r = E.compute(s, '2026-10-03');
  const p = E.paintRange(s, '2026-12-20', '2026-12-31', 'A', 5, r);
  assert.deepEqual(p.changed.map(c => c.date), ['2026-12-20', '2026-12-21', '2026-12-22']);
});
test('unione: vince l\'ultima modifica per giorno, nulla si perde', () => {
  let a = E.defaults(), b = E.defaults();
  a = E.setMark(a, '2026-10-10', 'A', 100); a = E.setMark(a, '2026-10-11', 'A', 100);
  b = E.setMark(b, '2026-10-10', 'P', 200);                 // annullata dopo sul secondo dispositivo
  b = E.setMark(b, '2026-10-12', 'G', 150);
  const m = E.merge(a, b);
  assert.equal(m.state.marks['2026-10-10'].s, ''); assert.equal(m.state.marks['2026-10-11'].s, 'A'); assert.equal(m.state.marks['2026-10-12'].s, 'G');
  const m2 = E.merge(b, a); assert.deepEqual(m2.state.marks, m.state.marks);  // commutativa
  assert.deepEqual(E.merge(m.state, b).state.marks, m.state.marks);           // idempotente
});
test('unione: impostazioni dal più recente', () => {
  const a = E.defaults(), b = E.defaults(); b.settings.inizio = '2026-10-20'; b.meta.settingsUpdatedAt = 50;
  assert.equal(E.merge(a, b).state.settings.inizio, '2026-10-20');
  assert.equal(E.merge(b, a).state.settings.inizio, '2026-10-20');
});
test('normalize rifiuta dati sporchi', () => {
  assert.throws(() => E.normalize(null));
  const n = E.normalize({ settings: { inizio: 'boh', soglia: 7 }, marks: { '2026-13-45': { s: 'A', t: 1 }, '2026-10-10': { s: 'X', t: 1 }, '2026-10-11': { s: 'A', t: 5 } } });
  assert.equal(n.settings.inizio, '2026-10-01'); assert.equal(n.settings.soglia, 0.56); assert.deepEqual(Object.keys(n.marks), ['2026-10-11']);
});
test('normalize blinda i dati ostili (t non finito, __proto__, soglie assurde, elenchi enormi)', () => {
  const evil = JSON.parse('{"settings":{"allerta":1e999,"inizio":"2026-10-01","fine":"2027-09-30"},"marks":{"2026-10-10":{"s":"A","t":1e999},"__proto__":{"s":"A","t":5},"constructor":{"s":"A","t":5},"2026-10-11":{"s":"<img src=x onerror=alert(1)>","t":1},"2026-10-12":{"s":"A","t":-5},"2026-10-13":{"s":"A","t":7}},"meta":{"settingsUpdatedAt":1e999,"lastExportAt":-3}}');
  evil.chiusure = Array.from({ length: 500 }, () => ({ dal: '2026-12-23', al: '2027-01-06', desc: 'x' }));
  evil.festivita = Array.from({ length: 900 }, () => ({ data: '2026-11-01', desc: 'y'.repeat(500) }));
  const n = E.normalize(evil);
  assert.deepEqual(Object.keys(n.marks), ['2026-10-13']);                 // t=Infinity, t<0, stato non valido e chiavi __proto__/constructor scartati
  assert.equal(n.settings.allerta, 10);                                    // 1e999 non e' una soglia: resta il valore predefinito
  assert.equal(n.meta.settingsUpdatedAt, 0); assert.equal(n.meta.lastExportAt, 0);
  assert.equal(n.chiusure.length, 200); assert.equal(n.festivita.length, 400); assert.equal(n.festivita[0].desc.length, 80);
  assert.equal(({}).s, undefined);                                         // nessun inquinamento di Object.prototype
});
