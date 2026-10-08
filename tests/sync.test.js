const test = require('node:test'), assert = require('node:assert/strict'), E = require('../engine.js');
const J = x => JSON.parse(JSON.stringify(x));

test('reconcileDays: adotta il piu\' recente, invia i locali piu\' recenti', () => {
  const local = { '2026-10-10': { s: 'A', t: 100 }, '2026-10-11': { s: 'A', t: 300 }, '2026-10-12': { s: 'G', t: 50 } };
  const remote = { '2026-10-10': { s: '', t: 200 }, '2026-10-11': { s: 'G', t: 250 }, '2026-10-13': { s: 'A', t: 10 } };
  const r = E.reconcileDays(local, remote);
  assert.deepEqual(J(r.adopt), { '2026-10-10': { s: '', t: 200 }, '2026-10-13': { s: 'A', t: 10 } });
  assert.deepEqual(r.push.sort(), ['2026-10-11', '2026-10-12']);
});
test('reconcileDays: pari t = nessuna azione; annullamento solo locale non inviato; dati sporchi ignorati', () => {
  const r = E.reconcileDays({ '2026-10-10': { s: 'A', t: 5 }, '2026-10-11': { s: '', t: 9 } }, { '2026-10-10': { s: 'A', t: 5 }, 'x': { s: 'A', t: 1 }, '2026-10-14': { s: 'Z', t: 1 }, '2026-10-15': { s: 'A' } });
  assert.deepEqual(J(r), { adopt: {}, push: [] });
});
test('reconcileDays: convergenza tra due dispositivi con scambi in qualunque ordine', () => {
  let A = {}, B = {}, cloud = {};
  const put = (m, d, s, t) => { m[d] = { s, t }; };
  const sync = (m) => { const r = E.reconcileDays(m, cloud); Object.assign(m, r.adopt); r.push.forEach(d => { if (!cloud[d] || m[d].t >= cloud[d].t) cloud[d] = { ...m[d] }; }); };
  put(A, '2026-10-10', 'A', 100); put(A, '2026-10-11', 'A', 101); sync(A); sync(B);
  put(B, '2026-10-10', '', 200); put(A, '2026-10-12', 'G', 150); put(A, '2026-10-11', '', 120); put(B, '2026-10-11', 'G', 130);
  sync(B); sync(A); sync(B); sync(A);
  assert.deepEqual(A, B); assert.deepEqual(A, cloud);
  assert.equal(A['2026-10-10'].s, ''); assert.equal(A['2026-10-11'].s, 'G'); assert.equal(A['2026-10-12'].s, 'G');
});
test('impostazioni: documento, confronto, applicazione', () => {
  const s = E.defaults(); assert.equal(E.reconcileSettings(s, null).action, 'none');
  s.meta.settingsUpdatedAt = 50; assert.equal(E.reconcileSettings(s, null).action, 'push');
  const doc = E.settingsDoc(s); assert.equal(doc.updatedAt, 50); assert.equal(doc.chiusure.length, 2);
  assert.equal(E.reconcileSettings(s, doc).action, 'none');
  assert.equal(E.reconcileSettings(s, { ...doc, updatedAt: 60 }).action, 'adopt');
  assert.equal(E.reconcileSettings(s, { ...doc, updatedAt: 10 }).action, 'push');
  const other = E.defaults(); other.marks['2026-10-10'] = { s: 'A', t: 1 };
  const d2 = { updatedAt: 99, settings: { ...s.settings, inizio: '2026-10-20', soglia: 0.6 }, chiusure: [{ dal: '2027-01-01', al: '2027-01-03', desc: 'x' }], festivita: [] };
  const r = E.applySettingsDoc(other, d2);
  assert.equal(r.settings.inizio, '2026-10-20'); assert.equal(r.settings.soglia, 0.6); assert.equal(r.chiusure.length, 1); assert.equal(r.meta.settingsUpdatedAt, 99);
  assert.equal(r.marks['2026-10-10'].s, 'A');                                   // i giorni non vengono toccati
  const bad = E.applySettingsDoc(other, { updatedAt: 5, settings: { inizio: 'boh', soglia: 9 }, chiusure: 'no', festivita: null });
  assert.equal(bad.settings.inizio, '2026-10-01'); assert.equal(bad.settings.soglia, 0.56);   // dati sporchi neutralizzati
});
