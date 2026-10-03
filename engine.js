/* Motore di calcolo "Assenze Collegio" - stessa logica del foglio Excel (fogli Giorni / Riepilogo / Mesi / Elenco).
   Funzioni pure, senza accesso al DOM: usato sia dall'app sia dai test. Le date sono stringhe ISO "AAAA-MM-GG". */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Engine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var MESI = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
  var MESI_AB = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
  var GIORNI = ['lunedì', 'martedì', 'mercoledì', 'giovedì', 'venerdì', 'sabato', 'domenica'];

  /* ---------- date ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function parse(s) { var p = s.split('-'); return { y: +p[0], m: +p[1], d: +p[2] }; }
  function isValidISO(s) {
    if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
    var p = parse(s);
    return p.m >= 1 && p.m <= 12 && p.d >= 1 && p.d <= daysInMonth(p.y, p.m);
  }
  function toDays(s) { var p = parse(s); return Math.round(Date.UTC(p.y, p.m - 1, p.d) / 86400000); }
  function fromDays(n) { var t = new Date(n * 86400000); return iso(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); }
  function dow(s) { return (((toDays(s) + 3) % 7) + 7) % 7; } // lunedì = 0
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function monthStart(s) { var p = parse(s); return iso(p.y, p.m, 1); }
  function endOfMonth(s) { var p = parse(s); return iso(p.y, p.m, daysInMonth(p.y, p.m)); }
  function addMonths(s, k) { var p = parse(s); var t = p.y * 12 + (p.m - 1) + k; return iso(Math.floor(t / 12), (t % 12) + 1, 1); }
  function todayISO() { var d = new Date(); return iso(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function fmtDate(s) { var p = parse(s); return pad(p.d) + '/' + pad(p.m) + '/' + p.y; }
  function fmtLong(s) { var p = parse(s); return GIORNI[dow(s)] + ' ' + p.d + ' ' + MESI[p.m - 1].toLowerCase() + ' ' + p.y; }

  /* ---------- stato di partenza ---------- */
  function defaults() {
    return {
      version: 1,
      settings: { inizio: '2026-10-01', fine: '2027-09-30', soglia: 0.56, oggiOverride: '', baseFrazione: 'validi', allerta: 10 },
      chiusure: [
        { dal: '2026-12-23', al: '2027-01-06', desc: 'Chiusura natalizia (date provvisorie)' },
        { dal: '2027-08-01', al: '2027-08-31', desc: 'Agosto' }
      ],
      festivita: [
        { data: '2026-11-01', desc: 'Ognissanti' }, { data: '2026-12-08', desc: 'Immacolata Concezione' },
        { data: '2026-12-25', desc: 'Natale' }, { data: '2026-12-26', desc: 'Santo Stefano' },
        { data: '2027-01-01', desc: 'Capodanno' }, { data: '2027-01-06', desc: 'Epifania' },
        { data: '2027-03-28', desc: 'Pasqua' }, { data: '2027-03-29', desc: "Lunedì dell'Angelo (Pasquetta)" },
        { data: '2027-04-25', desc: 'Festa della Liberazione' }, { data: '2027-05-01', desc: 'Festa del Lavoro' },
        { data: '2027-06-02', desc: 'Festa della Repubblica' }, { data: '2027-08-15', desc: 'Ferragosto' }
      ],
      marks: {},                       // { 'AAAA-MM-GG': { s: 'A' | 'G' | '', t: millisecondi } }  ('' = presenza / annullata)
      meta: { settingsUpdatedAt: 0, lastExportAt: 0 }
    };
  }

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* ---------- controllo impostazioni (stessi messaggi del foglio Impostazioni) ---------- */
  function check(state) {
    var s = state.settings;
    if (!isValidISO(s.inizio) || !isValidISO(s.fine)) return 'Inserisci una data di inizio e di fine valide';
    if (s.fine <= s.inizio) return "La fine del periodo deve essere dopo l'inizio";
    var first = addMonths(monthStart(s.fine), -11);
    if (s.inizio < first) return 'Il periodo supera i 12 mesi mostrati: sposta inizio o fine';
    if (!(s.soglia > 0 && s.soglia <= 1)) return 'La soglia deve essere tra 0% e 100%';
    for (var i = 0; i < state.chiusure.length; i++) {
      var c = state.chiusure[i];
      if (!!c.dal !== !!c.al) return 'Chiusure: controlla le date Dal/Al';
      if (c.dal && c.al && c.dal > c.al) return 'Chiusure: controlla le date Dal/Al';
    }
    return 'OK';
  }

  /* ---------- calcolo completo ---------- */
  function compute(state, todayStr) {
    var s = state.settings;
    var oggi = s.oggiOverride && isValidISO(s.oggiOverride) ? s.oggiOverride : (todayStr || todayISO());
    var chk = check(state);
    var out = { oggi: oggi, check: chk, ok: chk === 'OK' };
    if (!isValidISO(s.fine)) { out.days = []; out.mesi = []; return out; }

    var first = addMonths(monthStart(s.fine), -11);
    var last = endOfMonth(s.fine);
    var fest = {};
    state.festivita.forEach(function (f) { if (f.data) fest[f.data] = f.desc || 'Festività'; });
    var chius = state.chiusure.filter(function (c) { return c.dal && c.al; });

    var days = [], byDate = {};
    for (var n = toDays(first), n1 = toDays(last); n <= n1; n++) {
      var date = fromDays(n);
      var inPeriodo = date >= s.inizio && date <= s.fine;
      var festivo = Object.prototype.hasOwnProperty.call(fest, date);
      var chiuso = null;
      for (var k = 0; k < chius.length; k++) if (chius[k].dal <= date && date <= chius[k].al) { chiuso = chius[k].desc || 'Chiusura'; break; }
      var valido = inPeriodo && !festivo && !chiuso;
      var m = state.marks[date];
      var stato = m && (m.s === 'A' || m.s === 'G') ? m.s : '';
      var classe = !inPeriodo ? 3 : (festivo || chiuso ? 2 : (dow(date) > 4 ? 1 : 0));
      var day = {
        date: date, dow: dow(date), inPeriodo: inPeriodo, festivo: festivo, chiuso: !!chiuso,
        motivo: festivo ? fest[date] : (chiuso || ''), valido: valido, stato: stato,
        assenza: valido && stato === 'A' ? 1 : 0, giust: valido && stato === 'G' ? 1 : 0,
        residuo: valido && date >= oggi ? 1 : 0, classe: classe
      };
      days.push(day); byDate[date] = day;
    }
    out.days = days; out.byDate = byDate; out.first = first; out.last = last;

    var sum = function (f) { var t = 0; for (var i = 0; i < days.length; i++) t += f(days[i]); return t; };
    var gv = sum(function (d) { return d.valido ? 1 : 0; });
    var round6 = function (x) { return Math.round(x * 1e6) / 1e6; };
    out.giorniCalendario = out.ok ? toDays(s.fine) - toDays(s.inizio) + 1 : 0;
    out.gv = gv;
    out.giorniEsclusi = out.giorniCalendario - gv;
    out.minPres = Math.ceil(round6(s.soglia * gv));
    out.maxAss = gv - out.minPres;
    out.usate = sum(function (d) { return d.assenza; });
    out.giust = sum(function (d) { return d.giust; });
    out.rimanenti = out.maxAss - out.usate;
    out.gvRim = sum(function (d) { return d.residuo; });
    out.gvTrasc = gv - out.gvRim;
    out.assTrasc = sum(function (d) { return d.assenza && d.date < oggi ? 1 : 0; });
    out.presFinora = out.gvTrasc - out.assTrasc;
    out.pctFinora = out.gvTrasc === 0 ? null : out.presFinora / out.gvTrasc;
    out.presFine = gv - out.usate;
    out.pctFine = gv === 0 ? 0 : out.presFine / gv;
    out.margine = out.pctFine - s.soglia;

    /* mesi (foglio Mesi) */
    var mesi = [], sumW = 0;
    for (var i = 0; i < 12; i++) {
      var key = addMonths(first, i), p = parse(key), pref = key.slice(0, 7);
      var md = days.filter(function (d) { return d.date.slice(0, 7) === pref; });
      var validi = 0, ass = 0, giu = 0, res = 0;
      md.forEach(function (d) { validi += d.valido ? 1 : 0; ass += d.assenza; giu += d.giust; res += d.residuo; });
      var peso = 0;
      if (validi > 0) {
        if (s.baseFrazione === 'calendario') {
          var st = [key, oggi, s.inizio].sort().pop();
          var en = [endOfMonth(key), s.fine].sort()[0];
          peso = Math.max(0, toDays(en) - toDays(st) + 1) / daysInMonth(p.y, p.m);
        } else peso = res / validi;
      }
      sumW += peso;
      mesi.push({ key: key, y: p.y, m: p.m, label: MESI[p.m - 1] + ' ' + p.y, labelAb: MESI_AB[p.m - 1] + ' ' + String(p.y).slice(2),
        validi: validi, ass: ass, giu: giu, pres: validi - ass, pct: validi === 0 ? null : (validi - ass) / validi, peso: peso });
    }
    out.mesi = mesi; out.mesiRim = sumW;
    out.mensili = sumW === 0 ? 0 : Math.round(Math.max(0, out.rimanenti) / sumW * 1e4) / 1e4;

    /* stato complessivo (banner del Calendario) */
    out.status = !out.ok ? 'verifica' : (out.rimanenti < 0 ? 'superata' : (out.rimanenti <= s.allerta ? 'attenzione' : 'regola'));

    /* elenco (foglio Elenco) */
    var elenco = [];
    days.forEach(function (d) {
      if (!d.stato) return;
      var nota = d.valido ? (d.stato === 'G' ? 'Conta come presenza' : '')
        : (d.festivo ? 'Festività: non conta' : (d.chiuso ? 'Chiusura: non conta' : 'Fuori periodo: non conta'));
      elenco.push({ n: elenco.length + 1, date: d.date, dow: d.dow, tipo: d.stato === 'A' ? 'Assenza' : 'Giustificata', conta: d.valido, nota: nota });
    });
    out.elenco = elenco;
    return out;
  }

  /* ---------- modifica dei giorni (immutabile) ---------- */
  function setMark(state, date, s, now) {
    var next = clone(state);
    next.marks[date] = { s: s === 'A' || s === 'G' ? s : '', t: now };
    return next;
  }

  /* segna un periodo: solo sui giorni validi; restituisce { state, changed:[{date,prev}] } */
  function paintRange(state, from, to, s, now, result) {
    var next = clone(state), changed = [], target = s === 'A' || s === 'G' ? s : '';
    for (var n = toDays(from), n1 = toDays(to); n <= n1; n++) {
      var date = fromDays(n), day = result.byDate[date];
      if (!day || !day.valido) continue;
      var prev = state.marks[date] ? state.marks[date].s : '';
      if (prev === target) continue;
      changed.push({ date: date, prev: state.marks[date] ? clone(state.marks[date]) : null });
      next.marks[date] = { s: target, t: now };
    }
    return { state: next, changed: changed };
  }

  /* ---------- validazione e unione (backup / sincronizzazione tra dispositivi) ---------- */
  function normalize(obj) {
    if (!obj || typeof obj !== 'object') throw new Error('File non valido');
    var d = defaults(), st = obj.settings || {};
    var out = clone(d);
    if (isValidISO(st.inizio)) out.settings.inizio = st.inizio;
    if (isValidISO(st.fine)) out.settings.fine = st.fine;
    if (typeof st.soglia === 'number' && st.soglia > 0 && st.soglia <= 1) out.settings.soglia = st.soglia;
    if (st.oggiOverride === '' || isValidISO(st.oggiOverride)) out.settings.oggiOverride = st.oggiOverride || '';
    if (st.baseFrazione === 'validi' || st.baseFrazione === 'calendario') out.settings.baseFrazione = st.baseFrazione;
    if (typeof st.allerta === 'number' && st.allerta >= 0) out.settings.allerta = Math.round(st.allerta);
    if (Array.isArray(obj.chiusure)) out.chiusure = obj.chiusure.map(function (c) {
      return { dal: isValidISO(c && c.dal) ? c.dal : '', al: isValidISO(c && c.al) ? c.al : '', desc: String((c && c.desc) || '').slice(0, 80) };
    });
    if (Array.isArray(obj.festivita)) out.festivita = obj.festivita.filter(function (f) { return f && isValidISO(f.data); })
      .map(function (f) { return { data: f.data, desc: String(f.desc || '').slice(0, 80) }; });
    out.marks = {};
    var mk = obj.marks && typeof obj.marks === 'object' ? obj.marks : {};
    Object.keys(mk).forEach(function (k) {
      var v = mk[k];
      if (isValidISO(k) && v && (v.s === 'A' || v.s === 'G' || v.s === '') && typeof v.t === 'number') out.marks[k] = { s: v.s, t: v.t };
    });
    var me = obj.meta || {};
    out.meta = { settingsUpdatedAt: +me.settingsUpdatedAt || 0, lastExportAt: +me.lastExportAt || 0 };
    return out;
  }

  /* unisce due stati: per ogni giorno vince l'ultima modifica; le impostazioni vengono dallo stato più recente */
  function merge(local, incoming) {
    var out = clone(local), rep = { aggiunte: 0, modificate: 0, invariate: 0, impostazioni: 'locali' };
    Object.keys(incoming.marks).forEach(function (k) {
      var a = local.marks[k], b = incoming.marks[k];
      if (!a) { out.marks[k] = clone(b); if (b.s) rep.aggiunte++; return; }
      if (b.t > a.t && b.s !== a.s) { out.marks[k] = clone(b); rep.modificate++; }
      else if (b.t > a.t) out.marks[k] = clone(b);
      else rep.invariate++;
    });
    if (incoming.meta.settingsUpdatedAt > local.meta.settingsUpdatedAt) {
      out.settings = clone(incoming.settings); out.chiusure = clone(incoming.chiusure); out.festivita = clone(incoming.festivita);
      out.meta.settingsUpdatedAt = incoming.meta.settingsUpdatedAt; rep.impostazioni = 'importate';
    }
    return { state: out, report: rep };
  }

  function exportObject(state, now) {
    var o = clone(state);
    o.app = 'assenze-collegio'; o.exportedAt = now;
    return o;
  }

  return {
    MESI: MESI, MESI_AB: MESI_AB, GIORNI: GIORNI,
    pad: pad, iso: iso, parse: parse, isValidISO: isValidISO, toDays: toDays, fromDays: fromDays, dow: dow,
    daysInMonth: daysInMonth, monthStart: monthStart, endOfMonth: endOfMonth, addMonths: addMonths,
    todayISO: todayISO, fmtDate: fmtDate, fmtLong: fmtLong,
    defaults: defaults, clone: clone, check: check, compute: compute,
    setMark: setMark, paintRange: paintRange, normalize: normalize, merge: merge, exportObject: exportObject
  };
});
