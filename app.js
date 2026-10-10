/* StayTrack (assenze del collegio) - interfaccia. Il calcolo è in engine.js (stessa logica del foglio Excel). */
(function () {
  'use strict';
  var E = window.Engine;
  var KEY = 'assenze-collegio:v1', KEY_PREV = 'assenze-collegio:prev', KEY_UI = 'assenze-collegio:ui';
  var KEY_UID = 'assenze-collegio:uid', KEY_BAD = 'assenze-collegio:illeggibile';
  var VERSION = '2.6';
  var DAY = 86400000;

  var state, res, brush = 'A', undoStack = [], tab = 'calendario', calDirty = true;
  var fb = null, user = null, stopSub = null, known = {}, knownSet = null, pushedDay = {}, pushedSet = 0, firstDays = true;
  var net = { fromCache: true, pending: 0, setPending: 0, error: '', ok: false };
  var installEvt = null, lastToday = E.todayISO(), firstCalendarRender = true, toastTimer = null, calScroll = 0;

  /* ---------- utilità DOM ---------- */
  function $(s, r) { return (r || document).querySelector(s); }
  function append(el, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { append(el, x); }); return; }
    el.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k.slice(0, 2) === 'on') el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? '' : v);
    });
    for (var i = 2; i < arguments.length; i++) append(el, arguments[i]);
    return el;
  }
  function nf(x, d) { return x.toFixed(d).replace('.', ','); }
  function pc(x) { return x == null ? '-' : nf(x * 100, 1) + '%'; }

  /* ---------- memoria ---------- */
  var loadProblem = false, storageBroken = false;
  function load() {
    var raw = null;
    try { raw = localStorage.getItem(KEY); }
    catch (e) { storageBroken = true; return E.defaults(); }       // memoria del browser bloccata (es. navigazione privata): non sono dati danneggiati
    try {
      if (raw) return E.normalize(JSON.parse(raw));
    } catch (e) {
      /* dati locali danneggiati: NON vanno sovrascritti in silenzio, se ne conserva una copia */
      loadProblem = true;
      try { if (raw) localStorage.setItem(KEY_BAD, raw); } catch (e2) { /* ignora */ }
      console.warn('Dati non leggibili', e);
    }
    return E.defaults();
  }
  function save(noSync) {
    try { localStorage.setItem(KEY, JSON.stringify(state)); storageBroken = false; }
    catch (e) { storageBroken = true; toast('Attenzione: impossibile salvare su questo dispositivo'); }
    if (!noSync) noteChange();
  }
  function keepPrev() { try { var raw = localStorage.getItem(KEY); if (raw) localStorage.setItem(KEY_PREV, raw); } catch (e) { /* ignora */ } }
  var onboarded = false;
  function loadUI() { try { var u = JSON.parse(localStorage.getItem(KEY_UI) || '{}'); if (u.brush) brush = u.brush; onboarded = !!u.onboarded; } catch (e) { /* ignora */ } }
  function saveUI() { try { localStorage.setItem(KEY_UI, JSON.stringify({ brush: brush, onboarded: onboarded })); } catch (e) { /* ignora */ } }
  function recalc() { res = E.compute(state, E.todayISO()); }
  function hasMarks() { return Object.keys(state.marks).some(function (k) { return state.marks[k].s; }); }
  function needsBackup() { return hasMarks() && Date.now() - (state.meta.lastExportAt || 0) > 14 * DAY; }

  /* ---------- messaggi ---------- */
  function hideToast() { var t = $('#toast'); t.hidden = true; }
  function toast(msg, action) {
    var t = $('#toast');
    t.textContent = '';
    t.appendChild(document.createTextNode(msg));
    if (action) t.appendChild(h('button', { type: 'button', text: action.label, onclick: function () { hideToast(); action.fn(); } }));
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(hideToast, action ? 5500 : 2600);
  }

  /* ---------- navigazione ---------- */
  var ICONS = {
    calendario: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
    riepilogo: '<path d="M5 21V11M12 21V4M19 21v-7"/>',
    mesi: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16M15 4v16"/>',
    elenco: '<path d="M9 6h12M9 12h12M9 18h12"/><circle cx="4.5" cy="6" r="1"/><circle cx="4.5" cy="12" r="1"/><circle cx="4.5" cy="18" r="1"/>',
    impostazioni: '<path d="M4 7h10M18 7h2M4 17h2M10 17h10"/><circle cx="16" cy="7" r="2"/><circle cx="8" cy="17" r="2"/>'
  };
  var TABS = [['calendario', 'Calendario'], ['riepilogo', 'Riepilogo'], ['mesi', 'Mesi'], ['elenco', 'Elenco'], ['impostazioni', 'Impostazioni']];

  /* logo della barra laterale su PC: stesso disegno dell'icona dell'app (casa con griglia 3x3, navy su tessera bianca) */
  var LOGO = '<svg viewBox="0 0 512 512" aria-hidden="true"><rect x="2" y="2" width="508" height="508" rx="120" fill="#fff"/><g transform="translate(256 256) scale(1.375) translate(-158.5 -159.5)"><path d="M30 135L159 31.2L213.5 74.8V45.5H251.5V105.2L288 135" fill="none" stroke="#13243e" stroke-width="15" stroke-linecap="round" stroke-linejoin="round"/><path d="M65.5 143L158.5 68.6L251.5 143V269.5Q251.5 287.5 233.5 287.5H83.5Q65.5 287.5 65.5 269.5Z" fill="none" stroke="#13243e" stroke-width="15" stroke-linecap="round" stroke-linejoin="round"/><path d="M65.5 138.5H251.5" fill="none" stroke="#13243e" stroke-width="15" stroke-linecap="round" stroke-linejoin="round"/><path d="M127.5 138.5V287.5M189.5 138.5V287.5M65.5 188.5H251.5M65.5 238H251.5" fill="none" stroke="#13243e" stroke-width="15" stroke-linecap="round" stroke-linejoin="round"/></g></svg>';

  function buildNav() {
    var nav = $('#tabs');
    nav.textContent = '';
    var brand = h('div', { class: 'brand' });
    brand.innerHTML = LOGO;
    brand.appendChild(h('div', null, h('b', { text: 'StayTrack' }), h('small', { text: 'Assenze collegio' })));
    nav.appendChild(brand);
    var list = h('div', { class: 'navlist' });
    TABS.forEach(function (t) {
      var b = h('button', { type: 'button', 'data-tab': t[0], onclick: function () { setTab(t[0]); } });
      var ico = h('span', { class: 'ico' });
      ico.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICONS[t[0]] + '</svg>';
      b.appendChild(ico);
      b.appendChild(h('span', { class: 'lbl', text: t[1] }));
      list.appendChild(b);
    });
    nav.appendChild(list);
    /* piede della barra laterale (solo PC): stato del cloud e versione */
    nav.appendChild(h('div', { class: 'navfoot' },
      h('div', { class: 'cloudline off', id: 'cloudst3', role: 'button', tabindex: '0', onclick: gotoCloud, onkeydown: function (e) { if (e.key === 'Enter') gotoCloud(); } }),
      h('small', { text: 'Versione ' + VERSION })));
    updateNav();
    renderCloudStatus();
  }
  function updateNav() {
    document.querySelectorAll('#tabs button').forEach(function (b) {
      if (b.dataset.tab === tab) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
      var dot = b.querySelector('.dot');
      if (b.dataset.tab === 'impostazioni') {
        if (needsBackup() && !dot) b.appendChild(h('i', { class: 'dot' }));
        else if (!needsBackup() && dot) dot.remove();
      }
    });
  }
  function setTab(name) {
    if (tab === 'calendario') calScroll = window.scrollY;
    tab = name;
    TABS.forEach(function (t) { $('#v-' + t[0]).hidden = t[0] !== name; });
    renderTab(name);
    updateNav();
    window.scrollTo(0, name === 'calendario' ? calScroll : 0);
  }
  function renderTab(name) {
    ({ calendario: renderCalendario, riepilogo: renderRiepilogo, mesi: renderMesi, elenco: renderElenco, impostazioni: renderImpostazioni })[name]();
  }
  function refresh() { calDirty = true; renderTab(tab); updateNav(); }

  /* ---------- CALENDARIO ---------- */
  function statusLabel(st) {
    return { regola: 'In regola', attenzione: 'Attenzione: poche assenze rimaste', superata: 'Soglia superata', verifica: 'Verifica impostazioni' }[st];
  }
  function statusText() { return statusLabel(res.status); }
  function renderKpis() {
    if (!$('#k-rim')) return;
    var okd = res.days && res.days.length;
    $('#k-rim').textContent = okd ? res.rimanenti : '–';
    $('#k-rim').parentNode.classList.toggle('neg', okd && res.rimanenti < 0);
    $('#k-men').textContent = okd ? nf(res.mensili, 2) : '–';
    $('#k-pct').textContent = okd ? pc(res.pctFine) : '–';
    var st = $('#status');
    st.className = 'status' + (res.status && res.status !== 'regola' ? ' ' + res.status : '');
    st.textContent = okd ? statusText() : 'Verifica impostazioni';
    $('#info').textContent = okd
      ? 'Assenze usate ' + res.usate + ' su ' + res.maxAss + '  ·  Giustificate ' + res.giust + '  ·  Presenze finora ' + res.presFinora
        + (res.mesiRim > 0 && res.mesiRim < 1 ? '  ·  Meno di 1 mese rimasto: conta il totale delle assenze rimanenti' : '')
        + (storageBroken ? '  ·  ⚠ Dati NON salvati su questo dispositivo: esporta un backup o accedi con Google' : '')
      : res.check;
    (res.mesi || []).forEach(function (m) {
      var el = $('#mc-' + m.key.slice(0, 7));
      if (el) el.textContent = m.validi === 0 ? 'chiuso' : (m.ass || m.giu ? 'assenze ' + m.ass + ' · giust. ' + m.giu : m.validi + ' giorni validi');
    });
  }

  function paintDay(b, day) {
    var c = 'day';
    if (day.classe === 3) c += ' out';
    else if (day.classe === 2) c += ' fest';
    else { if (day.classe === 1) c += ' we'; if (day.stato) c += ' st-' + day.stato; }
    if (day.date === res.oggi) c += ' today';
    if (day.stato && !day.valido) c += ' nocount';
    b.className = c;
    b.querySelector('.s').setAttribute('data-s', day.stato || '');
    var what = day.festivo ? 'festività ' + day.motivo : (day.chiuso ? 'residenza chiusa' : (day.classe === 3 ? 'fuori periodo'
      : (day.stato === 'A' ? 'assenza' : (day.stato === 'G' ? 'assenza giustificata' : 'presente'))));
    b.setAttribute('aria-label', E.fmtLong(day.date) + ', ' + what);
  }
  function dayEl(day) {
    var b = h('button', { class: 'day', type: 'button', tabindex: '-1', 'data-date': day.date },
      h('span', { class: 'd', text: String(parseInt(day.date.slice(8), 10)) }), h('span', { class: 's' }));   // la lettera A/G è generata dal CSS (data-s): il testo del pulsante resta il solo numero del giorno
    paintDay(b, day);
    return b;
  }
  function updateDay(date) {
    var b = $('.day[data-date="' + date + '"]');
    if (b && res.byDate[date]) paintDay(b, res.byDate[date]);
  }

  function reason(day) {
    if (day.festivo) return 'Festività: ' + day.motivo + ' - non conta';
    if (day.chiuso) return 'Residenza chiusa (' + day.motivo + ') - non conta';
    return 'Giorno fuori dal periodo - non conta';
  }

  function onDay(date) {
    var day = res.byDate[date];
    if (!day) return;
    if (!day.valido) { toast(reason(day)); return; }
    var cur = day.stato, target = brush === 'P' ? '' : (cur === brush ? '' : brush);
    if (target === cur) { toast('Già presente'); return; }
    applyMark(date, target);
  }
  function applyMark(date, target) {
    var prev = state.marks[date] ? state.marks[date].s : '';
    state = E.setMark(state, date, target, Date.now());
    undoStack.push([{ date: date, prev: prev }]);
    if (undoStack.length > 50) undoStack.shift();
    afterChange([date]);
    if (navigator.vibrate && (!navigator.userActivation || navigator.userActivation.hasBeenActive)) { try { navigator.vibrate(8); } catch (e) { /* ignora */ } }
    var label = target === 'A' ? 'Assenza segnata' : (target === 'G' ? 'Giustificata segnata' : 'Segnato presente');
    toast(label + ': ' + E.fmtDate(date), { label: 'ANNULLA', fn: undo });
  }
  function afterChange(dates) {
    save(); recalc();
    dates.forEach(updateDay);
    renderKpis(); updateUndoBtn(); updateNav();
  }
  function undo() {
    var batch = undoStack.pop();
    if (!batch) return;
    var now = Date.now();
    batch.forEach(function (c) { state = E.setMark(state, c.date, c.prev, now); });
    afterChange(batch.map(function (c) { return c.date; }));
    toast('Modifica annullata');
  }
  function updateUndoBtn() { var b = $('#btn-undo'); if (b) b.disabled = undoStack.length === 0; }

  function setBrush(b) {
    brush = b; saveUI();
    document.querySelectorAll('.brush').forEach(function (el) { el.setAttribute('aria-checked', el.dataset.b === b ? 'true' : 'false'); });
  }

  function scrollToMonth(key, smooth) {
    var el = $('#m-' + key.slice(0, 7));
    if (el) el.scrollIntoView({ behavior: smooth ? 'smooth' : 'auto', block: 'start' });
  }
  function currentMonthKey() { return res.oggi.slice(0, 7); }

  function renderCalendario() {
    var v = $('#v-calendario');
    if (!calDirty && v.firstChild) { renderKpis(); return; }
    var y = window.scrollY;
    v.textContent = '';
    v.appendChild(h('h1', { class: 'sr-only', text: 'StayTrack: calendario' }));
    var sticky = h('div', { class: 'sticky', id: 'sticky' },
      h('div', { class: 'topline' },
        h('div', { class: 'status', id: 'status' }),
        h('div', { class: 'cloudst off', id: 'cloudst', role: 'button', tabindex: '0', onclick: gotoCloud, onkeydown: function (e) { if (e.key === 'Enter') gotoCloud(); } })),
      h('div', { class: 'kpis' },
        h('div', { class: 'kpi k1' }, h('small', { text: 'Assenze rimanenti' }), h('b', { id: 'k-rim' })),
        h('div', { class: 'kpi k2' }, h('small', { text: 'Mensili rimanenti' }), h('b', { id: 'k-men' })),
        h('div', { class: 'kpi k3' }, h('small', { text: 'Presenza a fine anno' }), h('b', { id: 'k-pct' }))),
      h('div', { class: 'toolbar' },
        h('div', { class: 'brushes', role: 'radiogroup', 'aria-label': 'Cosa segnare toccando un giorno' },
          brushBtn('A', 'Assenza'), brushBtn('G', 'Giustificata'), brushBtn('P', 'Presente')),
        h('button', { class: 'tbtn', type: 'button', text: 'Periodo…', onclick: openPeriodo }),
        h('button', { class: 'tbtn', id: 'btn-undo', type: 'button', 'aria-label': 'Annulla ultima modifica', text: '↶', onclick: undo })));
    v.appendChild(sticky);
    v.appendChild(h('div', { class: 'info', id: 'info' }));      // fuori dalla barra fissa: occupa meno spazio mentre si scorre

    if (!res.days || !res.days.length) {
      v.appendChild(h('div', { class: 'warnbox', text: 'Imposta le date del periodo in Impostazioni: ' + res.check }));
      calDirty = false; renderKpis(); return;
    }
    var chips = h('div', { class: 'chips' });
    chips.appendChild(h('button', { class: 'chip now', type: 'button', text: 'Oggi', onclick: function () { scrollToMonth(currentMonthKey(), true); } }));
    res.mesi.forEach(function (m) { chips.appendChild(h('button', { class: 'chip', type: 'button', text: m.labelAb, onclick: function () { scrollToMonth(m.key, true); } })); });
    v.appendChild(chips);

    var months = h('div', { class: 'months' });
    months.addEventListener('click', function (e) { var b = e.target.closest('.day'); if (b) { setRoving(b); onDay(b.dataset.date); } });
    months.addEventListener('keydown', onDayKey);
    res.mesi.forEach(function (m) {
      var grid = h('div', { class: 'grid' });
      'LMMGVSD'.split('').forEach(function (c, i) { grid.appendChild(h('div', { class: 'wd' + (i > 4 ? ' we' : ''), text: c })); });
      for (var i = 0, off = E.dow(m.key); i < off; i++) grid.appendChild(h('div'));
      for (var d = 1, dim = E.daysInMonth(m.y, m.m); d <= dim; d++) grid.appendChild(dayEl(res.byDate[E.iso(m.y, m.m, d)]));
      months.appendChild(h('div', { class: 'month', id: 'm-' + m.key.slice(0, 7) },
        h('div', { class: 'mh' + (m.key.slice(0, 7) === currentMonthKey() ? ' now' : '') }, h('b', { text: m.label }), h('span', { id: 'mc-' + m.key.slice(0, 7) })),
        grid));
    });
    v.appendChild(months);
    v.appendChild(h('div', { class: 'legend' },
      h('span', null, h('i', { style: 'background:var(--present)' }), 'presente (automatico)'),
      h('span', null, h('i', { style: 'background:var(--abs)' }), 'A assenza'),
      h('span', null, h('i', { style: 'background:var(--just)' }), 'G giustificata'),
      h('span', null, h('i', { style: 'background:var(--fest-d)' }), 'festa / chiuso')));
    setBrush(brush);
    setRoving($('.day[data-date="' + (res.byDate[res.oggi] ? res.oggi : res.days[0].date) + '"]'));
    calDirty = false;
    renderKpis(); updateUndoBtn(); renderCloudStatus();
    syncStickyHeight();
    if (firstCalendarRender) { firstCalendarRender = false; if (res.mesi.some(function (m) { return m.key.slice(0, 7) === currentMonthKey(); })) scrollToMonth(currentMonthKey(), false); }
    else window.scrollTo(0, y);
  }
  /* un solo giorno è raggiungibile con Tab; con le frecce ci si sposta tra i giorni, Invio/Spazio segna */
  function setRoving(el) {
    if (!el) return;
    var cur = $('.day[tabindex="0"]'); if (cur && cur !== el) cur.setAttribute('tabindex', '-1');
    el.setAttribute('tabindex', '0');
  }
  function onDayKey(e) {
    var b = e.target.closest && e.target.closest('.day'); if (!b) return;
    var d = b.dataset.date, n = E.toDays(d), t = null;
    if (e.key === 'ArrowLeft') t = E.fromDays(n - 1); else if (e.key === 'ArrowRight') t = E.fromDays(n + 1);
    else if (e.key === 'ArrowUp') t = E.fromDays(n - 7); else if (e.key === 'ArrowDown') t = E.fromDays(n + 7);
    else if (e.key === 'PageUp' || e.key === 'PageDown') {
      var p = E.parse(d), ms = E.addMonths(E.monthStart(d), e.key === 'PageUp' ? -1 : 1), pm = E.parse(ms);
      t = E.iso(pm.y, pm.m, Math.min(p.d, E.daysInMonth(pm.y, pm.m)));
    } else if (e.key === 'Home') t = E.monthStart(d); else if (e.key === 'End') t = E.endOfMonth(d);
    if (!t) return;
    var el = $('.day[data-date="' + t + '"]');
    if (el) { e.preventDefault(); setRoving(el); el.focus(); }
  }
  function brushBtn(b, label) {
    return h('button', { class: 'brush', type: 'button', role: 'radio', 'data-b': b, 'aria-checked': 'false', onclick: function () { setBrush(b); } },
      b, h('small', { text: label }));
  }
  var stickyH = -1;
  function syncStickyHeight() {
    var s = $('#sticky');
    if (!s) return;
    var hgt = s.offsetHeight;
    if (hgt === stickyH) return;           // invariata: evita di far ricalcolare lo stile a tutti i 365 giorni
    stickyH = hgt;
    document.documentElement.style.setProperty('--sticky-h', hgt + 'px');
  }

  /* segna un periodo */
  function openPeriodo() {
    var dlg = $('#dlg-periodo'), s = state.settings;
    var from = h('input', { type: 'date', id: 'p-from', value: res.oggi >= s.inizio && res.oggi <= s.fine ? res.oggi : s.inizio, min: s.inizio, max: s.fine });
    var to = h('input', { type: 'date', id: 'p-to', value: from.value, min: s.inizio, max: s.fine });
    var kind = 'A';
    var prev = h('div', { class: 'info', id: 'p-prev' });
    var aft = h('div', { class: 'after', id: 'p-after', 'aria-live': 'polite' });
    function radio(v, lab) {
      var i = h('input', { type: 'radio', name: 'pk', value: v, onchange: function () { kind = v; upd(); } });
      if (v === 'A') i.checked = true;
      return h('label', null, i, h('span', { text: lab }));
    }
    function upd() {
      if (!from.value || !to.value || to.value < from.value) { prev.textContent = 'Controlla le date'; aft.textContent = ''; return; }
      var r = E.paintRange(state, from.value, to.value, kind, 0, res);
      prev.textContent = r.changed.length === 1 ? '1 giorno valido verrà modificato' : r.changed.length + ' giorni validi verranno modificati';
      /* "cosa succede se": stato dopo l'applicazione (utile per pianificare ponti e viaggi) */
      if (r.changed.length) {
        var a = E.compute(r.state, E.todayISO()), dl = a.rimanenti - res.rimanenti;
        aft.className = 'after ' + a.status;
        aft.textContent = 'Dopo: ' + a.rimanenti + ' assenze rimanenti (' + (dl > 0 ? '+' : dl < 0 ? '−' : '±') + Math.abs(dl) + ') · ' + nf(a.mensili, 2) + ' al mese · ' + statusLabel(a.status);
      } else { aft.className = 'after'; aft.textContent = ''; }
      go.disabled = r.changed.length === 0;
    }
    var go = h('button', { class: 'btn', type: 'button', text: 'Applica', onclick: function () {
      var r = E.paintRange(state, from.value, to.value, kind, Date.now(), res);
      if (!r.changed.length) return;
      state = r.state;
      undoStack.push(r.changed.map(function (c) { return { date: c.date, prev: c.prev ? c.prev.s : '' }; }));
      afterChange(r.changed.map(function (c) { return c.date; }));
      dlg.close();
      toast(r.changed.length === 1 ? '1 giorno aggiornato' : r.changed.length + ' giorni aggiornati', { label: 'ANNULLA', fn: undo });
    } });
    from.addEventListener('change', function () { if (to.value < from.value) to.value = from.value; upd(); });
    to.addEventListener('change', upd);
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-periodo-t', text: 'Segna un periodo' }));
    dlg.appendChild(h('div', { class: 'db' },
      h('div', { class: 'field' }, h('label', { for: 'p-from', text: 'Dal' }), from),
      h('div', { class: 'field' }, h('label', { for: 'p-to', text: 'Al (incluso)' }), to),
      h('div', { class: 'radios' }, radio('A', 'A Assenza'), radio('G', 'G Giust.'), radio('P', 'P Presente')),
      prev, aft,
      h('div', { class: 'note', text: 'Vengono toccati solo i giorni validi: festività e chiusure restano escluse.' })));
    dlg.appendChild(h('div', { class: 'df' }, h('button', { class: 'btn ghost', type: 'button', text: 'Chiudi', onclick: function () { dlg.close(); } }), go));
    upd();
    dlg.showModal();
  }

  /* ---------- RIEPILOGO ---------- */
  function row(label, value, note, cls) {
    return h('div', { class: 'row' + (cls && cls.indexOf('big') >= 0 ? ' big' : '') },
      h('div', { class: 'l' }, label, note ? h('small', { text: note }) : null),
      h('div', { class: 'v' + (cls && cls.indexOf('neg') >= 0 ? ' neg' : ''), text: String(value) }));
  }
  function rows(list) { var c = h('div', { class: 'card' }); list.forEach(function (r) { c.appendChild(row.apply(null, r)); }); return c; }

  function summaryRows() {
    var s = state.settings;
    return {
      periodo: [
        ['Inizio periodo', E.fmtDate(s.inizio)], ['Fine periodo', E.fmtDate(s.fine)], ['Data usata nei calcoli', E.fmtDate(res.oggi), s.oggiOverride ? 'Data di riferimento impostata' : 'Oggi'],
        ['Giorni di calendario del periodo', res.giorniCalendario], ['Giorni esclusi (festività e chiusure)', res.giorniEsclusi, 'Unione senza doppi conteggi'],
        ['Giorni validi (base del ' + Math.round(s.soglia * 100) + '%)', res.gv, 'Weekend inclusi'], ['Soglia minima di presenza', nf(s.soglia * 100, 0) + '%'],
        ['Presenze minime richieste', res.minPres, 'Arrotondato per eccesso'], ['Assenze massime consentite', res.maxAss]],
      stato: [
        ['Assenze usate (non giustificate)', res.usate], ['Assenze giustificate', res.giust, 'Valgono come presenza'],
        ['Assenze rimanenti totali', res.rimanenti, 'Negativo = soglia superata', 'big' + (res.rimanenti < 0 ? ' neg' : '')],
        ['Giorni validi rimanenti (da oggi incluso)', res.gvRim], ['Giorni validi trascorsi', res.gvTrasc], ['Assenze nei giorni trascorsi', res.assTrasc],
        ['Presenze finora', res.presFinora], ['% presenza finora', pc(res.pctFinora), 'Sui giorni validi trascorsi'],
        ['Presenze a fine periodo', res.presFine, 'Se da oggi sempre presente'], ['% presenza a fine periodo', pc(res.pctFine), 'Se da oggi sempre presente'],
        ['Margine sulla soglia', (res.margine >= 0 ? '+' : '') + nf(res.margine * 100, 1) + ' punti', null, res.margine < 0 ? 'neg' : '']],
      proiezione: [
        ['Mesi rimanenti (con frazioni)', nf(res.mesiRim, 2), 'Frazione del mese corrente + mesi interi successivi'],
        ['Assenze mensili rimanenti', nf(res.mensili, 2), 'Assenze rimanenti ÷ mesi rimanenti', 'big'],
        ['Base della frazione di mese', s.baseFrazione === 'calendario' ? 'Giorni di calendario' : 'Giorni validi']]
    };
  }
  /* barra del budget: assenze usate sul massimo consentito, con il marcatore del ritmo uniforme (quante ne avresti usate a oggi spalmandole in modo regolare) */
  function budgetCard() {
    var max = res.maxAss, used = res.usate, over = res.rimanenti < 0;
    var pct = function (x) { return max > 0 ? Math.max(0, Math.min(100, x / max * 100)) : 0; };
    var pace = res.gv > 0 ? max * res.gvTrasc / res.gv : 0;
    var label = over ? 'Soglia superata di ' + (-res.rimanenti) + ': assenze usate ' + used + ' su ' + max
      : 'Assenze usate ' + used + ' su ' + max + ', rimanenti ' + res.rimanenti + '; ritmo uniforme a oggi ' + nf(pace, 1);
    var bar = h('div', { class: 'bar', role: 'img', 'aria-label': label },
      used > 0 ? h('i', { class: 'u' + (pct(used) >= 100 ? ' full' : ''), style: 'width:' + pct(used).toFixed(2) + '%' }) : null,
      max > 0 ? h('b', { class: 'pace', style: 'left:' + pct(pace).toFixed(2) + '%' }) : null);
    return h('div', { class: 'card budget' + (over ? ' neg' : ''), id: 'budget' },
      h('div', { class: 'bh' }, h('span', { class: 'bt', text: over ? 'Assenze oltre la soglia' : 'Assenze disponibili' }),
        h('span', { class: 'bn' }, String(over ? -res.rimanenti : res.rimanenti), h('small', { text: over ? ' in più' : ' su ' + max }))),
      bar,
      h('div', { class: 'bl' },
        h('span', null, h('i', { class: 'k-u' }), 'Usate ' + used),
        h('span', null, h('i', { class: 'k-j' }), 'Giustificate ' + res.giust + ' (non contano)'),
        h('span', null, h('i', { class: 'k-p' }), 'Ritmo uniforme a oggi ' + nf(pace, 1))));
  }
  function renderRiepilogo() {
    var v = $('#v-riepilogo'); v.textContent = '';
    v.classList.toggle('two', res.ok);       // su schermi larghi: due colonne
    v.appendChild(h('h1', { text: 'Riepilogo' }));
    if (!res.ok) { v.appendChild(h('div', { class: 'check bad', text: res.check })); return; }
    var r = summaryRows();
    v.appendChild(budgetCard());
    v.appendChild(h('section', { class: 'sec sec-p' }, h('h2', { text: 'Periodo e base di calcolo' }), rows(r.periodo)));
    v.appendChild(h('section', { class: 'sec sec-s' }, h('h2', { text: 'Stato attuale' }), rows(r.stato)));
    v.appendChild(h('section', { class: 'sec sec-m' }, h('h2', { text: 'Proiezione mensile' }), rows(r.proiezione)));
    v.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn', type: 'button', text: 'Esporta in Excel', onclick: function () { exportExcel(false); } }), h('button', { class: 'btn ghost', type: 'button', text: 'Stampa / salva PDF', onclick: printReport })));
  }

  /* ---------- MESI ---------- */
  function chart(mesi) {
    var W = 360, H = 170, base = 140, left = 22, mx = Math.max(3, Math.max.apply(null, mesi.map(function (m) { return Math.max(m.ass, m.giu); })));
    var bw = (W - left - 6) / mesi.length, svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Assenze e giustificate per mese">';
    for (var t = 0; t <= 3; t++) {
      var val = Math.round(mx * t / 3), y = base - (base - 14) * t / 3;
      svg += '<line class="gl" x1="' + left + '" x2="' + W + '" y1="' + y + '" y2="' + y + '"/><text x="' + (left - 4) + '" y="' + (y + 3) + '" text-anchor="end">' + val + '</text>';
    }
    mesi.forEach(function (m, i) {
      var x = left + i * bw + bw * 0.12, w = bw * 0.38, ha = (base - 14) * m.ass / mx, hg = (base - 14) * m.giu / mx;
      svg += '<rect class="ba" x="' + x.toFixed(1) + '" y="' + (base - ha).toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + ha.toFixed(1) + '" rx="2.5"/>';
      svg += '<rect class="bg" x="' + (x + w + 1).toFixed(1) + '" y="' + (base - hg).toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + hg.toFixed(1) + '" rx="2.5"/>';
      if (m.ass) svg += '<text class="vl" x="' + (x + w / 2).toFixed(1) + '" y="' + (base - ha - 3).toFixed(1) + '" text-anchor="middle">' + m.ass + '</text>';
      if (m.giu) svg += '<text class="vl" x="' + (x + w + 1 + w / 2).toFixed(1) + '" y="' + (base - hg - 3).toFixed(1) + '" text-anchor="middle">' + m.giu + '</text>';
      svg += '<text x="' + (x + w).toFixed(1) + '" y="' + (base + 12) + '" text-anchor="middle">' + E.MESI_AB[m.m - 1] + '</text>';
    });
    svg += '<rect class="ba" x="' + left + '" y="' + (H - 11) + '" width="8" height="8" rx="2.5"/><text class="lg" x="' + (left + 12) + '" y="' + (H - 3) + '">Assenze</text>';
    svg += '<rect class="bg" x="' + (left + 66) + '" y="' + (H - 11) + '" width="8" height="8" rx="2.5"/><text class="lg" x="' + (left + 78) + '" y="' + (H - 3) + '">Giustificate</text>';
    return svg + '</svg>';
  }
  function mesiTable() {
    var tb = h('tbody'), tot = { v: 0, a: 0, g: 0 };
    res.mesi.forEach(function (m) {
      tot.v += m.validi; tot.a += m.ass; tot.g += m.giu;
      tb.appendChild(h('tr', null, h('td', { text: E.MESI_AB[m.m - 1] + ' ' + m.y }), h('td', { text: m.validi }), h('td', { class: m.ass ? 'red' : '', text: m.ass }),
        h('td', { text: m.giu }), h('td', { text: m.pres }), h('td', { text: pc(m.pct) }), h('td', { text: nf(m.peso, 2) })));
    });
    return h('table', { class: 't' },
      h('thead', null, h('tr', null, ['Mese', 'Validi', 'Ass.', 'Giust.', 'Pres.', '%', 'Peso'].map(function (x) { return h('th', { text: x }); }))), tb,
      h('tfoot', null, h('tr', null, h('td', { text: 'Totale' }), h('td', { text: tot.v }), h('td', { text: tot.a }), h('td', { text: tot.g }),
        h('td', { text: tot.v - tot.a }), h('td', { text: tot.v ? pc((tot.v - tot.a) / tot.v) : '-' }), h('td', { text: nf(res.mesiRim, 2) }))));
  }
  function renderMesi() {
    var v = $('#v-mesi'); v.textContent = '';
    v.classList.add('two');                  // su schermi larghi: tabella e grafico affiancati
    v.appendChild(h('h1', { text: 'Dettaglio mensile' }));
    v.appendChild(h('p', { class: 'lead', text: 'Presenze = giorni validi − assenze (le giustificate valgono come presenza). Peso = parte del mese ancora da vivere, per la proiezione mensile.' }));
    if (!res.mesi.length) return;
    var c = h('div', { class: 'card' }, mesiTable());
    v.appendChild(c);
    var ch = h('div', { class: 'card chart' }); ch.innerHTML = chart(res.mesi); v.appendChild(ch);
  }

  /* ---------- ELENCO ---------- */
  function renderElenco() {
    var v = $('#v-elenco'); v.textContent = '';
    v.appendChild(h('h1', { text: 'Elenco delle assenze segnate' }));
    v.appendChild(h('p', { class: 'lead', text: 'In ordine di data: utile per confrontare le assenze con il registro del badge elettronico.' }));
    if (!res.elenco) return;
    var c = h('div', { class: 'card' });
    if (!res.elenco.length) c.appendChild(h('div', { class: 'empty', text: 'Nessuna assenza segnata: ogni giorno è presenza.' }));
    res.elenco.forEach(function (e) {
      var li = h('div', { class: 'li' + (e.conta ? '' : ' no') },
        h('div', { class: 'n', text: e.n }),
        h('div', { class: 'm' }, h('b', { text: E.fmtLong(e.date) }), e.nota ? h('small', { text: e.nota }) : null),
        h('span', { class: 'tag ' + (e.tipo === 'Assenza' ? 'A' : 'G'), text: e.tipo === 'Assenza' ? 'A' : 'G' }));
      c.appendChild(li);
    });
    v.appendChild(c);
    v.appendChild(h('p', { class: 'lead', style: 'margin-top:10px', text: 'Segnate: ' + res.elenco.length + '  ·  Contano come assenza: ' + res.usate + '  ·  Giustificate valide: ' + res.giust }));
  }

  /* ---------- IMPOSTAZIONI ---------- */
  function touchSettings() { state.meta.settingsUpdatedAt = Date.now(); save(); recalc(); calDirty = true; refreshCheck(); updateNav(); }
  function refreshCheck() {
    var c = $('#chk'); if (!c) return;
    c.className = 'check' + (res.ok ? '' : ' bad');
    c.textContent = res.ok ? 'Impostazioni OK' : res.check;
  }
  var fieldSeq = 0;
  function field(label, hint, input) {
    if (!input.id) input.id = 'f' + (++fieldSeq);
    return h('div', { class: 'field' }, h('label', { for: input.id }, label, hint ? h('small', { text: hint }) : null), input);
  }
  function dateInput(get, set, aria) {
    return h('input', { type: 'date', value: get() || '', 'aria-label': aria, onchange: function (e) { set(e.target.value || ''); touchSettings(); } });
  }
  function renderImpostazioni() {
    var y = window.scrollY, v = $('#v-impostazioni'), s = state.settings;
    v.textContent = '';
    v.appendChild(h('h1', { text: 'Impostazioni' }));
    v.appendChild(h('div', { class: 'check', id: 'chk' }));
    cloudCard(v);
    var par = h('div', { class: 'card' });
    par.appendChild(field('Data di ingresso', 'Inizio del periodo di conteggio', dateInput(function () { return s.inizio; }, function (x) { s.inizio = x; }, 'Data di ingresso')));
    par.appendChild(field('Fine periodo', null, dateInput(function () { return s.fine; }, function (x) { s.fine = x; }, 'Fine periodo')));
    par.appendChild(field('Soglia minima di presenza (%)', 'Regolamento: 56% su base annuale',
      h('input', { type: 'number', inputmode: 'decimal', min: 1, max: 100, step: 'any', value: String(Math.round(s.soglia * 10000) / 100),
        onchange: function (e) { var n = parseFloat(String(e.target.value).replace(',', '.')); s.soglia = isNaN(n) ? 0 : n / 100; touchSettings(); } })));
    par.appendChild(field('Avviso a assenze rimanenti ≤', "Compare 'ATTENZIONE' nel Calendario",
      h('input', { type: 'number', inputmode: 'numeric', min: 0, step: 1, value: String(s.allerta),
        onchange: function (e) { var n = parseInt(e.target.value, 10); s.allerta = isNaN(n) || n < 0 ? 0 : n; touchSettings(); } })));
    var base = h('select', { onchange: function (e) { s.baseFrazione = e.target.value; touchSettings(); } },
      h('option', { value: 'validi', text: 'Giorni validi' }), h('option', { value: 'calendario', text: 'Giorni di calendario' }));
    base.value = s.baseFrazione;
    par.appendChild(field('Base della frazione di mese', "Come il vecchio Excel: 'Giorni di calendario'", base));
    var ov = h('input', { type: 'date', value: s.oggiOverride || '', onchange: function (e) { s.oggiOverride = e.target.value || ''; touchSettings(); } });
    par.appendChild(field('Data di riferimento', 'Facoltativa, per simulare un altro giorno. Vuota = oggi', ov));
    v.appendChild(h('h2', { text: 'Parametri' })); v.appendChild(par);

    v.appendChild(h('h2', { text: 'Chiusure della residenza (escluse dal calcolo)' }));
    var cc = h('div', { class: 'card' });
    state.chiusure.forEach(function (c, i) {
      cc.appendChild(h('div', { class: 'lrow' },
        dateInput(function () { return c.dal; }, function (x) { c.dal = x; }, 'Chiusura ' + (i + 1) + ': dal'), dateInput(function () { return c.al; }, function (x) { c.al = x; }, 'Chiusura ' + (i + 1) + ': al'),
        h('button', { class: 'x', type: 'button', 'aria-label': 'Elimina chiusura', text: '✕', onclick: function () { state.chiusure.splice(i, 1); touchSettings(); renderImpostazioni(); } }),
        h('input', { class: 'desc', type: 'text', maxlength: 80, placeholder: 'Descrizione', 'aria-label': 'Chiusura ' + (i + 1) + ': descrizione', value: c.desc || '', onchange: function (e) { c.desc = e.target.value; touchSettings(); } })));
    });
    cc.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn ghost add', type: 'button', text: 'Aggiungi chiusura', onclick: function () { state.chiusure.push({ dal: '', al: '', desc: '' }); renderImpostazioni(); } })));
    v.appendChild(cc);

    v.appendChild(h('h2', { text: 'Festività (escluse dal calcolo)' }));
    var fc = h('div', { class: 'card' });
    state.festivita.forEach(function (f, i) {
      fc.appendChild(h('div', { class: 'lrow f' },
        dateInput(function () { return f.data; }, function (x) { f.data = x; }, 'Festività ' + (i + 1) + ': data'),
        h('input', { class: 'desc', type: 'text', maxlength: 80, placeholder: 'Descrizione', 'aria-label': 'Festività ' + (i + 1) + ': descrizione', value: f.desc || '', onchange: function (e) { f.desc = e.target.value; touchSettings(); } }),
        h('button', { class: 'x', type: 'button', 'aria-label': 'Elimina festività', text: '✕', onclick: function () { state.festivita.splice(i, 1); touchSettings(); renderImpostazioni(); } })));
    });
    fc.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn ghost add', type: 'button', text: 'Aggiungi festività', onclick: function () { state.festivita.push({ data: '', desc: '' }); renderImpostazioni(); } })));
    fc.appendChild(h('div', { class: 'note', text: 'Pasqua e Pasquetta incluse; San Nicola (6 maggio) no. Le date della chiusura natalizia sono provvisorie.' }));
    v.appendChild(fc);

    /* dati */
    v.appendChild(h('h2', { text: 'Backup manuale (copia di sicurezza)' }));
    if (needsBackup()) v.appendChild(h('div', { class: 'warnbox', text: 'Da più di 14 giorni non fai un backup: esporta un file e conservalo (Drive, email…).' }));
    var last = state.meta.lastExportAt ? new Date(state.meta.lastExportAt).toLocaleDateString('it-IT') : 'mai';
    var dc = h('div', { class: 'card' });
    var canShare = !!(navigator.canShare && typeof File !== 'undefined');
    dc.appendChild(h('div', { class: 'btns' },
      h('button', { class: 'btn', type: 'button', text: 'Esporta in Excel', onclick: function () { exportExcel(false); } }),
      canShare ? h('button', { class: 'btn ghost', type: 'button', text: 'Condividi Excel', onclick: function () { exportExcel(true); } }) : null,
      h('button', { class: 'btn ghost', type: 'button', text: 'Backup JSON', onclick: function () { exportFile(false); } }),
      h('button', { class: 'btn ghost', type: 'button', text: 'Importa Excel/JSON…', onclick: function () { $('#file-in').click(); } }),
      h('input', { type: 'file', id: 'file-in', accept: '.xlsx,.json,application/json,' + MIME_XLSX, hidden: true, onchange: importFile })));
    dc.appendChild(h('div', { class: 'btns', style: 'padding-top:0' },
      h('button', { class: 'btn ghost', type: 'button', text: 'Copia codice', onclick: copyCode }),
      h('button', { class: 'btn ghost', type: 'button', text: 'Incolla codice…', onclick: function () { openCodice(); } })));
    dc.appendChild(h('div', { class: 'note', text: 'Ultimo backup: ' + last + ". L'Excel è il foglio completo con formule: puoi continuare a usarlo in Excel e reimportarlo qui (aggiungendo le assenze o sostituendo tutto). Il JSON è una copia tecnica; importando i dati si uniscono (vince la modifica più recente)." }));
    v.appendChild(dc);

    v.appendChild(h('h2', { text: 'App' }));
    var ac = h('div', { class: 'card' });
    ac.appendChild(h('div', { class: 'row' }, h('div', { class: 'l' }, 'Memoria protetta', h('small', { id: 'persist-s', text: '…' })), installEvt ? h('button', { class: 'btn', type: 'button', text: 'Installa app', onclick: doInstall }) : null));
    ac.appendChild(h('div', { class: 'row' }, h('div', { class: 'l' }, 'Versione', h('small', { text: 'Funziona anche offline dopo il primo caricamento' })), h('div', { class: 'v', text: VERSION })));
    ac.appendChild(h('div', { class: 'btns' },
      h('button', { class: 'btn ghost', type: 'button', text: 'Ripristina stato precedente', onclick: restorePrev }),
      h('button', { class: 'btn danger', type: 'button', text: 'Cancella tutti i dati', onclick: resetAll })));
    ac.appendChild(h('div', { class: 'note', text: "Installazione: su Android, menu ⋮ di Chrome › Installa app; su PC, icona di installazione nella barra degli indirizzi di Chrome/Edge." }));
    v.appendChild(ac);
    refreshCheck(); renderCloudStatus();
    if (navigator.storage && navigator.storage.persisted) navigator.storage.persisted().then(function (p) {
      var el = $('#persist-s'); if (el) el.textContent = p ? 'Sì: il browser non cancellerà i dati da solo' : 'Non ancora garantita: fai ogni tanto un backup';
    });
    window.scrollTo(0, y);
  }

  /* ---------- backup / importazione ---------- */
  var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  function fileName(ext) { return 'staytrack-' + E.todayISO() + (ext || '.json'); }
  /* consegna un file: condivisione di sistema (telefono) oppure download */
  function deliver(data, name, mime, share, title) {
    var blob = new Blob([data], { type: mime });
    if (share && navigator.canShare && typeof File !== 'undefined') {
      var f = new File([blob], name, { type: mime });
      if (navigator.canShare({ files: [f] })) { navigator.share({ files: [f], title: title }).catch(function () { /* annullato */ }); return 'shared'; }
    }
    var link = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(link); link.click();
    setTimeout(function () { URL.revokeObjectURL(link.href); link.remove(); }, 1500);
    return 'saved';
  }
  function exportFile(share) {
    state.meta.lastExportAt = Date.now(); save(true);
    var how = deliver(JSON.stringify(E.exportObject(state, Date.now()), null, 1), fileName('.json'), 'application/json', share, 'Backup StayTrack');
    if (how === 'saved') toast('Backup salvato: ' + fileName('.json'));
    updateNav(); renderImpostazioni();
  }
  /* xlsx.js (16 KB) si carica solo alla prima esportazione/importazione Excel: l'avvio dell'app resta leggero.
     Il Service Worker lo tiene già in cache, quindi funziona anche offline. */
  var xlsxLoad = null;
  function loadXlsx() {
    if (window.Xlsx) return Promise.resolve(window.Xlsx);
    if (xlsxLoad) return xlsxLoad;
    xlsxLoad = new Promise(function (ok, ko) {
      var s = document.createElement('script');
      s.src = 'xlsx.js';
      s.onload = function () { if (window.Xlsx) ok(window.Xlsx); else { xlsxLoad = null; ko(new Error('Modulo Excel non valido: ricarica l\'app')); } };
      s.onerror = function () { xlsxLoad = null; ko(new Error('Modulo Excel non disponibile: apri l\'app una volta con la rete')); };
      document.head.appendChild(s);
    });
    return xlsxLoad;
  }
  /* Excel: il foglio "Assenze Collegio" completo (formule, tendine, grafico) con le assenze e le impostazioni già inserite */
  function exportExcel(share) {
    if (!res.ok) { toast(res.check); return; }
    if (typeof DecompressionStream === 'undefined') { toast("Questo browser non supporta l'esportazione Excel: aggiornalo"); return; }
    toast('Preparo il file Excel…');
    loadXlsx().then(function () { return fetch('modello.xlsx'); }).then(function (r) { if (!r.ok) throw new Error('modello'); return r.arrayBuffer(); })
      .then(function (buf) { return Xlsx.build(new Uint8Array(buf), state, res, Date.now()); })
      .then(function (out) {
        state.meta.lastExportAt = Date.now(); save(true);
        var how = deliver(out.bytes, fileName('.xlsx'), MIME_XLSX, share, 'StayTrack (Excel)');
        toast((how === 'saved' ? 'Excel salvato: ' + fileName('.xlsx') : 'Condivisione del file Excel…') + (out.warnings.length ? ' · ' + out.warnings.join('; ') : ''));
        updateNav(); if (tab === 'impostazioni') renderImpostazioni();
      }).catch(function () { toast('Esportazione Excel non riuscita (serve aver aperto l\'app almeno una volta con la rete)'); });
  }
  function importFile(e) {
    var f = e.target.files && e.target.files[0]; e.target.value = '';
    if (!f) return;
    f.arrayBuffer().then(function (buf) {
      var u = new Uint8Array(buf);
      if (u[0] === 0x50 && u[1] === 0x4B) previewImportXlsx(u);               // "PK": file Excel (zip)
      else previewImport(new TextDecoder().decode(u));
    }).catch(function () { toast('File non leggibile'); });
  }
  function toCode() { return 'ASSENZE1:' + btoa(unescape(encodeURIComponent(JSON.stringify(E.exportObject(state, Date.now()))))); }
  function fromCode(code) {
    var m = /^\s*ASSENZE1:([A-Za-z0-9+/=\s]+)$/.exec(code);
    if (!m) throw new Error('Codice non valido');
    return decodeURIComponent(escape(atob(m[1].replace(/\s+/g, ''))));
  }
  function copyCode() {
    state.meta.lastExportAt = Date.now(); save(true); updateNav();
    var code = toCode();
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(code).then(function () { toast('Codice copiato: incollalo sull\'altro dispositivo'); renderImpostazioni(); }, function () { openCodice(code); });
    else openCodice(code);
  }
  function openCodice(code, title, note) {
    var dlg = $('#dlg-codice'), ta = h('textarea', { 'aria-label': 'Codice di sincronizzazione', spellcheck: 'false', placeholder: 'Incolla qui il codice ASSENZE1:…' });
    if (code) ta.value = code;
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-codice-t', text: title || (code ? 'Codice di sincronizzazione' : 'Incolla il codice') }));
    dlg.appendChild(h('div', { class: 'db' }, ta, h('div', { class: 'note', text: note || (code ? 'Copia tutto il testo e incollalo sull\'altro dispositivo (Impostazioni › Incolla codice).' : 'Il codice si genera sull\'altro dispositivo con "Copia codice".') })));
    dlg.appendChild(h('div', { class: 'df' }, h('button', { class: 'btn ghost', type: 'button', text: 'Chiudi', onclick: function () { dlg.close(); } }),
      code ? null : h('button', { class: 'btn', type: 'button', text: 'Unisci', onclick: function () { try { var t = fromCode(ta.value); dlg.close(); previewImport(t); } catch (err) { toast('Codice non valido'); } } })));
    dlg.showModal(); if (code) { ta.focus(); ta.select(); }
  }
  function previewImportXlsx(bytes) {
    if (typeof DecompressionStream === 'undefined') { toast("Questo browser non supporta l'importazione Excel: aggiornalo"); return; }
    loadXlsx().then(function () { return Xlsx.parse(bytes); }).then(function (p) {
      var na = 0, ng = 0; Object.keys(p.marks).forEach(function (d) { if (p.marks[d] === 'A') na++; else ng++; });
      var dlg = $('#dlg-codice'), line = function (l, v) { return h('div', { class: 'row' }, h('div', { class: 'l', text: l }), h('div', { class: 'v', text: String(v) })); };
      dlg.textContent = '';
      dlg.appendChild(h('div', { class: 'dh', id: 'dlg-codice-t', text: 'Importare il file Excel?' }));
      dlg.appendChild(h('div', { class: 'db' },
        line('Periodo', E.fmtDate(p.settings.inizio) + ' - ' + E.fmtDate(p.settings.fine)), line('Assenze (A)', na), line('Giustificate (G)', ng),
        line('Chiusure · festività', p.chiusure.length + ' · ' + p.festivita.length),
        p.warnings.length ? h('div', { class: 'warnbox', text: p.warnings.join('; ') }) : null,
        h('div', { class: 'note', text: 'Aggiungi: le assenze del file si sommano a quelle presenti. Sostituisci tutto: assenze e impostazioni diventano quelle del file (le assenze attuali che il file non ha vengono annullate). Lo stato attuale resta ripristinabile da Impostazioni.' })));
      dlg.appendChild(h('div', { class: 'df', style: 'flex-wrap:wrap' },
        h('button', { class: 'btn ghost', type: 'button', text: 'Annulla', onclick: function () { dlg.close(); } }),
        h('button', { class: 'btn ghost', type: 'button', text: 'Aggiungi le assenze', onclick: function () { dlg.close(); applyXlsx(p, false); } }),
        h('button', { class: 'btn', type: 'button', text: 'Sostituisci tutto', onclick: function () { dlg.close(); applyXlsx(p, true); } })));
      dlg.showModal();
    }).catch(function (e) { toast(e && e.message ? e.message : 'File Excel non valido'); });
  }
  function applyXlsx(p, replace) {
    keepPrev();
    var now = Date.now(), n = 0, next = E.clone(state);
    if (replace) {
      var norm = E.normalize({ settings: p.settings, chiusure: p.chiusure, festivita: p.festivita, marks: {}, meta: {} });
      next.settings = norm.settings; next.chiusure = norm.chiusure; next.festivita = norm.festivita; next.meta.settingsUpdatedAt = now;
      var first = E.addMonths(E.monthStart(p.settings.fine), -11), last = E.endOfMonth(p.settings.fine);
      Object.keys(next.marks).forEach(function (d) {              // annulla le assenze attuali (nella finestra del file) che il file non ha
        if (d >= first && d <= last && next.marks[d].s && !p.marks[d]) { next.marks[d] = { s: '', t: now }; n++; }
      });
    }
    Object.keys(p.marks).forEach(function (d) {
      var cur = next.marks[d];
      if (!cur || cur.s !== p.marks[d]) { next.marks[d] = { s: p.marks[d], t: now }; n++; }
    });
    state = next; undoStack = []; save(); recalc(); refresh();
    toast('Importato da Excel: ' + n + (n === 1 ? ' giorno aggiornato' : ' giorni aggiornati'));
  }
  function previewImport(text) {
    var incoming;
    try { incoming = E.normalize(JSON.parse(text)); } catch (err) { toast('File o codice non valido'); return; }
    var m = E.merge(state, incoming), r = m.report, dlg = $('#dlg-codice');
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-codice-t', text: 'Unire i dati?' }));
    dlg.appendChild(h('div', { class: 'db' },
      h('div', { class: 'row' }, h('div', { class: 'l', text: 'Assenze aggiunte' }), h('div', { class: 'v', text: r.aggiunte })),
      h('div', { class: 'row' }, h('div', { class: 'l', text: 'Giorni modificati' }), h('div', { class: 'v', text: r.modificate })),
      h('div', { class: 'row' }, h('div', { class: 'l', text: 'Impostazioni' }), h('div', { class: 'v', text: r.impostazioni })),
      h('div', { class: 'note', text: "Per ogni giorno vince la modifica più recente. Lo stato attuale resta ripristinabile da Impostazioni." })));
    dlg.appendChild(h('div', { class: 'df' }, h('button', { class: 'btn ghost', type: 'button', text: 'Annulla', onclick: function () { dlg.close(); } }),
      h('button', { class: 'btn', type: 'button', text: 'Unisci', onclick: function () {
        keepPrev(); state = m.state; save(); recalc(); undoStack = []; dlg.close(); refresh(); toast('Dati uniti');
      } })));
    dlg.showModal();
  }
  function restorePrev() {
    var raw; try { raw = localStorage.getItem(KEY_PREV); } catch (e) { /* ignora */ }
    if (!raw) { toast('Nessuno stato precedente disponibile'); return; }
    if (!confirm('Ripristinare lo stato precedente all\'ultima importazione o cancellazione?')) return;
    var cur = localStorage.getItem(KEY);
    try { state = E.normalize(JSON.parse(raw)); } catch (e) { toast('Stato precedente non leggibile'); return; }
    if (cur) localStorage.setItem(KEY_PREV, cur);
    save(); recalc(); undoStack = []; refresh(); toast('Stato precedente ripristinato');
  }
  function resetAll() {
    if (!confirm('Cancellare TUTTE le assenze e le impostazioni di questo dispositivo?')) return;
    if (!confirm('Sicuro? Se non hai un backup i dati andranno persi.')) return;
    keepPrev(); state = E.defaults(); save(); recalc(); undoStack = []; refresh(); toast('Dati cancellati (ripristinabili da Impostazioni)');
  }
  function doInstall() { if (installEvt) { installEvt.prompt(); installEvt = null; renderImpostazioni(); } }


  /* ---------- ACCOUNT E SINCRONIZZAZIONE (Firebase: accesso Google + Firestore) ---------- */
  function cloudText() {
    if (!fb) return { cls: 'off', t: 'Cloud in caricamento…' };
    if (!fb.configured) return { cls: 'off', t: 'Sincronizzazione non attivata · tocca per i dettagli' };
    if (fb.failed) return { cls: 'off', t: 'Cloud non disponibile (serve rete la prima volta)' };
    if (!user) return { cls: 'off', t: 'Accedi con Google per sincronizzare' };
    if (net.error) return { cls: 'err', t: '' + net.error };
    var pend = net.pending + (net.setPending ? 1 : 0);
    if (!navigator.onLine) return { cls: 'wait', t: 'Offline · ' + (pend ? pend + (pend === 1 ? ' modifica da inviare' : ' modifiche da inviare') : 'dati salvati sul dispositivo') };
    if (pend) return { cls: 'wait', t: 'Sincronizzo… (' + pend + ')' };
    if (net.fromCache) return { cls: 'wait', t: 'Connessione al cloud…' };
    return { cls: 'ok', t: 'Sincronizzato · ' + user.email };
  }
  function renderCloudStatus() {
    var c = cloudText();
    [['#cloudst', 'cloudst'], ['#cloudst2', 'cloudline'], ['#cloudst3', 'cloudline']].forEach(function (p) {
      var el = $(p[0]); if (!el) return;
      el.className = p[1] + ' ' + c.cls;
      el.textContent = c.t;
      el.title = c.t;
    });
  }
  function gotoCloud() { setTab('impostazioni'); var el = $('#cloud-card'); if (el) el.scrollIntoView({ block: 'start' }); }

  function initFirebase() {
    if (!window.FB || !window.FB.ready) { window.addEventListener('fb-ready', initFirebase, { once: true }); return; }
    fb = window.FB;
    renderCloudStatus();
    if (tab === 'impostazioni') renderImpostazioni();
    if (fb.configured && !fb.failed && fb.onAuth) fb.onAuth(onAuth);
    if (fb.redirectCheck) fb.redirectCheck.then(function (r) {
      if (!r) return;
      toast(r.lost ? "Accesso non completato: il browser ha bloccato il ritorno da Google. Riprova da Chrome (non in finestra privata) o consenti i popup"
        : authError({ code: r.error }));
    });
  }
  function onAuth(u) {
    stopSync();
    user = u ? { uid: u.uid, email: u.email || '', name: u.displayName || '', photo: u.photoURL || '' } : null;
    renderCloudStatus(); if (tab === 'impostazioni') renderImpostazioni();
    if (!user) { maybeOnboard(); return; }
    var last = null; try { last = localStorage.getItem(KEY_UID); } catch (e) { /* ignora */ }
    if (last && last !== user.uid && hasMarks()) { askAccountSwitch(); return; }
    startSync();
  }
  function startSync() {
    try { localStorage.setItem(KEY_UID, user.uid); } catch (e) { /* ignora */ }
    known = {}; knownSet = null; pushedDay = {}; pushedSet = 0; firstDays = true;
    net = { fromCache: true, pending: 0, setPending: 0, error: '', ok: false };
    try { stopSub = fb.subscribe(user.uid, { days: onDays, settings: onSettings, error: onNetError }); }
    catch (e) { stopSub = null; onNetError(e); return; }
    renderCloudStatus();
  }
  function stopSync() {
    if (stopSub) { stopSub(); stopSub = null; }
    known = {}; knownSet = null; net = { fromCache: true, pending: 0, setPending: 0, error: '', ok: false };
  }

  /* arriva l'elenco dei giorni dal cloud (o dalla cache offline di Firestore): per ogni giorno vince l'ultima modifica */
  function onDays(map, meta) {
    known = map; net.fromCache = meta.fromCache; net.pending = meta.pending; net.error = ''; net.ok = true;
    var r = E.reconcileDays(state.marks, map), adopted = Object.keys(r.adopt);
    if (adopted.length) {
      var hadMarks = hasMarks();
      var visible = adopted.some(function (d) { return (state.marks[d] ? state.marks[d].s : '') !== r.adopt[d].s; });
      adopted.forEach(function (d) { state.marks[d] = r.adopt[d]; });
      undoStack = []; save(true); recalc(); calDirty = true; renderTab(tab); updateNav(); updateUndoBtn();
      if (visible) toast(firstDays && !hadMarks ? 'Dati ripristinati dal cloud' : "Aggiornato dall'altro dispositivo");
    }
    firstDays = false;
    pushDays(r.push);
    renderCloudStatus();
  }
  function onSettings(doc, meta) {
    knownSet = doc; net.setPending = meta.pending ? 1 : 0;
    var a = E.reconcileSettings(state, doc);
    if (a.action === 'adopt') { state = E.applySettingsDoc(state, doc); save(true); recalc(); calDirty = true; renderTab(tab); updateNav(); }
    else if (a.action === 'push') pushSettings();
    renderCloudStatus();
  }
  function onNetError(err) {
    var c = err && err.code;
    net.error = c === 'permission-denied' ? 'Permessi negati: pubblica le Security Rules (firestore.rules)' : 'Errore di sincronizzazione' + (c ? ' (' + c + ')' : '');
    renderCloudStatus(); if (tab === 'impostazioni') renderImpostazioni();
  }
  function onWriteError(err, date) {
    if (!err || err.code !== 'permission-denied') return;
    if (!net.ok) { onNetError(err); return; }                      // le regole non sono pubblicate / accesso negato
    /* una scrittura più vecchia di quella nel cloud è rifiutata di proposito e l'ascolto riporta il valore più recente.
       Se invece la modifica locale resta più recente del cloud, è l'orologio del dispositivo a essere sbagliato. */
    if (date) setTimeout(function () {
      if (user && E.reconcileDays(state.marks, known).push.indexOf(date) >= 0) toast('Una modifica non è stata accettata dal cloud: controlla che data e ora del telefono siano automatiche');
    }, 1500);
  }
  function pushDays(list) {
    list.forEach(function (d) {
      var m = state.marks[d];
      if (!m || pushedDay[d] === m.t) return;
      pushedDay[d] = m.t;
      Promise.resolve(fb.writeDay(user.uid, d, m)).catch(function (e) { onWriteError(e, d); });
    });
  }
  function pushSettings() {
    var d = E.settingsDoc(state);
    if (!d.updatedAt || pushedSet === d.updatedAt) return;
    pushedSet = d.updatedAt;
    Promise.resolve(fb.writeSettings(user.uid, d)).catch(onWriteError);
  }
  /* ogni salvataggio locale: invia ciò che è più recente del cloud (offline resta in coda nella cache di Firestore) */
  function noteChange() {
    if (user && stopSub) {
      pushDays(E.reconcileDays(state.marks, known).push);
      if (E.reconcileSettings(state, knownSet).action === 'push') pushSettings();
    }
    renderCloudStatus();
  }

  function authError(e) {
    var c = e && e.code;
    return ({
      'auth/popup-closed-by-user': 'Accesso annullato', 'auth/cancelled-popup-request': 'Accesso annullato',
      'auth/popup-blocked': 'Popup bloccato dal browser: consentilo e riprova',
      'auth/unauthorized-domain': 'Dominio non autorizzato: aggiungilo in Firebase › Authentication › Impostazioni › Domini autorizzati',
      'auth/operation-not-allowed': 'Abilita «Google» in Firebase › Authentication › Metodo di accesso',
      'auth/network-request-failed': 'Rete non disponibile: per accedere serve internet'
    })[c] || ('Accesso non riuscito' + (c ? ' (' + c + ')' : ''));
  }
  function signIn() {
    if (!fb || !fb.signIn) return;
    if (!navigator.onLine) { toast('Per accedere serve internet'); return; }
    fb.signIn().catch(function (e) { toast(authError(e)); });
  }
  function signOutUi() {
    if (!fb || !fb.signOut) return;
    fb.signOut().then(function () { toast('Disconnesso: i dati restano su questo dispositivo'); if (tab === 'impostazioni') renderImpostazioni(); }).catch(function () { toast('Disconnessione non riuscita'); });
  }
  /* "Esci e cancella i dati": per dispositivi condivisi, svuota anche la cache offline di Firestore */
  function signOutWipe() {
    if (!confirm('Uscire e cancellare dal dispositivo tutte le assenze (restano nel cloud)?')) return;
    keepPrev(); state = E.defaults(); undoStack = [];
    try { localStorage.removeItem(KEY_UID); localStorage.removeItem(KEY); } catch (e) { /* ignora */ }
    Promise.resolve(fb.signOut()).then(function () { return fb.wipeCache ? fb.wipeCache() : null; }).catch(function () { /* ignora */ }).then(function () { location.reload(); });
  }

  function askAccountSwitch() {
    var dlg = $('#dlg-codice'), decided = false;
    if (dlg.open) dlg.close();
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-codice-t', text: 'Altro account su questo dispositivo' }));
    dlg.appendChild(h('div', { class: 'db' }, h('div', { text: 'Qui ci sono assenze salvate con un altro account Google. Cosa fare con ' + user.email + '?' }),
      h('div', { class: 'note', text: 'Unisci: le assenze di questo dispositivo vengono aggiunte al nuovo account. Parti da zero: si cancellano i dati locali (ripristinabili da Impostazioni).' })));
    dlg.appendChild(h('div', { class: 'df', style: 'flex-wrap:wrap' },
      h('button', { class: 'btn ghost', type: 'button', text: 'Esci', onclick: function () { decided = true; dlg.close(); fb.signOut(); } }),
      h('button', { class: 'btn danger', type: 'button', text: 'Parti da zero', onclick: function () {
        decided = true; dlg.close(); keepPrev(); state = E.defaults(); save(true); recalc(); undoStack = []; refresh(); startSync();
      } }),
      h('button', { class: 'btn', type: 'button', text: 'Unisci', onclick: function () { decided = true; dlg.close(); startSync(); } })));
    dlg.addEventListener('close', function once() { dlg.removeEventListener('close', once); if (!decided && user) fb.signOut(); });
    dlg.showModal();
  }
  function maybeOnboard() {
    if (user || !fb || !fb.configured || fb.failed || hasMarks() || onboarded) return;
    var dlg = $('#dlg-codice');
    if (dlg.open) return;
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-codice-t', text: 'Benvenuto' }));
    dlg.appendChild(h('div', { class: 'db' },
      h('div', { text: 'Ogni giorno è già presenza: segni solo le assenze.' }),
      h('div', { class: 'note', text: 'Hai già usato l\'app o hai cambiato telefono? Accedi con Google per ritrovare i tuoi dati e sincronizzarli tra i dispositivi.' })));
    dlg.appendChild(h('div', { class: 'df' },
      h('button', { class: 'btn ghost', type: 'button', text: 'Inizia da zero', onclick: function () { dlg.close(); } }),
      h('button', { class: 'btn gbtn', type: 'button', text: 'Accedi con Google', onclick: function () { dlg.close(); signIn(); } })));
    dlg.addEventListener('close', function once() { dlg.removeEventListener('close', once); onboarded = true; saveUI(); });
    dlg.showModal();
  }

  function cloudCard(v) {
    v.appendChild(h('h2', { id: 'cloud-card', text: 'Account e sincronizzazione (Google)' }));
    var c = h('div', { class: 'card' });
    if (!fb) c.appendChild(h('div', { class: 'note', text: 'Caricamento…' }));
    else if (!fb.configured) c.appendChild(h('div', { class: 'note', text: "Sincronizzazione non ancora attivata. Crea un progetto Firebase gratuito e incolla le sue chiavi nel blocco «firebaseConfig» di index.html (passaggi nella guida PDF). Fino ad allora l'app funziona solo su questo dispositivo." }));
    else if (fb.failed) c.appendChild(h('div', { class: 'note', text: "Impossibile caricare Firebase (serve la rete almeno la prima volta). L'app funziona comunque in locale." }));
    else if (!user) {
      c.appendChild(h('div', { class: 'note', text: "Accedi con il tuo account Google: le assenze si sincronizzano tra i dispositivi e, anche se disinstalli l'app o cambi telefono, le ritrovi accedendo di nuovo. I dati sono privati: solo il tuo account può leggerli." }));
      c.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn gbtn', type: 'button', text: 'Accedi con Google', onclick: signIn })));
    } else {
      c.appendChild(h('div', { class: 'row acct' },
        user.photo ? h('img', { class: 'avatar', src: user.photo, alt: '', referrerpolicy: 'no-referrer' }) : null,
        h('div', { class: 'l' }, h('b', { text: user.name || user.email }), user.name ? h('small', { text: user.email }) : null)));
      c.appendChild(h('div', { class: 'row' }, h('div', { class: 'l' }, h('span', { id: 'cloudst2', class: 'cloudline' }))));
      c.appendChild(h('div', { class: 'btns' },
        h('button', { class: 'btn ghost', type: 'button', text: 'Esci', onclick: signOutUi }),
        h('button', { class: 'btn ghost', type: 'button', text: 'Esci e cancella i dati qui', onclick: signOutWipe })));
      c.appendChild(h('div', { class: 'note', text: "Offline le modifiche restano sul dispositivo (cache di Firestore) e partono da sole quando torna la rete. Le assenze fatte sull'altro dispositivo compaiono in pochi secondi. I dati sono privati: solo il tuo account può leggerli." }));
    }
    v.appendChild(c);
  }

  /* ---------- aggiornamento automatico dell'app ---------- */
  var buildId = null, updating = false, CORE_FILES = ['index.html', 'app.js', 'engine.js', 'xlsx.js', 'styles.css', 'manifest.webmanifest'];
  function hashText(s) {
    var h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (var i = 0, ch; i < s.length; i++) { ch = s.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  }
  function fingerprint() {
    return Promise.all(CORE_FILES.map(function (f) {
      return fetch(f, { cache: 'no-cache' }).then(function (r) { if (!r.ok) throw new Error(f); return r.text(); });
    })).then(function (t) { return hashText(t.join('\u0000')); });
  }
  /* se i file su GitHub sono cambiati rispetto a quelli in uso, ricarica da sola (i dati sono già salvati) */
  function checkUpdate() {
    if (updating || !navigator.onLine) return;
    if (navigator.serviceWorker && navigator.serviceWorker.getRegistration) navigator.serviceWorker.getRegistration().then(function (r) { if (r) r.update(); }).catch(function () { /* ignora */ });
    fingerprint().then(function (id) {
      if (!buildId) { buildId = id; return; }
      if (id === buildId || document.querySelector('dialog[open]')) return;
      updating = true; toast('Nuova versione: aggiorno…');
      setTimeout(function () { location.reload(); }, 1200);
    }).catch(function () { /* offline o errore: riprova più tardi */ });
  }

  /* ---------- stampa / PDF ---------- */
  function printReport() {
    if (!res.ok) { toast(res.check); return; }
    var r = summaryRows(), rep = $('#report'); rep.textContent = '';
    rep.appendChild(h('h1', { text: 'StayTrack - Riepilogo' }));
    rep.appendChild(h('p', { text: 'Aggiornato al ' + E.fmtDate(res.oggi) + '  ·  periodo ' + E.fmtDate(state.settings.inizio) + ' - ' + E.fmtDate(state.settings.fine) }));
    function tbl(list) { return h('table', null, h('tbody', null, list.map(function (x) { return h('tr', null, h('th', { text: x[0] }), h('td', { text: String(x[1]) })); }))); }
    rep.appendChild(h('h2', { text: 'Stato attuale' })); rep.appendChild(tbl(r.stato));
    rep.appendChild(h('h2', { text: 'Periodo e base di calcolo' })); rep.appendChild(tbl(r.periodo));
    rep.appendChild(h('h2', { text: 'Proiezione mensile' })); rep.appendChild(tbl(r.proiezione));
    rep.appendChild(h('h2', { text: 'Dettaglio mensile' })); rep.appendChild(mesiTable());
    rep.appendChild(h('h2', { text: 'Assenze segnate' }));
    rep.appendChild(res.elenco.length
      ? h('table', null, h('thead', null, h('tr', null, ['N°', 'Data', 'Tipo', 'Nota'].map(function (x) { return h('th', { text: x }); }))),
        h('tbody', null, res.elenco.map(function (e) { return h('tr', null, h('td', { text: e.n }), h('td', { text: E.fmtLong(e.date) }), h('td', { text: e.tipo }), h('td', { text: e.nota })); })))
      : h('p', { text: 'Nessuna assenza segnata.' }));
    document.body.classList.add('printing');
    window.print();
    setTimeout(function () { document.body.classList.remove('printing'); }, 500);
  }

  /* ---------- avvio ---------- */
  function startServiceWorker() {
    if (!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol)) return;
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(function (e) { console.warn('Service worker non registrato', e); });
  }

  /* scorciatoie dell'icona (tieni premuto sull'icona): ?azione=assenza-oggi  e  ?tab=riepilogo */
  function handleLaunch() {
    var q; try { q = new URLSearchParams(location.search); } catch (e) { return; }
    var act = q.get('azione'), t = q.get('tab');
    if (!act && !t) return;
    try { history.replaceState(null, '', location.pathname); } catch (e) { /* ignora */ }
    if (t && TABS.some(function (x) { return x[0] === t; })) setTab(t);
    if (act === 'assenza-oggi') {
      var d = res.byDate && res.byDate[res.oggi];
      if (!d) toast('Oggi è fuori dal periodo mostrato');
      else if (!d.valido) toast(reason(d));
      else if (d.stato === 'A') toast('Oggi è già segnato come assenza');
      else applyMark(res.oggi, 'A');
    }
  }
  /* a mezzanotte "oggi" cambia anche se l'app resta aperta */
  function scheduleMidnight() {
    var n = new Date(), next = new Date(n.getFullYear(), n.getMonth(), n.getDate() + 1, 0, 0, 5);
    setTimeout(function () { lastToday = E.todayISO(); recalc(); refresh(); scheduleMidnight(); }, Math.max(1000, next - n));
  }

  function init() {
    if ('scrollRestoration' in history) history.scrollRestoration = 'manual';   // all'apertura si va sempre al mese corrente
    try { localStorage.removeItem('assenze-collegio:cloud'); localStorage.removeItem('assenze-collegio:cloud-draft'); } catch (e) { /* ignora: chiavi della v1.x */ }
    state = load(); loadUI(); recalc(); buildNav();
    TABS.forEach(function (t) { $('#v-' + t[0]).hidden = t[0] !== tab; });
    renderTab(tab);
    window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; if (tab === 'impostazioni') renderImpostazioni(); });
    window.addEventListener('resize', syncStickyHeight);
    window.addEventListener('online', function () { renderCloudStatus(); checkUpdate(); });
    window.addEventListener('offline', renderCloudStatus);
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState !== 'visible') return;
      if (E.todayISO() !== lastToday) { lastToday = E.todayISO(); recalc(); refresh(); }
      checkUpdate();
    });
    window.addEventListener('afterprint', function () { document.body.classList.remove('printing'); });
    document.addEventListener('click', function once() {
      document.removeEventListener('click', once);
      if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () { /* ignora */ });
    });
    startServiceWorker();
    initFirebase();
    handleLaunch();
    scheduleMidnight();
    if (loadProblem) toast('I dati locali non erano leggibili: ne ho conservato una copia. Accedi con Google o importa un backup per ripristinarli.');
    else if (storageBroken) toast('Memoria del browser non disponibile (navigazione privata?): i dati non restano alla chiusura. Accedi con Google o esporta un backup.');
    /* un'altra scheda/finestra dell'app ha modificato i dati: aggiorna questa */
    window.addEventListener('storage', function (e) {
      if (e.key !== KEY || !e.newValue) return;
      try { state = E.normalize(JSON.parse(e.newValue)); undoStack = []; recalc(); calDirty = true; renderTab(tab); updateNav(); } catch (err) { /* ignora */ }
    });
    /* errore imprevisto nel codice dell'app: avvisa senza bloccare (i dati sono già salvati) */
    var errShown = false;
    window.addEventListener('error', function (e) {
      if (errShown || !e.filename || !/(app|engine|xlsx)\.js/.test(e.filename)) return;
      errShown = true; toast("Errore imprevisto: i dati sono salvati. Se qualcosa non va, ricarica l'app.");
    });
    setTimeout(checkUpdate, 2500);
    setInterval(checkUpdate, 10 * 60 * 1000);
  }
  init();
})();
