/* Assenze Collegio - interfaccia. Il calcolo è in engine.js (stessa logica del foglio Excel). */
(function () {
  'use strict';
  var E = window.Engine;
  var KEY = 'assenze-collegio:v1', KEY_PREV = 'assenze-collegio:prev', KEY_UI = 'assenze-collegio:ui';
  var VERSION = '1.0';
  var DAY = 86400000;

  var state, res, brush = 'A', undoStack = [], tab = 'calendario', calDirty = true, storageOk = true;
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
  function load() {
    try {
      var raw = localStorage.getItem(KEY);
      if (raw) return E.normalize(JSON.parse(raw));
    } catch (e) { console.warn('Dati non leggibili', e); }
    return E.defaults();
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(state)); storageOk = true; }
    catch (e) { storageOk = false; toast('Attenzione: impossibile salvare su questo dispositivo'); }
  }
  function keepPrev() { try { var raw = localStorage.getItem(KEY); if (raw) localStorage.setItem(KEY_PREV, raw); } catch (e) { /* ignora */ } }
  function loadUI() { try { var u = JSON.parse(localStorage.getItem(KEY_UI) || '{}'); if (u.brush) brush = u.brush; } catch (e) { /* ignora */ } }
  function saveUI() { try { localStorage.setItem(KEY_UI, JSON.stringify({ brush: brush })); } catch (e) { /* ignora */ } }
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

  function buildNav() {
    var nav = $('#tabs');
    nav.textContent = '';
    TABS.forEach(function (t) {
      var b = h('button', { type: 'button', 'data-tab': t[0], onclick: function () { setTab(t[0]); } });
      b.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICONS[t[0]] + '</svg>';
      b.appendChild(h('span', { text: t[1] }));
      nav.appendChild(b);
    });
    updateNav();
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
  function statusText() {
    return { regola: 'IN REGOLA', attenzione: 'ATTENZIONE: POCHE ASSENZE RIMASTE', superata: 'SOGLIA SUPERATA', verifica: 'VERIFICA IMPOSTAZIONI' }[res.status];
  }
  function renderKpis() {
    if (!$('#k-rim')) return;
    var okd = res.days && res.days.length;
    $('#k-rim').textContent = okd ? res.rimanenti : '–';
    $('#k-rim').parentNode.classList.toggle('neg', okd && res.rimanenti < 0);
    $('#k-men').textContent = okd ? nf(res.mensili, 2) : '–';
    $('#k-pct').textContent = okd ? pc(res.pctFine) : '–';
    var st = $('#status');
    st.className = 'status' + (res.status && res.status !== 'regola' ? ' ' + res.status : '');
    st.textContent = okd ? statusText() : 'VERIFICA IMPOSTAZIONI';
    $('#info').textContent = okd
      ? 'Assenze usate ' + res.usate + ' su ' + res.maxAss + '  ·  Giustificate ' + res.giust + '  ·  Presenze finora ' + res.presFinora
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
    b.querySelector('.s').textContent = day.stato;
    var what = day.festivo ? 'festività ' + day.motivo : (day.chiuso ? 'residenza chiusa' : (day.classe === 3 ? 'fuori periodo'
      : (day.stato === 'A' ? 'assenza' : (day.stato === 'G' ? 'assenza giustificata' : 'presente'))));
    b.setAttribute('aria-label', E.fmtLong(day.date) + ', ' + what);
  }
  function dayEl(day) {
    var b = h('button', { class: 'day', type: 'button', 'data-date': day.date },
      h('span', { class: 'd', text: String(parseInt(day.date.slice(8), 10)) }), h('span', { class: 's' }));
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
    var prev = state.marks[date] ? state.marks[date].s : '';
    state = E.setMark(state, date, target, Date.now());
    undoStack.push([{ date: date, prev: prev }]);
    if (undoStack.length > 50) undoStack.shift();
    afterChange([date]);
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
    var sticky = h('div', { class: 'sticky', id: 'sticky' },
      h('div', { class: 'kpis' },
        h('div', { class: 'kpi k1' }, h('small', { text: 'Assenze rimanenti' }), h('b', { id: 'k-rim' })),
        h('div', { class: 'kpi k2' }, h('small', { text: 'Mensili rimanenti' }), h('b', { id: 'k-men' })),
        h('div', { class: 'kpi k3' }, h('small', { text: 'Presenza a fine anno' }), h('b', { id: 'k-pct' }))),
      h('div', { class: 'status', id: 'status' }),
      h('div', { class: 'info', id: 'info' }),
      h('div', { class: 'toolbar' },
        h('div', { class: 'brushes', role: 'radiogroup', 'aria-label': 'Cosa segnare toccando un giorno' },
          brushBtn('A', 'Assenza'), brushBtn('G', 'Giustificata'), brushBtn('P', 'Presente')),
        h('button', { class: 'tbtn', type: 'button', text: 'Periodo…', onclick: openPeriodo }),
        h('button', { class: 'tbtn', id: 'btn-undo', type: 'button', 'aria-label': 'Annulla ultima modifica', text: '↶', onclick: undo })));
    v.appendChild(sticky);

    if (!res.days || !res.days.length) {
      v.appendChild(h('div', { class: 'warnbox', text: 'Imposta le date del periodo in Impostazioni: ' + res.check }));
      calDirty = false; renderKpis(); return;
    }
    var chips = h('div', { class: 'chips' });
    chips.appendChild(h('button', { class: 'chip now', type: 'button', text: '▸ Oggi', onclick: function () { scrollToMonth(currentMonthKey(), true); } }));
    res.mesi.forEach(function (m) { chips.appendChild(h('button', { class: 'chip', type: 'button', text: m.labelAb, onclick: function () { scrollToMonth(m.key, true); } })); });
    v.appendChild(chips);

    var months = h('div', { class: 'months' });
    months.addEventListener('click', function (e) { var b = e.target.closest('.day'); if (b) onDay(b.dataset.date); });
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
    calDirty = false;
    renderKpis(); updateUndoBtn();
    syncStickyHeight();
    if (firstCalendarRender) { firstCalendarRender = false; if (res.mesi.some(function (m) { return m.key.slice(0, 7) === currentMonthKey(); })) scrollToMonth(currentMonthKey(), false); }
    else window.scrollTo(0, y);
  }
  function brushBtn(b, label) {
    return h('button', { class: 'brush', type: 'button', role: 'radio', 'data-b': b, 'aria-checked': 'false', onclick: function () { setBrush(b); } },
      b, h('small', { text: label }));
  }
  function syncStickyHeight() {
    var s = $('#sticky');
    if (s) document.documentElement.style.setProperty('--sticky-h', s.offsetHeight + 'px');
  }

  /* segna un periodo */
  function openPeriodo() {
    var dlg = $('#dlg-periodo'), s = state.settings;
    var from = h('input', { type: 'date', id: 'p-from', value: res.oggi >= s.inizio && res.oggi <= s.fine ? res.oggi : s.inizio, min: s.inizio, max: s.fine });
    var to = h('input', { type: 'date', id: 'p-to', value: from.value, min: s.inizio, max: s.fine });
    var kind = 'A';
    var prev = h('div', { class: 'info', id: 'p-prev' });
    function radio(v, lab) {
      var i = h('input', { type: 'radio', name: 'pk', value: v, onchange: function () { kind = v; upd(); } });
      if (v === 'A') i.checked = true;
      return h('label', null, i, h('span', { text: lab }));
    }
    function upd() {
      if (!from.value || !to.value || to.value < from.value) { prev.textContent = 'Controlla le date'; return; }
      var r = E.paintRange(state, from.value, to.value, kind, 0, res);
      prev.textContent = r.changed.length + ' giorni validi verranno modificati';
      go.disabled = r.changed.length === 0;
    }
    var go = h('button', { class: 'btn', type: 'button', text: 'Applica', onclick: function () {
      var r = E.paintRange(state, from.value, to.value, kind, Date.now(), res);
      if (!r.changed.length) return;
      state = r.state;
      undoStack.push(r.changed.map(function (c) { return { date: c.date, prev: c.prev ? c.prev.s : '' }; }));
      afterChange(r.changed.map(function (c) { return c.date; }));
      dlg.close();
      toast(r.changed.length + ' giorni aggiornati', { label: 'ANNULLA', fn: undo });
    } });
    from.addEventListener('change', function () { if (to.value < from.value) to.value = from.value; upd(); });
    to.addEventListener('change', upd);
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-periodo-t', text: 'Segna un periodo' }));
    dlg.appendChild(h('div', { class: 'db' },
      h('div', { class: 'field' }, h('label', { for: 'p-from', text: 'Dal' }), from),
      h('div', { class: 'field' }, h('label', { for: 'p-to', text: 'Al (incluso)' }), to),
      h('div', { class: 'radios' }, radio('A', 'A Assenza'), radio('G', 'G Giust.'), radio('P', 'P Presente')),
      prev,
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
        ['ASSENZE RIMANENTI TOTALI', res.rimanenti, 'Negativo = soglia superata', 'big' + (res.rimanenti < 0 ? ' neg' : '')],
        ['Giorni validi rimanenti (da oggi incluso)', res.gvRim], ['Giorni validi trascorsi', res.gvTrasc], ['Assenze nei giorni trascorsi', res.assTrasc],
        ['Presenze finora', res.presFinora], ['% presenza finora', pc(res.pctFinora), 'Sui giorni validi trascorsi'],
        ['Presenze a fine periodo', res.presFine, 'Se da oggi sempre presente'], ['% presenza a fine periodo', pc(res.pctFine), 'Se da oggi sempre presente'],
        ['Margine sulla soglia', (res.margine >= 0 ? '+' : '') + nf(res.margine * 100, 1) + ' punti', null, res.margine < 0 ? 'neg' : '']],
      proiezione: [
        ['Mesi rimanenti (con frazioni)', nf(res.mesiRim, 2), 'Frazione del mese corrente + mesi interi successivi'],
        ['ASSENZE MENSILI RIMANENTI', nf(res.mensili, 2), 'Assenze rimanenti ÷ mesi rimanenti', 'big'],
        ['Base della frazione di mese', s.baseFrazione === 'calendario' ? 'Giorni di calendario' : 'Giorni validi']]
    };
  }
  function renderRiepilogo() {
    var v = $('#v-riepilogo'); v.textContent = '';
    v.appendChild(h('h1', { text: 'Riepilogo' }));
    if (!res.ok) { v.appendChild(h('div', { class: 'check bad', text: res.check })); return; }
    var r = summaryRows();
    v.appendChild(h('h2', { text: 'Periodo e base di calcolo' })); v.appendChild(rows(r.periodo));
    v.appendChild(h('h2', { text: 'Stato attuale' })); v.appendChild(rows(r.stato));
    v.appendChild(h('h2', { text: 'Proiezione mensile' })); v.appendChild(rows(r.proiezione));
    v.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn ghost', type: 'button', text: 'Stampa / salva PDF', onclick: printReport })));
  }

  /* ---------- MESI ---------- */
  function chart(mesi) {
    var W = 360, H = 170, base = 140, left = 22, mx = Math.max(3, Math.max.apply(null, mesi.map(function (m) { return Math.max(m.ass, m.giu); })));
    var bw = (W - left - 6) / mesi.length, svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Assenze e giustificate per mese">';
    for (var t = 0; t <= 3; t++) {
      var val = Math.round(mx * t / 3), y = base - (base - 14) * t / 3;
      svg += '<line x1="' + left + '" x2="' + W + '" y1="' + y + '" y2="' + y + '" stroke="currentColor" opacity=".12"/><text x="' + (left - 4) + '" y="' + (y + 3) + '" text-anchor="end">' + val + '</text>';
    }
    mesi.forEach(function (m, i) {
      var x = left + i * bw + bw * 0.12, w = bw * 0.38, ha = (base - 14) * m.ass / mx, hg = (base - 14) * m.giu / mx;
      svg += '<rect x="' + x.toFixed(1) + '" y="' + (base - ha).toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + ha.toFixed(1) + '" rx="2" fill="#d64545"/>';
      svg += '<rect x="' + (x + w + 1).toFixed(1) + '" y="' + (base - hg).toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + hg.toFixed(1) + '" rx="2" fill="#f5b942"/>';
      svg += '<text x="' + (x + w).toFixed(1) + '" y="' + (base + 12) + '" text-anchor="middle">' + E.MESI_AB[m.m - 1] + '</text>';
    });
    svg += '<rect x="' + left + '" y="' + (H - 11) + '" width="8" height="8" rx="2" fill="#d64545"/><text class="lg" x="' + (left + 12) + '" y="' + (H - 3) + '">Assenze</text>';
    svg += '<rect x="' + (left + 66) + '" y="' + (H - 11) + '" width="8" height="8" rx="2" fill="#f5b942"/><text class="lg" x="' + (left + 78) + '" y="' + (H - 3) + '">Giustificate</text>';
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
    v.appendChild(h('h1', { text: 'Dettaglio mensile' }));
    v.appendChild(h('p', { class: 'lead', text: 'Presenze = giorni validi − assenze (le giustificate valgono come presenza). Peso = parte del mese ancora da vivere, per la proiezione mensile.' }));
    if (!res.mesi.length) return;
    var c = h('div', { class: 'card' }, mesiTable());
    v.appendChild(c);
    var ch = h('div', { class: 'card chart', style: 'margin-top:12px' }); ch.innerHTML = chart(res.mesi); v.appendChild(ch);
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
  function field(label, hint, input) {
    return h('div', { class: 'field' }, h('label', null, label, hint ? h('small', { text: hint }) : null), input);
  }
  function dateInput(get, set) {
    return h('input', { type: 'date', value: get() || '', onchange: function (e) { set(e.target.value || ''); touchSettings(); } });
  }
  function renderImpostazioni() {
    var y = window.scrollY, v = $('#v-impostazioni'), s = state.settings;
    v.textContent = '';
    v.appendChild(h('h1', { text: 'Impostazioni' }));
    v.appendChild(h('div', { class: 'check', id: 'chk' }));
    var par = h('div', { class: 'card' });
    par.appendChild(field('Data di ingresso', 'Inizio del periodo di conteggio', dateInput(function () { return s.inizio; }, function (x) { s.inizio = x; })));
    par.appendChild(field('Fine periodo', null, dateInput(function () { return s.fine; }, function (x) { s.fine = x; })));
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
        dateInput(function () { return c.dal; }, function (x) { c.dal = x; }), dateInput(function () { return c.al; }, function (x) { c.al = x; }),
        h('button', { class: 'x', type: 'button', 'aria-label': 'Elimina chiusura', text: '✕', onclick: function () { state.chiusure.splice(i, 1); touchSettings(); renderImpostazioni(); } }),
        h('input', { class: 'desc', type: 'text', maxlength: 80, placeholder: 'Descrizione', value: c.desc || '', onchange: function (e) { c.desc = e.target.value; touchSettings(); } })));
    });
    cc.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn ghost', type: 'button', text: '＋ Aggiungi chiusura', onclick: function () { state.chiusure.push({ dal: '', al: '', desc: '' }); renderImpostazioni(); } })));
    v.appendChild(cc);

    v.appendChild(h('h2', { text: 'Festività (escluse dal calcolo)' }));
    var fc = h('div', { class: 'card' });
    state.festivita.forEach(function (f, i) {
      fc.appendChild(h('div', { class: 'lrow f' },
        dateInput(function () { return f.data; }, function (x) { f.data = x; }),
        h('input', { class: 'desc', type: 'text', maxlength: 80, placeholder: 'Descrizione', value: f.desc || '', onchange: function (e) { f.desc = e.target.value; touchSettings(); } }),
        h('button', { class: 'x', type: 'button', 'aria-label': 'Elimina festività', text: '✕', onclick: function () { state.festivita.splice(i, 1); touchSettings(); renderImpostazioni(); } })));
    });
    fc.appendChild(h('div', { class: 'btns' }, h('button', { class: 'btn ghost', type: 'button', text: '＋ Aggiungi festività', onclick: function () { state.festivita.push({ data: '', desc: '' }); renderImpostazioni(); } })));
    fc.appendChild(h('div', { class: 'note', text: 'Pasqua e Pasquetta incluse; San Nicola (6 maggio) no. Le date della chiusura natalizia sono provvisorie.' }));
    v.appendChild(fc);

    /* dati */
    v.appendChild(h('h2', { text: 'Backup e sincronizzazione tra dispositivi' }));
    if (needsBackup()) v.appendChild(h('div', { class: 'warnbox', text: 'Da più di 14 giorni non fai un backup: esporta un file e conservalo (Drive, email…).' }));
    var last = state.meta.lastExportAt ? new Date(state.meta.lastExportAt).toLocaleDateString('it-IT') : 'mai';
    var dc = h('div', { class: 'card' });
    var canShare = !!(navigator.canShare && typeof File !== 'undefined');
    dc.appendChild(h('div', { class: 'btns' },
      h('button', { class: 'btn', type: 'button', text: 'Salva backup', onclick: function () { exportFile(false); } }),
      canShare ? h('button', { class: 'btn ghost', type: 'button', text: 'Condividi backup', onclick: function () { exportFile(true); } }) : null,
      h('button', { class: 'btn ghost', type: 'button', text: 'Importa file…', onclick: function () { $('#file-in').click(); } }),
      h('input', { type: 'file', id: 'file-in', accept: '.json,application/json', hidden: true, onchange: importFile })));
    dc.appendChild(h('div', { class: 'btns', style: 'padding-top:0' },
      h('button', { class: 'btn ghost', type: 'button', text: 'Copia codice', onclick: copyCode }),
      h('button', { class: 'btn ghost', type: 'button', text: 'Incolla codice…', onclick: function () { openCodice(); } })));
    dc.appendChild(h('div', { class: 'note', text: 'Ultimo backup: ' + last + '. Per usare due dispositivi: esporta (file o codice) da uno e importa nell\'altro. I dati si uniscono: per ogni giorno vince la modifica più recente e nulla va perso.' }));
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
    refreshCheck();
    if (navigator.storage && navigator.storage.persisted) navigator.storage.persisted().then(function (p) {
      var el = $('#persist-s'); if (el) el.textContent = p ? 'Sì: il browser non cancellerà i dati da solo' : 'Non ancora garantita: fai ogni tanto un backup';
    });
    window.scrollTo(0, y);
  }

  /* ---------- backup / importazione ---------- */
  function fileName() { return 'assenze-collegio-' + E.todayISO() + '.json'; }
  function exportFile(share) {
    state.meta.lastExportAt = Date.now(); save();
    var json = JSON.stringify(E.exportObject(state, Date.now()), null, 1);
    var blob = new Blob([json], { type: 'application/json' });
    if (share && navigator.canShare) {
      var f = new File([blob], fileName(), { type: 'application/json' });
      if (navigator.canShare({ files: [f] })) { navigator.share({ files: [f], title: 'Backup assenze collegio' }).catch(function () { /* annullato */ }); updateNav(); renderImpostazioni(); return; }
    }
    var a = h('a', { href: URL.createObjectURL(blob), download: fileName() });
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    toast('Backup salvato: ' + fileName()); updateNav(); renderImpostazioni();
  }
  function importFile(e) {
    var f = e.target.files && e.target.files[0]; e.target.value = '';
    if (!f) return;
    f.text().then(function (t) { previewImport(t); }).catch(function () { toast('File non leggibile'); });
  }
  function toCode() { return 'ASSENZE1:' + btoa(unescape(encodeURIComponent(JSON.stringify(E.exportObject(state, Date.now()))))); }
  function fromCode(code) {
    var m = /^\s*ASSENZE1:([A-Za-z0-9+/=\s]+)$/.exec(code);
    if (!m) throw new Error('Codice non valido');
    return decodeURIComponent(escape(atob(m[1].replace(/\s+/g, ''))));
  }
  function copyCode() {
    state.meta.lastExportAt = Date.now(); save(); updateNav();
    var code = toCode();
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(code).then(function () { toast('Codice copiato: incollalo sull\'altro dispositivo'); renderImpostazioni(); }, function () { openCodice(code); });
    else openCodice(code);
  }
  function openCodice(code) {
    var dlg = $('#dlg-codice'), ta = h('textarea', { 'aria-label': 'Codice di sincronizzazione', spellcheck: 'false', placeholder: 'Incolla qui il codice ASSENZE1:…' });
    if (code) ta.value = code;
    dlg.textContent = '';
    dlg.appendChild(h('div', { class: 'dh', id: 'dlg-codice-t', text: code ? 'Codice di sincronizzazione' : 'Incolla il codice' }));
    dlg.appendChild(h('div', { class: 'db' }, ta, h('div', { class: 'note', text: code ? 'Copia tutto il testo e incollalo sull\'altro dispositivo (Impostazioni › Incolla codice).' : 'Il codice si genera sull\'altro dispositivo con "Copia codice".' })));
    dlg.appendChild(h('div', { class: 'df' }, h('button', { class: 'btn ghost', type: 'button', text: 'Chiudi', onclick: function () { dlg.close(); } }),
      code ? null : h('button', { class: 'btn', type: 'button', text: 'Unisci', onclick: function () { try { var t = fromCode(ta.value); dlg.close(); previewImport(t); } catch (err) { toast('Codice non valido'); } } })));
    dlg.showModal(); if (code) { ta.focus(); ta.select(); }
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

  /* ---------- stampa / PDF ---------- */
  function printReport() {
    if (!res.ok) { toast(res.check); return; }
    var r = summaryRows(), rep = $('#report'); rep.textContent = '';
    rep.appendChild(h('h1', { text: 'Assenze Collegio - Riepilogo' }));
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
    var hadController = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('sw.js').catch(function (e) { console.warn('Service worker non registrato', e); });
    navigator.serviceWorker.addEventListener('controllerchange', function () {
      if (hadController) toast('App aggiornata', { label: 'RICARICA', fn: function () { location.reload(); } });
    });
  }

  function init() {
    state = load(); loadUI(); recalc(); buildNav();
    TABS.forEach(function (t) { $('#v-' + t[0]).hidden = t[0] !== tab; });
    renderTab(tab);
    window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; if (tab === 'impostazioni') renderImpostazioni(); });
    window.addEventListener('resize', syncStickyHeight);
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'visible' && E.todayISO() !== lastToday) { lastToday = E.todayISO(); recalc(); refresh(); }
    });
    window.addEventListener('afterprint', function () { document.body.classList.remove('printing'); });
    document.addEventListener('click', function once() {
      document.removeEventListener('click', once);
      if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () { /* ignora */ });
    });
    startServiceWorker();
  }
  init();
})();
