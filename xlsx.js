/* Assenze Collegio - export/import in Excel.
   Il file esportato e' il foglio Excel "Assenze Collegio" (modello.xlsx) con le assenze e le impostazioni dell'app gia' inserite:
   formule, tendine, colori, grafico e protezione restano quelli del foglio originale, quindi si puo' continuare a usarlo in Excel.
   Nessuna libreria esterna: lettura/scrittura ZIP minimale (deflate tramite DecompressionStream/CompressionStream del browser). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./engine.js'));
  else root.Xlsx = factory(root.Engine);
})(typeof self !== 'undefined' ? self : this, function (E) {
  'use strict';

  var enc = new TextEncoder(), dec = new TextDecoder();
  var BLK0 = 14, HGT = 15, STATUS0 = BLK0 + 3;          // prima riga di stato = 17 (come nel foglio Excel)
  var COLS = 'ABCDEFG';
  var EPOCH = 25569;                                      // numero seriale Excel di 1970-01-01

  /* ---------- ZIP ---------- */
  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function pipe(bytes, stream) { return new Response(new Blob([bytes]).stream().pipeThrough(stream)).arrayBuffer().then(function (b) { return new Uint8Array(b); }); }
  function inflateRaw(b) { return pipe(b, new DecompressionStream('deflate-raw')); }
  function deflateRaw(b) { return typeof CompressionStream === 'undefined' ? Promise.resolve(null) : pipe(b, new CompressionStream('deflate-raw')); }

  function unzip(buf) {
    var u = buf instanceof Uint8Array ? buf : new Uint8Array(buf), dv = new DataView(u.buffer, u.byteOffset, u.byteLength);
    var i = u.length - 22;
    while (i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
    if (i < 0) return Promise.reject(new Error('File non valido'));
    var n = dv.getUint16(i + 10, true), p = dv.getUint32(i + 16, true), list = [];
    for (var k = 0; k < n; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) return Promise.reject(new Error('File non valido'));
      var nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      list.push({ method: dv.getUint16(p + 10, true), csize: dv.getUint32(p + 20, true), lho: dv.getUint32(p + 42, true), name: dec.decode(u.subarray(p + 46, p + 46 + nl)) });
      p += 46 + nl + el + cl;
    }
    var out = {};
    return list.reduce(function (prev, e) {
      return prev.then(function () {
        if (/\/$/.test(e.name)) return;
        var start = e.lho + 30 + dv.getUint16(e.lho + 26, true) + dv.getUint16(e.lho + 28, true), data = u.subarray(start, start + e.csize);
        if (e.method === 0) { out[e.name] = data.slice(); return; }
        if (e.method !== 8) throw new Error('Compressione non supportata');
        return inflateRaw(data).then(function (d) { out[e.name] = d; });
      });
    }, Promise.resolve()).then(function () { return out; });
  }

  function u16(n) { return [n & 255, (n >>> 8) & 255]; }
  function u32(n) { return [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]; }
  function zip(files, nowMs) {
    var d = new Date(nowMs || Date.now()), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    var date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    var names = Object.keys(files), chunks = [], central = [], offset = 0;
    return names.reduce(function (prev, name) {
      return prev.then(function () {
        var data = files[name], nm = enc.encode(name), crc = crc32(data);
        return deflateRaw(data).then(function (c) {
          var useC = c && c.length < data.length, body = useC ? c : data, method = useC ? 8 : 0;
          var head = new Uint8Array([].concat(u32(0x04034b50), u16(20), u16(0x0800), u16(method), u16(time), u16(date), u32(crc), u32(body.length), u32(data.length), u16(nm.length), u16(0)));
          chunks.push(head, nm, body);
          central.push(new Uint8Array([].concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(method), u16(time), u16(date), u32(crc), u32(body.length), u32(data.length),
            u16(nm.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset))), nm);
          offset += head.length + nm.length + body.length;
        });
      });
    }, Promise.resolve()).then(function () {
      var cdSize = central.reduce(function (s, x) { return s + x.length; }, 0);
      var end = new Uint8Array([].concat(u32(0x06054b50), u16(0), u16(0), u16(names.length), u16(names.length), u32(cdSize), u32(offset), u16(0)));
      return concat(chunks.concat(central, [end]));
    });
  }
  function concat(parts) {
    var len = parts.reduce(function (s, x) { return s + x.length; }, 0), out = new Uint8Array(len), p = 0;
    parts.forEach(function (x) { out.set(x, p); p += x.length; });
    return out;
  }

  /* ---------- XML ---------- */
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function unesc(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(+d); }).replace(/&amp;/g, '&'); }
  var CELL = /<c r="([A-Z]+[0-9]+)"([^>]*?)(\/>|>[\s\S]*?<\/c>)/g;

  /* edits: { 'B4': {k:'num'|'str'|'blank'|'cache', v} }; 'cache' aggiorna solo il valore memorizzato di una cella con formula */
  function patchSheet(xml, edits) {
    return xml.replace(CELL, function (m, ref, attrs, rest) {
      var e = edits[ref];
      if (!e) return m;
      var s = / s="(\d+)"/.exec(attrs), sa = s ? ' s="' + s[1] + '"' : '';
      if (e.k === 'num') return '<c r="' + ref + '"' + sa + ' t="n"><v>' + e.v + '</v></c>';
      if (e.k === 'str') return '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + esc(e.v) + '</t></is></c>';
      if (e.k === 'blank') return '<c r="' + ref + '"' + sa + '/>';
      var f = /<f[^>]*>[\s\S]*?<\/f>|<f[^>]*\/>/.exec(rest);
      if (!f) return m;
      return typeof e.v === 'number'
        ? '<c r="' + ref + '"' + sa + ' t="n">' + f[0] + '<v>' + e.v + '</v></c>'
        : '<c r="' + ref + '"' + sa + ' t="str">' + f[0] + '<v>' + esc(e.v) + '</v></c>';
    });
  }

  function sheetPaths(files) {
    var wb = dec.decode(files['xl/workbook.xml']), rels = dec.decode(files['xl/_rels/workbook.xml.rels']), map = {}, re = /<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"/g, m;
    while ((m = re.exec(wb))) {
      var t = new RegExp('<Relationship [^>]*?Id="' + m[2] + '"[^>]*?Target="([^"]+)"').exec(rels) || new RegExp('<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="' + m[2] + '"').exec(rels);
      if (t) map[unesc(m[1])] = 'xl/' + t[1].replace(/^\/?(xl\/)?/, '');
    }
    return map;
  }
  function sharedStrings(files) {
    var f = files['xl/sharedStrings.xml']; if (!f) return [];
    var xml = dec.decode(f), out = [], re = /<si>([\s\S]*?)<\/si>/g, m;
    while ((m = re.exec(xml))) { var t = '', r = /<t[^>]*>([\s\S]*?)<\/t>/g, x; while ((x = r.exec(m[1]))) t += x[1]; out.push(unesc(t)); }
    return out;
  }
  function readCells(xml, ss) {
    var out = {}, m;
    CELL.lastIndex = 0;
    while ((m = CELL.exec(xml))) {
      var ref = m[1], attrs = m[2], inner = m[3], t = /\bt="([^"]+)"/.exec(attrs);
      t = t ? t[1] : 'n';
      if (inner === '/>') continue;
      var v = /<v>([\s\S]*?)<\/v>/.exec(inner);
      if (t === 'inlineStr') { var s = '', r = /<t[^>]*>([\s\S]*?)<\/t>/g, x; while ((x = r.exec(inner))) s += x[1]; out[ref] = unesc(s); }
      else if (!v || v[1] === '') continue;
      else if (t === 's') out[ref] = ss[+v[1]];
      else if (t === 'str' || t === 'e') out[ref] = unesc(v[1]);
      else if (t === 'b') out[ref] = v[1] === '1';
      else out[ref] = parseFloat(v[1]);
    }
    return out;
  }

  /* ---------- mappa calendario <-> celle ---------- */
  function statusRef(date, first) {                       // cella di stato del giorno nel foglio Calendario
    var a = E.parse(date), f = E.parse(first), k = (a.y - f.y) * 12 + a.m - f.m;
    if (k < 0 || k > 11) return null;
    var w = Math.floor((a.d - 1 + E.dow(E.monthStart(date))) / 7);
    return COLS[E.dow(date)] + (STATUS0 + HGT * k + 2 * w);
  }
  function serial(iso) { return E.toDays(iso) + EPOCH; }
  function fromSerial(n) { return typeof n === 'number' && isFinite(n) ? E.fromDays(Math.round(n) - EPOCH) : ''; }
  function nn(x) { return Math.round(x * 1e10) / 1e10; }

  var STATUS_TXT = { regola: 'IN REGOLA', attenzione: 'ATTENZIONE: POCHE ASSENZE RIMASTE', superata: 'SOGLIA SUPERATA', verifica: 'VERIFICA IMPOSTAZIONI' };

  /* ---------- esportazione ---------- */
  function build(template, state, res, nowMs) {
    return unzip(template).then(function (files) {
      var paths = sheetPaths(files), s = state.settings, warnings = [];
      ['Calendario', 'Riepilogo', 'Mesi', 'Impostazioni'].forEach(function (n) { if (!paths[n]) throw new Error('Modello Excel non valido'); });
      var get = function (n) { return dec.decode(files[paths[n]]); }, put = function (n, xml) { files[paths[n]] = enc.encode(xml); };

      /* Calendario: assenze nelle celle di stato + valori memorizzati dei riquadri in alto */
      var e1 = {};
      if (res.days) res.days.forEach(function (d) { if (d.stato) { var r = statusRef(d.date, res.first); if (r) e1[r] = { k: 'str', v: d.stato }; } });
      if (res.ok) {
        e1.A2 = { k: 'cache', v: 'Periodo: ' + E.fmtDate(s.inizio) + ' - ' + E.fmtDate(s.fine) + '  ·  soglia ' + Math.round(s.soglia * 100) + '%' };
        e1.A4 = { k: 'cache', v: res.rimanenti }; e1.C4 = { k: 'cache', v: res.mensili }; e1.F4 = { k: 'cache', v: nn(res.pctFine) };
        e1.A5 = { k: 'cache', v: 'Assenze usate ' + res.usate + ' su ' + res.maxAss + '  ·  Giustificate ' + res.giust + '  ·  Presenze finora ' + res.presFinora };
        e1.A6 = { k: 'cache', v: STATUS_TXT[res.status] };
      }
      put('Calendario', patchSheet(get('Calendario'), e1));

      /* Riepilogo e Mesi: valori memorizzati (Excel li ricalcola comunque all'apertura) */
      if (res.ok) {
        var r2 = { B4: serial(s.inizio), B5: serial(s.fine), B6: serial(res.oggi), B7: res.giorniCalendario, B8: res.giorniEsclusi, B9: res.gv, B10: s.soglia,
          B11: res.minPres, B12: res.maxAss, B15: res.usate, B16: res.giust, B17: res.rimanenti, B18: res.gvRim, B19: res.gvTrasc, B20: res.assTrasc, B21: res.presFinora,
          B22: res.pctFinora == null ? '-' : nn(res.pctFinora), B23: res.presFine, B24: nn(res.pctFine), B25: nn(res.margine), B28: nn(res.mesiRim), B29: res.mensili,
          B30: s.baseFrazione === 'calendario' ? 'Giorni di calendario' : 'Giorni validi' };
        var e2 = {}; Object.keys(r2).forEach(function (k) { e2[k] = { k: 'cache', v: r2[k] }; });
        put('Riepilogo', patchSheet(get('Riepilogo'), e2));
        var e3 = {}, tot = { v: 0, a: 0, g: 0 };
        res.mesi.forEach(function (m, i) {
          var r = 5 + i; tot.v += m.validi; tot.a += m.ass; tot.g += m.giu;
          e3['A' + r] = { k: 'cache', v: m.label }; e3['B' + r] = { k: 'cache', v: m.validi }; e3['C' + r] = { k: 'cache', v: m.ass }; e3['D' + r] = { k: 'cache', v: m.giu };
          e3['E' + r] = { k: 'cache', v: m.pres }; e3['F' + r] = { k: 'cache', v: m.pct == null ? '-' : nn(m.pct) }; e3['G' + r] = { k: 'cache', v: nn(m.peso) };
        });
        e3.B17 = { k: 'cache', v: tot.v }; e3.C17 = { k: 'cache', v: tot.a }; e3.D17 = { k: 'cache', v: tot.g }; e3.E17 = { k: 'cache', v: tot.v - tot.a };
        e3.F17 = { k: 'cache', v: tot.v ? nn((tot.v - tot.a) / tot.v) : '-' }; e3.G17 = { k: 'cache', v: nn(res.mesiRim) };
        put('Mesi', patchSheet(get('Mesi'), e3));
      }

      /* Impostazioni */
      var e5 = {}, num = function (v) { return { k: 'num', v: v }; }, str = function (v) { return v ? { k: 'str', v: v } : { k: 'blank' }; };
      var dt = function (iso) { return iso && E.isValidISO(iso) ? num(serial(iso)) : { k: 'blank' }; };
      e5.B4 = dt(s.inizio); e5.B5 = dt(s.fine); e5.B6 = num(s.soglia); e5.B7 = dt(s.oggiOverride);
      e5.B9 = str(s.baseFrazione === 'calendario' ? 'Giorni di calendario' : 'Giorni validi'); e5.B10 = num(s.allerta);
      if (state.chiusure.length > 8) warnings.push('chiusure: esportate le prime 8 su ' + state.chiusure.length);
      for (var i = 0; i < 8; i++) { var c = state.chiusure[i] || {}, r = 16 + i; e5['A' + r] = dt(c.dal); e5['B' + r] = dt(c.al); e5['C' + r] = str(c.desc); }
      if (state.festivita.length > 30) warnings.push('festività: esportate le prime 30 su ' + state.festivita.length);
      for (var j = 0; j < 30; j++) { var f = state.festivita[j] || {}, rr = 27 + j; e5['A' + rr] = dt(f.data); e5['C' + rr] = str(f.desc); }
      put('Impostazioni', patchSheet(get('Impostazioni'), e5));

      /* l'Excel ricalcola tutto all'apertura */
      var wb = dec.decode(files['xl/workbook.xml']);
      wb = /<calcPr[^>]*fullCalcOnLoad/.test(wb) ? wb : wb.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
      files['xl/workbook.xml'] = enc.encode(wb);
      return zip(files, nowMs).then(function (bytes) { return { bytes: bytes, warnings: warnings }; });
    });
  }

  /* ---------- importazione ---------- */
  function parse(bytes) {
    return unzip(bytes).then(function (files) {
      if (!files['xl/workbook.xml']) throw new Error('Non è un file Excel valido');
      var paths = sheetPaths(files), ss = sharedStrings(files);
      if (!paths.Calendario || !paths.Impostazioni) throw new Error('File Excel non riconosciuto (mancano i fogli Calendario e Impostazioni)');
      var imp = readCells(dec.decode(files[paths.Impostazioni]), ss), cal = readCells(dec.decode(files[paths.Calendario]), ss), warnings = [];
      var inizio = fromSerial(imp.B4), fine = fromSerial(imp.B5);
      if (!E.isValidISO(inizio) || !E.isValidISO(fine)) throw new Error('Date del periodo non valide nel foglio Impostazioni');
      var soglia = typeof imp.B6 === 'number' ? imp.B6 : 0.56;
      var settings = { inizio: inizio, fine: fine, soglia: soglia > 1 ? soglia / 100 : soglia, oggiOverride: fromSerial(imp.B7) || '',
        baseFrazione: /calendario/i.test(String(imp.B9 || '')) ? 'calendario' : 'validi', allerta: typeof imp.B10 === 'number' && imp.B10 >= 0 ? Math.round(imp.B10) : 10 };
      var chiusure = [], festivita = [];
      for (var r = 16; r <= 23; r++) { var a = fromSerial(imp['A' + r]), b = fromSerial(imp['B' + r]), d = imp['C' + r] ? String(imp['C' + r]) : ''; if (a || b) chiusure.push({ dal: a, al: b, desc: d }); }
      for (var q = 27; q <= 56; q++) { var fd = fromSerial(imp['A' + q]); if (fd) festivita.push({ data: fd, desc: imp['C' + q] ? String(imp['C' + q]) : '' }); }
      var first = E.addMonths(E.monthStart(fine), -11), marks = {}, other = 0;
      for (var k = 0; k < 12; k++) {
        var ms = E.addMonths(first, k), pm = E.parse(ms), dim = E.daysInMonth(pm.y, pm.m);
        for (var day = 1; day <= dim; day++) {
          var date = E.iso(pm.y, pm.m, day), ref = statusRef(date, first), v = ref ? cal[ref] : null;
          if (v == null || v === '') continue;
          var u = String(v).trim().toUpperCase();
          if (u === 'A' || u === 'G') marks[date] = u; else if (u !== 'P') other++;
        }
      }
      if (other) warnings.push(other + ' celle con codici non riconosciuti ignorate');
      return { settings: settings, chiusure: chiusure, festivita: festivita, marks: marks, warnings: warnings };
    });
  }

  return { build: build, parse: parse, unzip: unzip, zip: zip, crc32: crc32, statusRef: statusRef, serial: serial, fromSerial: fromSerial };
});
