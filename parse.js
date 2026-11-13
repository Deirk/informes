var SPACES = ['PASILLO', 'SALA COMEDOR', 'COCINA', 'HALL', 'ALCOBA SERVICIO', 'ZONA DE OFICIO', 'ESTAR TV', 'BAÑO 1', 'BAÑO 2', 'BAÑO 3', 'ALCOBA 1', 'ALCOBA 2', 'ALCOBA 3'];

function parseInforme(buf, JSZipLib, DOMParserCls) {
  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  }
  var ALIAS = { 'SALA': 'SALA COMEDOR', 'COMEDOR': 'SALA COMEDOR', 'SALA COMEDOR': 'SALA COMEDOR', 'ZONA DE OFICIOS': 'ZONA DE OFICIO', 'ZONA OFICIO': 'ZONA DE OFICIO', 'ZONA DE OFICIO': 'ZONA DE OFICIO', 'ALCOBA DE SERVICIO': 'ALCOBA SERVICIO', 'ALCOBA SERVICIO': 'ALCOBA SERVICIO', 'ESTAR': 'ESTAR TV', 'ESTAR TV': 'ESTAR TV' };
  function matchSpace(t) {
    var n = norm(t);
    if (!n || n.length > 30) return null;
    if (ALIAS[n]) return ALIAS[n];
    for (var i = 0; i < SPACES.length; i++) if (norm(SPACES[i]) === n) return SPACES[i];
    return null;
  }
  function imgSize(b) {
    try {
      if (b[0] === 0x89 && b[1] === 0x50) {
        return { w: ((b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19]) >>> 0, h: ((b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23]) >>> 0 };
      }
      if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) };
      if (b[0] === 0x42 && b[1] === 0x4d) {
        var bw = (b[18] | (b[19] << 8) | (b[20] << 16) | (b[21] << 24)), bh = (b[22] | (b[23] << 8) | (b[24] << 16) | (b[25] << 24));
        return { w: Math.abs(bw), h: Math.abs(bh) };
      }
      if (b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45) {
        var fcc = String.fromCharCode(b[12], b[13], b[14], b[15]);
        if (fcc === 'VP8X') return { w: 1 + (b[24] | (b[25] << 8) | (b[26] << 16)), h: 1 + (b[27] | (b[28] << 8) | (b[29] << 16)) };
        if (fcc === 'VP8 ') return { w: (b[26] | (b[27] << 8)) & 0x3fff, h: (b[28] | (b[29] << 8)) & 0x3fff };
        if (fcc === 'VP8L') { var v = b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24); return { w: (v & 0x3fff) + 1, h: ((v >> 14) & 0x3fff) + 1 }; }
      }
      if (b[0] === 0xff && b[1] === 0xd8) {
        var i = 2;
        while (i < b.length - 9) {
          if (b[i] !== 0xff) { i++; continue; }
          var m = b[i + 1];
          if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
            return { h: (b[i + 5] << 8) | b[i + 6], w: (b[i + 7] << 8) | b[i + 8] };
          }
          i += 2 + ((b[i + 2] << 8) | b[i + 3]);
        }
      }
    } catch (e) {}
    return null;
  }
  function kids(node, name) {
    var out = [];
    for (var c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.nodeName === name) out.push(c);
    return out;
  }
  function descend(node, fn) {
    for (var c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      fn(c);
      descend(c, fn);
    }
  }

  return JSZipLib.loadAsync(buf).then(function (zip) {
    var docF = zip.file('word/document.xml');
    if (!docF) throw new Error('nodocx');
    var relF = zip.file('word/_rels/document.xml.rels');
    var numF = zip.file('word/numbering.xml');
    return Promise.all([docF.async('string'), relF ? relF.async('string') : Promise.resolve(''), numF ? numF.async('string') : Promise.resolve('')]).then(function (res) {
      var numFmt = {};
      (function () {
        var nx = new DOMParserCls().parseFromString(res[2] || '<r/>', 'application/xml');
        var absFmt = {};
        Array.prototype.forEach.call(nx.getElementsByTagName('w:abstractNum'), function (a) {
          var lv = a.getElementsByTagName('w:lvl')[0], nf = lv && lv.getElementsByTagName('w:numFmt')[0];
          absFmt[a.getAttribute('w:abstractNumId')] = nf ? nf.getAttribute('w:val') : 'bullet';
        });
        Array.prototype.forEach.call(nx.getElementsByTagName('w:num'), function (n) {
          var ab = n.getElementsByTagName('w:abstractNumId')[0];
          numFmt[n.getAttribute('w:numId')] = ab ? absFmt[ab.getAttribute('w:val')] : 'bullet';
        });
      })();
      var xml = new DOMParserCls().parseFromString(res[0], 'application/xml');
      var rels = new DOMParserCls().parseFromString(res[1] || '<r/>', 'application/xml');
      var rid = {};
      Array.prototype.forEach.call(rels.getElementsByTagName('Relationship'), function (r) { rid[r.getAttribute('Id')] = r.getAttribute('Target'); });
      var body = xml.getElementsByTagName('w:body')[0];
      var paras = [];
      descend(body, function (n) { if (n.nodeName === 'w:p') paras.push(n); });

      var mode = 'cover', cur = null, newSect = false, photos = 0;
      var cover = [], intro = [], planosTitulo = '', planos = [], espacios = {}, presu = { img: null, nota: '' }, concl = [], conclTitulo = 'CONCLUSIONES', firma = null, firmasAll = [];
      var imgJobs = [];

      function img(target, ew, eh) {
        var o = { t: 'img', target: target, bytes: null, ext: (target.split('.').pop() || '').toLowerCase(), w: 0, h: 0, ew: ew || 0, eh: eh || 0 };
        imgJobs.push(o);
        return o;
      }
      function isIntro(T) { return /^\s*(\d+[.)]\s*)?INTRODUCCI[OÓ]N\s*:?\s*$/i.test(T); }
      function isPlano(T) { return T.length <= 70 && /^\s*(\d+[.)]\s*)?PLANOS?\b/i.test(T); }
      function isConcl(T) { return T.length <= 40 && /^\s*(\d+[.)]\s*)?CONCLUSIONES\b/i.test(T); }

      paras.forEach(function (p) {
        var toks = [];
        descend(p, function (n) {
          if (n.nodeName === 'w:t') {
            var s = n.textContent;
            if (toks.length && toks[toks.length - 1].text != null) toks[toks.length - 1].text += s; else toks.push({ text: s });
          } else if (n.nodeName === 'w:tab' || n.nodeName === 'w:br') {
            if (toks.length && toks[toks.length - 1].text != null) toks[toks.length - 1].text += ' ';
          } else if (n.nodeName === 'a:blip' || n.nodeName === 'v:imagedata') {
            var id = n.getAttribute('r:embed') || n.getAttribute('r:id') || n.getAttribute('r:pict');
            var inFallback = false;
            for (var fb = n.parentNode; fb; fb = fb.parentNode) if (fb.nodeName === 'mc:Fallback') { inFallback = true; break; }
            if (id && rid[id] && !inFallback) {
              var ex = null;
              for (var an = n.parentNode; an && !ex; an = an.parentNode) {
                if (an.nodeName === 'wp:inline' || an.nodeName === 'wp:anchor') ex = an.getElementsByTagName('wp:extent')[0] || null;
              }
              toks.push({ target: rid[id], ew: ex ? Math.round(Number(ex.getAttribute('cx')) / 9525) : 0, eh: ex ? Math.round(Number(ex.getAttribute('cy')) / 9525) : 0 });
            }
          }
        });
        var ppr = kids(p, 'w:pPr')[0];
        var sectAfter = !!(ppr && kids(ppr, 'w:sectPr').length);
        var numPr = ppr && kids(ppr, 'w:numPr')[0];
        var isLi = !!numPr;
        var numId = numPr && numPr.getElementsByTagName('w:numId')[0];
        var liType = numId && numFmt[numId.getAttribute('w:val')] === 'decimal' ? 'ol' : 'li';
        var style = ppr && kids(ppr, 'w:pStyle')[0];
        var isHeadStyle = !!(style && /^Heading|^Title/i.test(style.getAttribute('w:val') || ''));
        var T = toks.filter(function (k) { return k.text != null; }).map(function (k) { return k.text; }).join('').replace(/\s+/g, ' ').trim();
        var hasImg = toks.some(function (k) { return k.target; });
        var spaceHere = null;
        toks.forEach(function (k) { if (k.text != null && !spaceHere) spaceHere = matchSpace(k.text); });
        var started = newSect;
        newSect = sectAfter;

        if (mode === 'cover') {
          if (T && isIntro(T)) { mode = 'intro'; return; }
          if (T) cover.push(T);
          return;
        }
        if (T && isConcl(T) && mode !== 'cover') { mode = 'conclusiones'; conclTitulo = T.replace(/^\s*\d+[.)]\s*/, ''); return; }
        if ((mode === 'intro') && T && isPlano(T)) { mode = 'planos'; planosTitulo = T.replace(/^\s*\d+[.)]\s*/, ''); return; }
        if ((mode === 'planos' || mode === 'imagenes') && T && T.length <= 60 && /^\s*(\d+[.)]\s*)?presupuesto\b/i.test(T) && !spaceHere) { mode = 'presupuesto'; return; }
        if ((mode === 'intro' || mode === 'planos' || mode === 'imagenes') && spaceHere) { mode = 'imagenes'; cur = spaceHere; if (!espacios[cur]) espacios[cur] = []; }
        else if (mode === 'imagenes' && started && photos > 0 && hasImg && !spaceHere) { mode = 'presupuesto'; }

        if (mode === 'intro') {
          if (T && !isHeadStyle) intro.push({ t: isLi ? liType : 'p', x: T });
        } else if (mode === 'planos') {
          toks.forEach(function (k) {
            if (k.target) planos.push(img(k.target, k.ew, k.eh));
          });
          if (T) planos.push({ t: isLi ? liType : (T === T.toUpperCase() && T.length <= 40 ? 'h' : 'p'), x: T, _at: toks.length });
        } else if (mode === 'imagenes') {
          toks.forEach(function (k) {
            if (k.target && cur) { espacios[cur].push(img(k.target)); photos++; }
          });
        } else if (mode === 'presupuesto') {
          toks.forEach(function (k) { if (k.target && !presu.img) presu.img = img(k.target); });
          if (T) presu.nota = presu.nota ? presu.nota + ' ' + T : T;
        } else if (mode === 'conclusiones') {
          toks.forEach(function (k) { if (k.target) { var fi = img(k.target); if (!firma) firma = fi; firmasAll.push(fi); } });
          if (T && !isHeadStyle) concl.push({ t: isLi ? liType : 'p', x: T });
        }
      });

      if (!presu.img && mode === 'imagenes') {
        var all = [];
        Object.keys(espacios).forEach(function (k) { espacios[k].forEach(function (im) { all.push({ k: k, im: im }); }); });
        var last = all[all.length - 1];
        if (last && last.im.ext === 'png') { presu.img = last.im; espacios[last.k].pop(); }
      }
      Object.keys(espacios).forEach(function (k) { if (!espacios[k].length) delete espacios[k]; });

      var cache = {};
      return Promise.all(imgJobs.map(function (o) {
        if (!cache[o.target]) {
          var f = zip.file('word/' + o.target.replace(/^\.?\//, ''));
          cache[o.target] = f ? f.async('uint8array') : Promise.resolve(null);
        }
        return cache[o.target].then(function (b) {
          o.bytes = b;
          var sz = b && imgSize(b);
          if (sz && sz.w > 0 && sz.h > 0) { o.w = sz.w; o.h = sz.h; } else if (o.ew && o.eh) { o.w = o.ew; o.h = o.eh; }
        });
      })).then(function () {
        function ok(o) { return o && o.bytes && /^(png|jpe?g|gif|bmp|webp)$/.test(o.ext); }
        var droppedPlanos = planos.filter(function (b) { return b.t === 'img' && !ok(b); }).map(function (b) { return b.ext || 'desconocido'; });
        planos = planos.filter(function (b) { return b.t !== 'img' || ok(b); });
        var maxArea = 0;
        planos.forEach(function (b) { if (b.t === 'img') maxArea = Math.max(maxArea, b.w * b.h); });
        var plans = planos.filter(function (b) { return b.t === 'img' && (maxArea === 0 || b.w * b.h >= maxArea * 0.35); });
        plans.forEach(function (b, i) { b.role = i === 0 ? 'grupo' : i === 1 ? 'individual' : 'fijo'; });
        planos.forEach(function (b) { if (b.t === 'img' && !b.role) b.role = 'fijo'; });
        if (!planos.some(function (b) { return b.role === 'individual'; })) {
          var gi = -1;
          planos.forEach(function (b, i) { if (b.role === 'grupo') gi = i; });
          var lastImg = -1;
          planos.forEach(function (b, i) { if (b.t === 'img') lastImg = i; });
          planos.splice(Math.max(gi, lastImg) + 1, 0, { t: 'img', role: 'individual', bytes: null, w: 0, h: 0, ext: '' });
        }
        if (!planos.some(function (b) { return b.role === 'grupo'; })) planos.unshift({ t: 'img', role: 'grupo', bytes: null, w: 0, h: 0, ext: '' });
        Object.keys(espacios).forEach(function (k) { espacios[k] = espacios[k].filter(ok); });
        if (!ok(presu.img)) presu.img = null;
        if (!ok(firma)) firma = null;
        firmasAll = firmasAll.filter(ok);

        var cv = { subtitulo: '', apartamento: '', conjunto: '', elaboradores: '', fecha: '' };
        var rest = cover.slice();
        function take(re) {
          for (var i = 0; i < rest.length; i++) if (re.test(rest[i])) return rest.splice(i, 1)[0];
          return '';
        }
        cv.apartamento = take(/^\s*(apartamento|apto|casa|local|oficina)\b/i);
        cv.conjunto = take(/^\s*(conjunto|edificio|urbanizaci|torres|condominio|residencial|agrupaci)/i);
        var ei = -1;
        rest.forEach(function (s, i) { if (ei < 0 && /^\s*elaborad[oa]\s+por/i.test(s)) ei = i; });
        var MON = /(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)/i;
        var fi = -1;
        rest.forEach(function (s, i) { if (MON.test(s) && /(19|20)\d\d/.test(s)) fi = i; });
        if (fi > -1) cv.fecha = rest[fi];
        if (ei > -1) cv.elaboradores = rest.filter(function (s, i) { return i > ei && i !== fi; }).join('\n');
        var used = {};
        if (ei > -1) used[ei] = 1;
        if (fi > -1) used[fi] = 1;
        var sub = rest.filter(function (s, i) { return !used[i] && (ei < 0 || i < ei); });
        cv.subtitulo = sub[0] || '';

        var vars = { seguro: '', poliza: '', siniestro: '', valor: '' };
        var blocks = intro.concat(concl);
        function scan(re) {
          for (var i = 0; i < blocks.length; i++) { var m = re.exec(blocks[i].x); if (m) return m[1].trim(); }
          return '';
        }
        vars.seguro = scan(/\bante\s+(.+?),\s+en el marco/i);
        vars.poliza = scan(/p[oó]liza\s+(?:No\.?|n[uú]mero|N°)\s*([\w\-.]+)/i).replace(/[.,;]+$/, '');
        vars.siniestro = scan(/siniestro\s+(?:No\.?|n[uú]mero|N°)\s*([\w\-.]+)/i).replace(/[.,;]+$/, '');
        vars.valor = scan(/asciende a\s+(.+?),\s+e incluye/i);
        function tag(v, name) {
          if (!v) return;
          blocks.forEach(function (b) { b.x = b.x.split(v).join('{{' + name + '}}'); });
        }
        tag(vars.seguro, 'SEGURO');
        tag(vars.siniestro, 'SINIESTRO');
        tag(vars.poliza, 'POLIZA');
        tag(vars.valor, 'VALOR');

        planos.forEach(function (b) { delete b._at; });
        return { droppedPlanos: droppedPlanos, cover: cv, vars: vars, intro: intro, planosTitulo: planosTitulo, planos: planos, espacios: espacios, presupuesto: presu, concl: concl, conclTitulo: conclTitulo, firma: firma, firmas: firmasAll };
      });
    });
  });
}
if (typeof module !== 'undefined') module.exports = { parseInforme: parseInforme, SPACES: SPACES };
