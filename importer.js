const crypto = require('crypto');
const sharp = require('sharp');
const JSZip = require('jszip');
const { JSDOM } = require('jsdom');
const { parseInforme } = require('./parse');
const { has, apartamentoFrom, fillApartamento } = require('./state');

const DOMParserCls = new JSDOM('').window.DOMParser;
const locks = {};

function uid() { return crypto.randomBytes(8).toString('hex'); }
function sha(bytes) { return crypto.createHash('sha256').update(bytes).digest('hex'); }

function withLock(key, fn) {
  const prev = locks[key] || Promise.resolve();
  const next = prev.then(fn, fn);
  locks[key] = next.catch(() => {});
  return next;
}

async function processImage(bytes, maxDim) {
  const out = await sharp(Buffer.from(bytes), { failOn: 'none' })
    .rotate()
    .resize({ width: maxDim, height: maxDim, fit: 'inside', withoutEnlargement: true })
    .flatten({ background: '#ffffff' })
    .jpeg({ quality: 90 })
    .toBuffer({ resolveWithObject: true });
  return { buf: out.data, w: out.info.width, h: out.info.height };
}

async function pool(items, size, fn) {
  const out = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(size, items.length) }, async () => {
    while (next < items.length) {
      const k = next++;
      out[k] = await fn(items[k], k);
    }
  });
  await Promise.all(workers);
  return out;
}

function fillDefaults(project, res) {
  const d = project.defaults = project.defaults || {};
  const c = res.cover || {};
  const v = res.vars || {};
  const src = { conjunto: c.conjunto, subtitulo: c.subtitulo, elaboradores: c.elaboradores, seguro: v.seguro, poliza: v.poliza, siniestro: v.siniestro };
  Object.keys(src).forEach((k) => { if (has(src[k]) && !has(d[k])) d[k] = src[k]; });
}

async function seedFirma(store, res) {
  if (!res.firma) return false;
  return withLock('config', async () => {
    const cur = await store.getDoc('config', 'global');
    if (cur && cur.firma && cur.firma.aid) return false;
    const pr = await processImage(res.firma.bytes, 1600);
    const aid = await store.putAsset(pr.buf, 'image/jpeg');
    await store.putDoc('config', 'global', { firma: { aid: aid, w: pr.w, h: pr.h }, updatedAt: Date.now() });
    return true;
  });
}

async function importDocx(store, buf, o) {
  const warns = [];
  const res = await parseInforme(buf, JSZip, DOMParserCls);
  const p0 = await store.getDoc('proyectos', o.projectId);
  if (!p0) throw new Error('Proyecto no encontrado');
  const groups = p0.grupos || [];
  const knownFijos = p0.fijos || {};
  const fx = { fijos: {}, group: null, newGroup: null };
  const tg = groups.filter((x) => x.id === o.groupId)[0] || null;

  async function up(img, maxDim) {
    try {
      const pr = await processImage(img.bytes, maxDim);
      const aid = await store.putAsset(pr.buf, 'image/jpeg');
      return { aid: aid, w: pr.w, h: pr.h, cap: '' };
    } catch (e) {
      warns.push('una imagen no se pudo procesar (' + (img.ext || 'formato desconocido') + ')');
      return null;
    }
  }

  async function shared(img, maxDim) {
    const key = sha(img.bytes);
    const k = knownFijos[key] || fx.fijos[key];
    if (k) return { aid: k.aid, w: k.w, h: k.h, cap: '' };
    const r = await up(img, maxDim);
    if (r) fx.fijos[key] = { aid: r.aid, w: r.w, h: r.h };
    return r;
  }

  const tm = /TORRE\s+[A-Z0-9]+/i.exec((res.cover && res.cover.apartamento) || '');
  let grupoId = tg ? tg.id : '';
  const planos = [];
  for (const b of res.planos) {
    if (b.t !== 'img') { planos.push({ t: b.t, x: b.x }); continue; }
    if (!b.bytes) { planos.push({ t: 'img', role: b.role, ref: null }); continue; }
    if (b.role === 'grupo') {
      const key = sha(b.bytes);
      let g = tg || groups.filter((x) => x.hash === key)[0] || null;
      if (!g || !g.aid) {
        const r = await up(b, 2200);
        if (r) {
          if (g) fx.group = { id: g.id, aid: r.aid, w: r.w, h: r.h, hash: key };
          else {
            const nm = tm ? 'Torre ' + tm[0].split(/\s+/)[1].toUpperCase() : 'Grupo ' + (groups.length + 1);
            g = { id: uid(), name: nm, aid: r.aid, w: r.w, h: r.h, hash: key };
            fx.newGroup = g;
          }
        } else warns.push('el plano de grupo no se pudo procesar');
      }
      if (g) grupoId = g.id;
      planos.push({ t: 'img', role: 'grupo', ref: null });
    } else {
      planos.push({ t: 'img', role: b.role, ref: b.role === 'fijo' ? await shared(b, 2200) : await up(b, 2200) });
    }
  }

  const espacios = {};
  for (const name of Object.keys(res.espacios)) {
    const refs = (await pool(res.espacios[name], 3, (img) => up(img, 1600))).filter(Boolean);
    if (refs.length) espacios[name] = refs;
  }

  const fm = p0.formato || null;
  if (fm && fm.planos) {
    const hasExtras = planos.some((b) => b.t !== 'img' || (b.role === 'fijo' && b.ref));
    if (!hasExtras) {
      const keep = {};
      planos.forEach((b) => { if (b.t === 'img') keep[b.role] = b; });
      const rebuilt = fm.planos.map((b) => b.t === 'img' ? (b.role === 'fijo' ? { t: 'img', role: 'fijo', ref: b.ref || null } : (keep[b.role] || { t: 'img', role: b.role, ref: null })) : { t: b.t, x: b.x });
      ['grupo', 'individual'].forEach((role) => { if (!rebuilt.some((b) => b.t === 'img' && b.role === role)) { const it = keep[role] || { t: 'img', role: role, ref: null }; if (role === 'grupo') rebuilt.unshift(it); else rebuilt.push(it); } });
      planos.length = 0;
      rebuilt.forEach((b) => planos.push(b));
    }
  }
  const presRef = res.presupuesto.img ? await up(res.presupuesto.img, 2200) : null;
  const firma = res.firma ? await shared(res.firma, 1600) : null;

  await withLock('p:' + o.projectId, async () => {
    const d = await store.getDoc('proyectos', o.projectId);
    if (!d) return;
    d.grupos = d.grupos || [];
    d.fijos = d.fijos || {};
    Object.assign(d.fijos, fx.fijos);
    if (fx.newGroup && !d.grupos.some((x) => x.id === fx.newGroup.id)) d.grupos.push(fx.newGroup);
    if (fx.group) {
      const gg = d.grupos.filter((x) => x.id === fx.group.id)[0];
      if (gg && !gg.aid) { gg.aid = fx.group.aid; gg.w = fx.group.w; gg.h = fx.group.h; gg.hash = fx.group.hash; }
    }
    fillDefaults(d, res);
    d.updatedAt = Date.now();
    await store.putDoc('proyectos', o.projectId, d);
  });

  const missing = [];
  if (!Object.keys(espacios).length) missing.push('no se encontraron fotos por espacio');
  if (!has(res.presupuesto.nota) && !(fm && has(fm.presupuestoNota))) missing.push('no se encontró el texto del presupuesto');
  if (!has(res.vars && res.vars.valor)) missing.push('no se detectó el valor total');

  const now = Date.now();
  const id = uid();
  const doc = {
    v: 2, projectId: o.projectId, createdAt: now, updatedAt: now, step: 'portada', source: o.name || '',
    cover: res.cover, vars: res.vars, intro: res.intro, planosTitulo: res.planosTitulo || 'PLANO CON EL REGISTRO DE LAS AFECTACIONES',
    grupoId: grupoId, planos: planos, espacios: espacios, orden: Object.keys(espacios), presupuesto: { ref: presRef, nota: res.presupuesto.nota || (fm && fm.presupuestoNota) || '' },
    concl: res.concl, conclTitulo: res.conclTitulo || 'CONCLUSIONES', firma: firma,
    completado: false, drive: o.drive || null, importError: ''
  };
  {
    const ng = (p0.grupos || []).concat(fx.newGroup ? [fx.newGroup] : []);
    fillApartamento(doc, { grupos: ng });
    if (!doc.drive && !has(doc.cover.apartamento)) doc.cover.apartamento = apartamentoFrom({ name: o.name }, (ng.filter((x) => x.id === grupoId)[0] || {}).name) || '';
  }
  await store.putDoc('informes2', id, doc);
  try { await seedFirma(store, res); } catch (e) {}
  const infos = [];
  if (!res.presupuesto.img) infos.push('Sin imagen de presupuesto. Es opcional: solo el texto de la sección es obligatorio.');
  return { id: id, doc: doc, warns: warns.concat(missing), infos: infos };
}

async function recordFailure(store, o, message) {
  const now = Date.now();
  const id = uid();
  const doc = {
    v: 2, projectId: o.projectId, createdAt: now, updatedAt: now, step: 'portada', source: o.name || '',
    cover: { subtitulo: '', apartamento: o.name || '', conjunto: '', elaboradores: '', fecha: '' },
    vars: { seguro: '', poliza: '', siniestro: '', valor: '' }, intro: [], planosTitulo: 'PLANO CON EL REGISTRO DE LAS AFECTACIONES',
    grupoId: o.groupId || '', planos: [], espacios: {}, presupuesto: { ref: null, nota: '' }, concl: [], conclTitulo: 'CONCLUSIONES', firma: null,
    completado: false, drive: o.drive || null, importError: message
  };
  await store.putDoc('informes2', id, doc);
  return { id: id, doc: doc };
}

async function extractFirma(store, buf) {
  if (!buf || buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) { const e = new Error('Ese archivo no es un .docx. Si es .doc o PDF, guárdalo como .docx desde Word.'); throw e; }
  let res;
  try { res = await parseInforme(buf, JSZip, DOMParserCls); }
  catch (e) { throw new Error('No se pudo leer el Word: ' + String((e && e.message) || 'formato no reconocido') + '.'); }
  if (!res.firma) return null;
  const pr = await processImage(res.firma.bytes, 1600);
  return { aid: await store.putAsset(pr.buf, 'image/jpeg'), w: pr.w, h: pr.h };
}

async function backfillApartamentos(store) {
  const projects = {};
  let n = 0;
  for (const r of await store.listCol('informes2')) {
    const d = r.data;
    if (!d || d.importError || has(d.cover && d.cover.apartamento)) continue;
    if (!projects[d.projectId]) projects[d.projectId] = await store.getDoc('proyectos', d.projectId);
    if (fillApartamento(d, projects[d.projectId])) { d.updatedAt = Date.now(); await store.putDoc('informes2', r.id, d); n++; }
  }
  return n;
}

module.exports = { backfillApartamentos: backfillApartamentos, extractFirma: extractFirma, importDocx: importDocx, recordFailure: recordFailure, withLock: withLock, uid: uid };
