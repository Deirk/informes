const D = require('docx');
const buildReport = require('./gen');
const { SPACES } = require('./parse');
const { has, cardState, moduleGaps, planoGrupoOf, fijaOf } = require('./state');
const { withLock } = require('./importer');

const jobs = {};
const NOTA_PRESUPUESTO = 'Nota: Los items cuya cantidad se encuentra en cero, corresponde a que no se requiere.';

function blank() { return { busy: false, msg: '', done: 0, total: 0, summary: '', errs: [] }; }

function status(projectId) { return jobs[projectId] || blank(); }

function pending(doc, project) {
  return !!(doc && !doc.importError && doc.drive && doc.drive.folderId && (!doc.exportado || (doc.updatedAt || 0) > (doc.exportado.at || 0)));
}

async function asImg(store, ref, caption) {
  if (!ref || !ref.aid) return null;
  const a = await store.getAsset(ref.aid);
  if (!a) return null;
  return { data: a.buf, w: ref.w || 1, h: ref.h || 1, caption: caption || '' };
}

function presupuestoFor(project, doc, img) {
  let nota = (doc.presupuesto && doc.presupuesto.nota) || '';
  if (project && project.presupuestoObligatorio === false && !img && !has(nota)) nota = NOTA_PRESUPUESTO;
  return { img: img, nota: nota };
}

async function buildRep(store, doc, project, cfg) {
  const groupRef = planoGrupoOf(project, doc);
  const fija = fijaOf(project, doc);
  const blocks = (doc.planos || []).slice();
  if (fija && !blocks.some((b) => b.t === 'img' && b.role === 'fijo')) {
    const gi = blocks.findIndex((b) => b.t === 'img' && b.role === 'grupo');
    blocks.splice(gi < 0 ? 0 : gi, 0, { t: 'img', role: 'fijo', ref: null });
  }
  const planos = [];
  let fijoN = 0;
  for (const b of blocks) {
    if (b.t === 'img') {
      let ref = b.role === 'grupo' ? groupRef : b.ref;
      if (b.role === 'fijo') { if (fijoN === 0 && fija) ref = fija; fijoN++; }
      planos.push({ t: 'img', role: b.role, img: await asImg(store, ref) });
    } else planos.push({ t: b.t, x: b.x });
  }
  const esp = doc.espacios || {};
  const has2 = (n) => (esp[n] || []).length;
  const ord = (doc.orden || []).filter(has2);
  const names = ord.concat(SPACES.filter((n) => has2(n) && ord.indexOf(n) < 0)).concat(Object.keys(esp).filter((n) => has2(n) && ord.indexOf(n) < 0 && SPACES.indexOf(n) < 0));
  const espacios = [];
  for (const n of names) {
    const photos = [];
    for (const r of esp[n]) { const im = await asImg(store, r, r.cap); if (im) photos.push(im); }
    espacios.push({ name: n, photos: photos });
  }
  if (project && project.fotosPorEspacio === false) {
    const all = [];
    espacios.forEach((x) => x.photos.forEach((ph) => all.push(ph)));
    espacios.length = 0;
    if (all.length) espacios.push({ name: 'FOTOS', general: true, photos: all });
  }
  const presImg = await asImg(store, doc.presupuesto && doc.presupuesto.ref);
  const memoria = [];
  for (const r of doc.memoria || []) { const im = await asImg(store, r); if (im) memoria.push(im); }
  const firmaRef = cfg && cfg.firma && cfg.firma.aid ? cfg.firma : doc.firma;
  return {
    cover: doc.cover, vars: doc.vars, intro: doc.intro, planosTitulo: doc.planosTitulo, planos: planos, espacios: espacios, memoria: memoria,
    presupuesto: presupuestoFor(project, doc, presImg), estiloPropio: !!(project && project.estiloPropio),
    concl: doc.concl, conclTitulo: doc.conclTitulo, firma: await asImg(store, firmaRef)
  };
}

async function buildDocx(store, doc, project, cfg) {
  return D.Packer.toBuffer(buildReport(D, await buildRep(store, doc, project, cfg)));
}

function copyName(name) { return String(name || 'Informe').replace(/\.docx$/i, '').trim() + ' (Corregido).docx'; }

function slug(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function photoPrefix(doc, project) {
  const segs = String((doc.drive && doc.drive.path) || '').split('/').map((x) => x.trim()).filter(Boolean);
  const proj = slug((project && project.name) || segs[0] || '');
  const grp = ((project && project.grupos) || []).filter((x) => x.id === doc.grupoId)[0];
  const torre = slug(segs.filter((x) => /torre/i.test(x))[0] || (grp && grp.name) || '');
  const rango = segs.map((x) => /apartamentos?\s*0*(\d+)/i.exec(x)).filter(Boolean)[0];
  const apSrc = String((doc.cover && doc.cover.apartamento) || (doc.drive && doc.drive.name) || '');
  const ap = (/(\d{3,4})/.exec(apSrc) || [])[1] || '';
  const torreNum = (/(\d+)\s*$/.exec(torre.replace(/_/g, ' ')) || [])[1];
  const parts = [proj, 'inspeccion', 'zonas_privadas', /^torre/.test(torre) || !torre ? torre : 'torre_' + (torreNum || torre), rango ? 'aptos_' + String(rango[1]).padStart(2, '0') : '', ap];
  return parts.filter(Boolean).join('_');
}

function stamp(t) {
  const f = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Bogota', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(new Date(t));
  const g = (k) => f.filter((x) => x.type === k)[0].value;
  return g('year') + '-' + g('month') + '-' + g('day') + '_' + g('hour') + '-' + g('minute') + '-' + g('second');
}

async function sendPhotos(ctx, doc, project) {
  const sent = Object.assign({}, doc.fotosDrive || {});
  const esp = doc.espacios || {};
  const ord = (doc.orden || []).filter((n) => (esp[n] || []).length);
  const names = ord.concat(Object.keys(esp).filter((n) => (esp[n] || []).length && ord.indexOf(n) < 0));
  const refs = [];
  names.forEach((n) => esp[n].forEach((r) => { if (r && r.aid) refs.push(r); }));
  let next = Object.keys(sent).reduce((m, k) => Math.max(m, sent[k].n || 0), 0);
  const prefix = photoPrefix(doc, project);
  const at = stamp(Date.now());
  let err = '';
  for (const r of refs) {
    if (sent[r.aid]) continue;
    try {
      const a = await ctx.store.getAsset(r.aid);
      if (!a) continue;
      const png = /png/i.test(a.type);
      const n = next + 1;
      const fname = prefix + '_' + at + '_' + n + (png ? '.png' : '.jpg');
      const up = await ctx.drive.upload({ parentId: doc.drive.folderId, name: fname, buf: a.buf, mime: png ? 'image/png' : 'image/jpeg' });
      next = n;
      sent[r.aid] = { id: up.id, name: fname, n: n };
    } catch (e) {
      err = (e && e.message) || 'no se pudo subir una foto';
      break;
    }
  }
  return { sent: sent, err: err };
}

async function sendOne(ctx, project, cfgDoc, r) {
  const store = ctx.store;
  const buf = await buildDocx(store, r.data, project, cfgDoc);
  const file = await ctx.drive.replace({ fileId: r.data.drive.id, buf: buf });
  const name = file.name || r.data.drive.name;
  const ph = project && project.fotosADrive === false ? { sent: r.data.fotosDrive || {}, err: '' } : await sendPhotos(ctx, r.data, project);
  await withLock('r:' + r.id, async () => {
    const fresh = await store.getDoc('informes2', r.id);
    if (!fresh) return;
    fresh.exportado = { fileId: file.id, name: name, at: Date.now() };
    fresh.fotosDrive = ph.sent;
    if (file.modifiedTime && fresh.drive) fresh.drive.modified = file.modifiedTime;
    await store.putDoc('informes2', r.id, fresh);
  });
  if (ph.err) throw new Error('El Word se envió, pero una foto no se pudo subir: ' + ph.err);
  return { name: name, updated: !!(r.data.exportado && r.data.exportado.fileId) };
}

async function exportOne(ctx, reportId) {
  const doc = await ctx.store.getDoc('informes2', reportId);
  if (!doc) throw new Error('Informe no encontrado');
  if (doc.importError) throw new Error('Este informe tiene un error de importación y no se puede enviar.');
  if (!doc.drive || !doc.drive.folderId) throw new Error('Este informe no tiene ruta de Drive. Sincroniza el proyecto con Drive para obtenerla.');
  const project = await ctx.store.getDoc('proyectos', doc.projectId);
  const cfgDoc = await ctx.store.getDoc('config', 'global');
  const res = await sendOne(ctx, project, cfgDoc, { id: reportId, data: doc });
  ctx.notify('informes2');
  return res;
}

async function run(ctx, projectId, job, groupId) {
  const store = ctx.store;
  const project = await store.getDoc('proyectos', projectId);
  if (!project) throw new Error('Proyecto no encontrado');
  const cfgDoc = await store.getDoc('config', 'global');
  const all = (await store.listCol('informes2')).filter((r) => r.data.projectId === projectId && (!groupId || (r.data.grupoId || '') === groupId) && !r.data.importError);
  const bad = [];
  const todo = all.filter((r) => pending(r.data, project));
  const noDrive = all.filter((r) => !(r.data.drive && r.data.drive.folderId)).length;
  const upToDate = Math.max(0, all.length - todo.length - noDrive - bad.length);
  job.total = todo.length;
  let created = 0;
  let updated = 0;
  for (let i = 0; i < todo.length; i++) {
    const r = todo[i];
    const label = (r.data.cover && r.data.cover.apartamento) || r.data.drive.name;
    job.done = i;
    job.msg = label;
    try {
      const res = await sendOne(ctx, project, cfgDoc, r);
      if (res.updated) updated++; else created++;
    } catch (e) {
      job.errs.push(label + ': ' + ((e && e.message) || 'no se pudo exportar'));
    }
    ctx.notify('informes2');
  }
  job.done = todo.length;
  const parts = [];
  parts.push((created + updated) + ((created + updated) === 1 ? ' original reemplazado' : ' originales reemplazados'));
  if (upToDate > 0) parts.push(upToDate + ' ya estaban al día');
  if (bad.length) parts.push(bad.length + (bad.length === 1 ? ' con pendientes se omitió' : ' con pendientes se omitieron'));
  if (noDrive) parts.push(noDrive + (noDrive === 1 ? ' no viene' : ' no vienen') + ' de Drive y se omitió');
  job.summary = todo.length || noDrive || upToDate || bad.length ? 'Exportación terminada: ' + parts.join(', ') + '.' : 'No hay informes para exportar.';
}

function start(ctx, projectId, groupId) {
  if (jobs[projectId] && jobs[projectId].busy) return jobs[projectId];
  const job = jobs[projectId] = blank();
  job.busy = true;
  job.gid = groupId || '';
  job.msg = 'Preparando…';
  run(ctx, projectId, job, groupId || '')
    .catch((e) => { job.errs.push((e && e.message) || 'No se pudo exportar.'); job.summary = ''; })
    .finally(() => { job.busy = false; ctx.notify('informes2'); });
  return job;
}

module.exports = { exportOne: exportOne, start: start, status: status, pending: pending, buildDocx: buildDocx, copyName: copyName };
