const D = require('docx');
const buildReport = require('./gen');
const { SPACES } = require('./parse');
const { has } = require('./state');
const { withLock } = require('./importer');

const jobs = {};
const NOTA_PRESUPUESTO = 'Nota: Los items cuya cantidad se encuentra en cero, corresponde a que no se requiere.';

function blank() { return { busy: false, msg: '', done: 0, total: 0, summary: '', errs: [] }; }

function status(projectId) { return jobs[projectId] || blank(); }

function pending(doc) {
  return !!(doc && doc.completado && !doc.importError && doc.drive && doc.drive.folderId && (!doc.exportado || (doc.updatedAt || 0) > (doc.exportado.at || 0)));
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
  const g = ((project && project.grupos) || []).filter((x) => x.id === doc.grupoId)[0];
  const groupRef = g && g.aid ? { aid: g.aid, w: g.w, h: g.h } : null;
  const planos = [];
  for (const b of doc.planos || []) {
    if (b.t === 'img') planos.push({ t: 'img', role: b.role, img: await asImg(store, b.role === 'grupo' ? groupRef : b.ref) });
    else planos.push({ t: b.t, x: b.x });
  }
  const esp = doc.espacios || {};
  const names = SPACES.filter((n) => (esp[n] || []).length).concat(Object.keys(esp).filter((n) => SPACES.indexOf(n) < 0 && esp[n].length));
  const espacios = [];
  for (const n of names) {
    const photos = [];
    for (const r of esp[n]) { const im = await asImg(store, r, r.cap); if (im) photos.push(im); }
    espacios.push({ name: n, photos: photos });
  }
  const presImg = await asImg(store, doc.presupuesto && doc.presupuesto.ref);
  const firmaRef = cfg && cfg.firma && cfg.firma.aid ? cfg.firma : doc.firma;
  return {
    cover: doc.cover, vars: doc.vars, intro: doc.intro, planosTitulo: doc.planosTitulo, planos: planos, espacios: espacios,
    presupuesto: presupuestoFor(project, doc, presImg), estiloPropio: !!(project && project.estiloPropio),
    concl: doc.concl, conclTitulo: doc.conclTitulo, firma: await asImg(store, firmaRef)
  };
}

async function buildDocx(store, doc, project, cfg) {
  return D.Packer.toBuffer(buildReport(D, await buildRep(store, doc, project, cfg)));
}

function copyName(name) { return String(name || 'Informe').replace(/\.docx$/i, '').trim() + ' (Corregido).docx'; }

async function sendOne(ctx, project, cfgDoc, r) {
  const store = ctx.store;
  const buf = await buildDocx(store, r.data, project, cfgDoc);
  const name = copyName(r.data.drive.name);
  const prev = r.data.exportado && r.data.exportado.fileId;
  let file;
  let updated = false;
  if (prev) {
    try { file = await ctx.drive.upload({ parentId: r.data.drive.folderId, name: name, buf: buf, existingId: prev }); updated = true; }
    catch (e) { if (e.status === 404) file = await ctx.drive.upload({ parentId: r.data.drive.folderId, name: name, buf: buf }); else throw e; }
  } else {
    file = await ctx.drive.upload({ parentId: r.data.drive.folderId, name: name, buf: buf });
  }
  await withLock('r:' + r.id, async () => {
    const fresh = await store.getDoc('informes2', r.id);
    if (!fresh) return;
    fresh.exportado = { fileId: file.id, name: file.name || name, at: Date.now() };
    await store.putDoc('informes2', r.id, fresh);
  });
  return { name: file.name || name, updated: updated };
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
  const all = (await store.listCol('informes2')).filter((r) => r.data.projectId === projectId && (!groupId || (r.data.grupoId || '') === groupId) && r.data.completado && !r.data.importError);
  const todo = all.filter((r) => pending(r.data));
  const noDrive = all.filter((r) => !(r.data.drive && r.data.drive.folderId)).length;
  const upToDate = all.length - todo.length - noDrive;
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
  parts.push(created + (created === 1 ? ' copia creada' : ' copias creadas'));
  if (updated) parts.push(updated + (updated === 1 ? ' actualizada' : ' actualizadas'));
  if (upToDate > 0) parts.push(upToDate + ' ya estaban al día');
  if (noDrive) parts.push(noDrive + (noDrive === 1 ? ' no viene' : ' no vienen') + ' de Drive y se omitió');
  job.summary = todo.length || noDrive || upToDate ? 'Exportación terminada: ' + parts.join(', ') + '.' : 'No hay informes completados para exportar.';
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
