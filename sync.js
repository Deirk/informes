const { importDocx, recordFailure, withLock, uid, backfillApartamentos } = require('./importer');
const { cardState } = require('./state');

const jobs = {};

function blank() { return { busy: false, msg: '', done: 0, total: 0, summary: '', errs: [] }; }

function status(projectId) { return jobs[projectId] || blank(); }

function friendly(message) {
  const m = String(message || 'No se pudo importar');
  if (/zip|central directory|end of central/i.test(m)) return 'el archivo no es un documento de Word válido';
  return m;
}

async function run(ctx, projectId, job) {
  const store = ctx.store;
  const p0 = await store.getDoc('proyectos', projectId);
  if (!p0) throw new Error('Proyecto no encontrado');
  if (!p0.drive || !p0.drive.id) throw new Error('El proyecto no tiene carpeta de Drive vinculada');
  job.msg = 'Buscando informes en Drive…';
  const scan = await ctx.drive.scan(p0.drive.id);
  const reports = await store.listCol('informes2');
  const known = new Set(reports.filter((r) => r.data.projectId === projectId && r.data.drive && r.data.drive.id).map((r) => r.data.drive.id));
  const byDrive = {};
  scan.items.forEach((it) => { byDrive[it.id] = it; });
  for (const r of reports) {
    const dm = r.data.drive;
    if (r.data.projectId !== projectId || !dm || !dm.id || (dm.folderId && dm.path)) continue;
    const it = byDrive[dm.id];
    if (!it) continue;
    const fresh = await store.getDoc('informes2', r.id);
    if (fresh && fresh.drive && fresh.drive.id === dm.id) {
      fresh.drive.folderId = it.parentId || it.towerId;
      fresh.drive.path = it.path || it.towerName;
      await store.putDoc('informes2', r.id, fresh);
    }
  }
  const todo = scan.items.filter((it) => !known.has(it.id));
  const omitted = scan.items.length - todo.length;
  const towers = {};
  todo.forEach((it) => { towers[it.towerId] = it.towerName; });

  const gmap = {};
  await withLock('p:' + projectId, async () => {
    const d = await store.getDoc('proyectos', projectId);
    d.grupos = d.grupos || [];
    Object.keys(towers).forEach((tid) => {
      const nm = towers[tid];
      let g = d.grupos.filter((x) => x.driveId === tid)[0] || d.grupos.filter((x) => String(x.name || '').trim().toLowerCase() === String(nm).trim().toLowerCase())[0];
      if (g) g.driveId = tid;
      else { g = { id: uid(), name: nm, aid: null, w: 0, h: 0, hash: '', driveId: tid }; d.grupos.push(g); }
      gmap[tid] = g.id;
    });
    d.updatedAt = Date.now();
    await store.putDoc('proyectos', projectId, d);
  });

  job.total = todo.length;
  let amarillos = 0;
  let rojos = 0;
  for (let i = 0; i < todo.length; i++) {
    const it = todo[i];
    const groupId = gmap[it.towerId] || '';
    job.done = i;
    job.msg = it.towerName + ' · ' + it.name;
    const meta = { id: it.id, modified: it.modifiedTime || '', name: it.name, folderId: it.parentId || it.towerId, path: it.path || it.towerName };
    try {
      const buf = await ctx.drive.doc(it.id);
      const r = await importDocx(store, buf, { projectId: projectId, groupId: groupId, name: it.name, drive: meta });
      const project = await store.getDoc('proyectos', projectId);
      if (cardState(r.doc, project) === 'amarillo') amarillos++; else rojos++;
      r.warns.forEach((w) => job.errs.push(it.name + ': ' + w));
    } catch (e) {
      const m = friendly(e && e.message);
      rojos++;
      job.errs.push(it.name + ': ' + m);
      await recordFailure(store, { projectId: projectId, groupId: groupId, name: it.name, drive: meta }, 'No se pudo importar: ' + m);
    }
    ctx.notify('informes2');
    ctx.notify('proyectos');
    ctx.notify('config');
  }
  job.done = todo.length;
  try { await backfillApartamentos(store); } catch (e) {}
  ctx.notify('informes2');
  job.summary = 'Sincronizado. ' + scan.items.length + ' informes en Drive, ' + todo.length + ' nuevos (' + amarillos + ' con el mismo formato, ' + rojos + ' en rojo) y ' + omitted + ' que ya estaban.';
}

function start(ctx, projectId) {
  if (jobs[projectId] && jobs[projectId].busy) return jobs[projectId];
  const job = jobs[projectId] = blank();
  job.busy = true;
  job.msg = 'Iniciando…';
  run(ctx, projectId, job)
    .catch((e) => { job.errs.push((e && e.message) || 'No se pudo leer Drive.'); job.summary = ''; })
    .finally(() => { job.busy = false; ctx.notify('informes2'); ctx.notify('proyectos'); });
  return job;
}

module.exports = { start: start, status: status };
