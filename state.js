const has = (v) => !!String(v == null ? '' : v).trim();

function normText(x) { return String(x || '').toLowerCase().replace(/\{\{[a-z]+\}\}/g, '{}').replace(/\[[^\]]*\]/g, '[]').replace(/[\u2018\u2019\u201c\u201d]/g, '"').replace(/\s+/g, ' ').trim(); }
function blocksKey(bl) { return (bl || []).map((b) => (b.t || 'p') + '|' + normText(b.x)).filter((s) => s.length > 2).join('\n'); }
function formatoDiffs(d, p) {
  const f = p && p.formato;
  const out = [];
  if (!f) return out;
  if (blocksKey(d.intro) !== blocksKey(f.intro)) out.push('intro');
  if (blocksKey(d.concl) !== blocksKey(f.concl)) out.push('concl');
  return out;
}

function moduleGaps(d, project) {
  const g = [];
  const c = d.cover || {};
  const v = d.vars || {};
  if (!has(c.apartamento) || !has(c.conjunto)) g.push('portada');
  if (!(d.intro || []).length || !has(v.seguro) || !has(v.poliza) || !has(v.siniestro)) g.push('introducción');
  else if (formatoDiffs(d, project).indexOf('intro') > -1) g.push('introducción distinta al formato');
  const grp = project && (project.grupos || []).filter((x) => x.id === d.grupoId)[0];
  const ind = (d.planos || []).filter((b) => b.t === 'img' && b.role === 'individual')[0];
  if (!(grp && grp.aid) || !(ind && ind.ref)) g.push('planos');
  if (!Object.keys(d.espacios || {}).length) g.push('imágenes');
  if (!(project && project.presupuestoObligatorio === false) && !has(d.presupuesto && d.presupuesto.nota)) g.push('presupuesto');
  if (!(d.concl || []).length || !has(v.valor)) g.push('conclusiones');
  else if (formatoDiffs(d, project).indexOf('concl') > -1) g.push('conclusiones distintas al formato');
  return g;
}

function cardState(d, project) {
  if (d.importError || moduleGaps(d, project).length) return 'rojo';
  return d.completado ? 'verde' : 'amarillo';
}

function apartamentoFrom(meta, groupName) {
  const m = meta || {};
  const file = String(m.name || '').replace(/\.docx$/i, '');
  const parts = String(m.path || '').split(' / ');
  const folder = parts.length ? parts[parts.length - 1] : '';
  let num = '';
  let x = /(?:^|[^0-9])(\d{2,4})(?![0-9])/.exec(file.replace(/^\s*inf[\s\-_.]*/i, ''));
  if (x) num = x[1];
  if (!num) { x = /(?:^|[^0-9])(\d{2,4})(?![0-9])/.exec(folder); if (x) num = x[1]; }
  if (!num) return '';
  let tower = '';
  const t = /torre\s*([A-Za-z0-9]+)/i.exec(String(groupName || ''));
  if (t) tower = t[1].toUpperCase();
  if (!tower) { const y = /\d{2,4}\s*[-_ ]?\s*([A-Za-z])(?![A-Za-z0-9])/.exec(file) || /\d{2,4}\s*[-_ ]?\s*([A-Za-z])(?![A-Za-z0-9])/.exec(folder); if (y) tower = y[1].toUpperCase(); }
  return 'APARTAMENTO ' + num + (tower ? ' TORRE ' + tower : '');
}

function fillApartamento(doc, project) {
  if (!doc || doc.importError || !doc.cover || has(doc.cover.apartamento)) return false;
  const g = ((project && project.grupos) || []).filter((x) => x.id === doc.grupoId)[0];
  const meta = doc.drive || { name: doc.source };
  const v = apartamentoFrom(meta, g && g.name);
  if (!v) return false;
  doc.cover.apartamento = v;
  return true;
}

module.exports = { apartamentoFrom: apartamentoFrom, fillApartamento: fillApartamento, moduleGaps: moduleGaps, cardState: cardState, has: has };
