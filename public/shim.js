(function () {
  window.HYA_LOCAL = true;
  const listeners = {};
  function emit(col) { (listeners[col] ? Array.from(listeners[col]) : []).forEach((fn) => fn()); }
  try {
    const es = new EventSource('/api/events');
    es.onmessage = (e) => { try { emit(JSON.parse(e.data).col); } catch (x) {} };
  } catch (e) {}
  setInterval(() => { if (!document.hidden) Object.keys(listeners).forEach(emit); }, 10000);
  async function jf(url, opts) {
    const r = await fetch(url, opts);
    if (r.status === 401) { location.href = '/login'; throw { code: 'unavailable', message: 'Sesión expirada' }; }
    if (!r.ok) { let m = ''; try { m = (await r.json()).error; } catch (e) {} throw { code: 'unavailable', message: m || ('HTTP ' + r.status) }; }
    return r.json();
  }
  function mkDoc(path) {
    const parts = path.split('/'), col = parts[0], id = parts[1];
    return {
      set: (d) => jf('/api/doc/' + col + '/' + id, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(d) }).then(() => undefined),
      delete: () => jf('/api/doc/' + col + '/' + id, { method: 'DELETE' }).then(() => undefined),
      get: async () => { const r = await fetch('/api/doc/' + col + '/' + id); if (!r.ok) return { id: id, exists: false, data: () => undefined }; const d = await r.json(); return { id: id, exists: true, data: () => d }; }
    };
  }
  const db = {
    doc: mkDoc,
    collection: (name) => {
      let ord = null, dir = 'asc', lim = null;
      const q = {
        orderBy: (f, d) => { ord = f; dir = d || 'asc'; return q; },
        limit: (n) => { lim = n; return q; },
        doc: (id) => mkDoc(name + '/' + id),
        onSnapshot: (cb, err) => {
          let last = null;
          const run = async () => {
            try {
              let rows = await jf('/api/col/' + name);
              const sig = JSON.stringify(rows); if (sig === last) return; last = sig;
              if (ord) rows.sort((a, b) => { const x = a.data[ord], y = b.data[ord]; const c = x === y ? 0 : (x > y ? 1 : -1); return dir === 'desc' ? -c : c; });
              if (lim) rows = rows.slice(0, lim);
              cb({ docs: rows.map((r) => ({ id: r.id, exists: true, data: () => r.data })) });
            } catch (e) { if (err) err(e); }
          };
          (listeners[name] = listeners[name] || new Set()).add(run);
          run();
          return () => listeners[name].delete(run);
        }
      };
      return q;
    }
  };
  const assets = {
    upload: async (blob, opts) => {
      const type = (opts && opts.type) || blob.type || 'application/octet-stream';
      const r = await jf('/api/asset', { method: 'POST', headers: { 'Content-Type': type }, body: blob });
      return { id: r.id, url: '/_blob/' + r.id, sizeBytes: blob.size, contentType: type };
    },
    delete: (id) => jf('/api/asset/' + id, { method: 'DELETE' })
  };
  const downloads = {
    save: async (o) => {
      const url = URL.createObjectURL(o.data instanceof Blob ? o.data : new Blob([o.data]));
      const a = document.createElement('a');
      a.href = url; a.download = o.filename; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
    }
  };
  window.claude = { use: async (name) => ({ db: db, assets: assets, downloads: downloads })[name] || null };
})();
