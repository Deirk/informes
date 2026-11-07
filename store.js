const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function newAssetId() { return crypto.randomBytes(16).toString('hex'); }

function createLocal(cfg, root) {
  const dataDir = path.resolve(root, cfg.dataDir);
  const docsDir = path.join(dataDir, 'docs');
  const assetsDir = path.join(dataDir, 'assets');
  const colDir = (col) => path.join(docsDir, col);

  return {
    kind: 'local',
    async init() {
      fs.mkdirSync(docsDir, { recursive: true });
      fs.mkdirSync(assetsDir, { recursive: true });
    },
    async listCol(col) {
      const dir = colDir(col);
      if (!fs.existsSync(dir)) return [];
      const out = [];
      for (const f of fs.readdirSync(dir)) {
        if (!f.endsWith('.json')) continue;
        try { out.push({ id: f.slice(0, -5), data: JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')) }); } catch (e) {}
      }
      return out;
    },
    async getDoc(col, id) {
      const file = path.join(colDir(col), id + '.json');
      if (!fs.existsSync(file)) return null;
      return JSON.parse(fs.readFileSync(file, 'utf8'));
    },
    async putDoc(col, id, data) {
      fs.mkdirSync(colDir(col), { recursive: true });
      const file = path.join(colDir(col), id + '.json');
      const tmp = file + '.tmp';
      fs.writeFileSync(tmp, JSON.stringify(data));
      fs.renameSync(tmp, file);
    },
    async delDoc(col, id) {
      const file = path.join(colDir(col), id + '.json');
      if (fs.existsSync(file)) fs.unlinkSync(file);
    },
    async putAsset(buf, type) {
      const id = newAssetId();
      fs.writeFileSync(path.join(assetsDir, id), buf);
      fs.writeFileSync(path.join(assetsDir, id + '.type'), String(type || 'application/octet-stream'));
      return id;
    },
    async getAsset(id) {
      const fp = path.join(assetsDir, id);
      if (!fs.existsSync(fp)) return null;
      let type = 'application/octet-stream';
      try { type = fs.readFileSync(fp + '.type', 'utf8'); } catch (e) {}
      return { buf: fs.readFileSync(fp), type: type };
    },
    async delAsset(id) {
      [id, id + '.type'].forEach((f) => { const fp = path.join(assetsDir, f); if (fs.existsSync(fp)) fs.unlinkSync(fp); });
    }
  };
}

function createSupabase(cfg) {
  const base = cfg.supabaseUrl.replace(/\/+$/, '');
  const key = cfg.supabaseKey;
  const bucket = cfg.supabaseBucket;
  const jwt = key.startsWith('eyJ');

  async function call(pathname, opts) {
    const o = opts || {};
    const headers = Object.assign({ apikey: key }, jwt ? { Authorization: 'Bearer ' + key } : {}, o.headers || {});
    const res = await fetch(base + pathname, { method: o.method || 'GET', headers: headers, body: o.body });
    if (!res.ok && !(o.allow && o.allow.indexOf(res.status) > -1)) {
      let detail = '';
      try { detail = (await res.text()).slice(0, 300); } catch (e) {}
      throw new Error('Supabase ' + res.status + (detail ? ': ' + detail : ''));
    }
    return res;
  }

  const enc = encodeURIComponent;

  return {
    kind: 'supabase',
    async init() {
      await call('/storage/v1/bucket', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: bucket, name: bucket, public: false }),
        allow: [400, 409]
      });
      const res = await call('/rest/v1/docs?select=id&limit=1', { allow: [404] });
      if (res.status === 404) throw new Error('No existe la tabla docs en Supabase. Ejecuta supabase.sql en el editor SQL del proyecto.');
    },
    async listCol(col) {
      const out = [];
      const page = 1000;
      for (let offset = 0; ; offset += page) {
        const res = await call('/rest/v1/docs?col=eq.' + enc(col) + '&select=id,data&order=id&limit=' + page + '&offset=' + offset);
        const rows = await res.json();
        rows.forEach((r) => out.push({ id: r.id, data: r.data }));
        if (rows.length < page) break;
      }
      return out;
    },
    async getDoc(col, id) {
      const res = await call('/rest/v1/docs?col=eq.' + enc(col) + '&id=eq.' + enc(id) + '&select=data&limit=1');
      const rows = await res.json();
      return rows.length ? rows[0].data : null;
    },
    async putDoc(col, id, data) {
      await call('/rest/v1/docs?on_conflict=col,id', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Prefer: 'resolution=merge-duplicates,return=minimal' },
        body: JSON.stringify({ col: col, id: id, data: data, updated_at: new Date().toISOString() })
      });
    },
    async delDoc(col, id) {
      await call('/rest/v1/docs?col=eq.' + enc(col) + '&id=eq.' + enc(id), { method: 'DELETE', headers: { Prefer: 'return=minimal' } });
    },
    async putAsset(buf, type) {
      const id = newAssetId();
      await call('/storage/v1/object/' + enc(bucket) + '/' + id, {
        method: 'POST',
        headers: { 'Content-Type': type || 'application/octet-stream', 'x-upsert': 'true' },
        body: buf
      });
      return id;
    },
    async getAsset(id) {
      const res = await call('/storage/v1/object/' + enc(bucket) + '/' + enc(id), { allow: [400, 404] });
      if (!res.ok) return null;
      return { buf: Buffer.from(await res.arrayBuffer()), type: res.headers.get('content-type') || 'application/octet-stream' };
    },
    async delAsset(id) {
      await call('/storage/v1/object/' + enc(bucket) + '/' + enc(id), { method: 'DELETE', allow: [400, 404] });
    }
  };
}

module.exports = function createStore(cfg, root) {
  return cfg.supabaseUrl && cfg.supabaseKey ? createSupabase(cfg) : createLocal(cfg, root);
};
