const path = require('path');
const fs = require('fs');
const { GoogleAuth, OAuth2Client } = require('google-auth-library');

const SCOPES = ['https://www.googleapis.com/auth/drive'];

const FOLDER = 'application/vnd.google-apps.folder';
const GDOC = 'application/vnd.google-apps.document';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const API = 'https://www.googleapis.com/drive/v3/files';
const ID = /^[A-Za-z0-9_-]{10,}$/;

module.exports = function createDrive(cfg, root, deps) {
  const keyFile = path.resolve(root, cfg.credentials || 'credentials.json');
  let auth = null;

  function readCredentials() {
    if (cfg.credentialsJson) {
      try { return JSON.parse(cfg.credentialsJson); } catch (e) { throw new Error('GOOGLE_CREDENTIALS_JSON no es un JSON válido.'); }
    }
    if (!fs.existsSync(keyFile)) {
      throw new Error('No se encontraron credenciales de Google. Define GOOGLE_CREDENTIALS_JSON o GOOGLE_CREDENTIALS en el archivo .env, con las que usa el bot de Telegram.');
    }
    try { return JSON.parse(fs.readFileSync(keyFile, 'utf8')); } catch (e) { throw new Error('El archivo de credenciales no es un JSON válido.'); }
  }

  function getClient() {
    if (auth) return auth;
    const c = readCredentials();
    if (c.client_email && c.private_key) {
      auth = new GoogleAuth({ credentials: c, scopes: SCOPES });
    } else if (c.refresh_token && c.client_id && c.client_secret) {
      const oc = new OAuth2Client(c.client_id, c.client_secret);
      oc.setCredentials({ refresh_token: c.refresh_token });
      auth = oc;
    } else if (c.installed || c.web) {
      throw new Error('Ese JSON es el secreto de cliente OAuth, no un token. Ejecuta npm run auth para generar token.json y apunta GOOGLE_CREDENTIALS a ese archivo.');
    } else {
      throw new Error('Formato de credenciales no reconocido. Se espera una cuenta de servicio (client_email y private_key) o un token OAuth (refresh_token, client_id y client_secret).');
    }
    return auth;
  }

  async function token() {
    const client = getClient();
    const c = client.getClient ? await client.getClient() : client;
    const t = await c.getAccessToken();
    return typeof t === 'string' ? t : t.token;
  }

  async function request(url) {
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + await token() } });
    return res;
  }

  async function realJson(url) {
    const res = await request(url);
    if (!res.ok) {
      let msg = 'Drive respondió ' + res.status;
      try { const j = await res.json(); if (j.error && j.error.message) msg += ': ' + j.error.message; } catch (e) {}
      const err = new Error(msg);
      err.status = res.status;
      throw err;
    }
    return res.json();
  }

  const json = (deps && deps.json) || realJson;
  const insideCache = {};

  function rootId() {
    const id = String(cfg.driveRootFolderId || '').trim();
    if (!id) throw new Error('Falta DRIVE_ROOT_FOLDER_ID en el archivo .env. Es la carpeta de Drive que contiene los proyectos.');
    if (!ID.test(id)) throw new Error('DRIVE_ROOT_FOLDER_ID no es un identificador de carpeta válido.');
    return id;
  }

  async function inside(id) {
    const top = rootId();
    if (id === top) return true;
    if (insideCache[id] !== undefined) return insideCache[id];
    let cur = id;
    const trail = [];
    let result = false;
    for (let depth = 0; depth < 20; depth++) {
      if (cur === top) { result = true; break; }
      if (insideCache[cur] !== undefined) { result = insideCache[cur]; break; }
      trail.push(cur);
      const meta = await json(API + '/' + cur + '?fields=id,parents&supportsAllDrives=true');
      const parents = meta.parents || [];
      if (!parents.length) break;
      if (parents.indexOf(top) > -1) { result = true; break; }
      cur = parents[0];
    }
    trail.forEach((t) => { insideCache[t] = result; });
    insideCache[id] = result;
    return result;
  }

  async function guard(id) {
    safeId(id);
    if (!(await inside(id))) throw new Error('Esa carpeta o archivo está fuera de la carpeta raíz permitida.');
    return id;
  }

  function q(parts) { return encodeURIComponent(parts.join(' and ')); }

  async function listAll(query, fields) {
    const out = [];
    let pageToken = '';
    do {
      const url = API + '?q=' + query + '&fields=' + encodeURIComponent('nextPageToken,files(' + fields + ')') +
        '&pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&orderBy=name' + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      const j = await json(url);
      (j.files || []).forEach((f) => out.push(f));
      pageToken = j.nextPageToken || '';
    } while (pageToken);
    return out;
  }

  function safeId(id) {
    if (!ID.test(String(id || ''))) throw new Error('Identificador de Drive no válido.');
    return id;
  }

  async function folders(parent) {
    const top = rootId();
    const base = !parent || parent === 'root' ? top : parent;
    await guard(base);
    const list = await listAll(q(["'" + base + "' in parents", "mimeType='" + FOLDER + "'", 'trashed=false']), 'id,name');
    let rootInfo = null;
    if (base === top) {
      const meta = await json(API + '/' + top + '?fields=id,name&supportsAllDrives=true');
      rootInfo = { id: meta.id, name: meta.name };
    }
    return { folders: list, root: rootInfo };
  }

  async function folder(ref) {
    let id = String(ref || '').trim();
    const m = /\/folders\/([A-Za-z0-9_-]+)/.exec(id) || /[?&]id=([A-Za-z0-9_-]+)/.exec(id);
    if (m) id = m[1];
    await guard(id);
    const f = await json(API + '/' + id + '?fields=id,name,mimeType&supportsAllDrives=true');
    if (f.mimeType !== FOLDER) throw new Error('Ese enlace no es una carpeta de Drive.');
    return { id: f.id, name: f.name };
  }

  function natural(a, b) { return String(a).localeCompare(String(b), 'es', { numeric: true, sensitivity: 'base' }); }

  async function scan(folderId) {
    await guard(folderId);
    const root = await folder(folderId);
    const wanted = [GDOC].concat(cfg.incluirDocx ? [DOCX] : []);
    const isInf = (f) => /^inf/i.test(String(f.name || '').trim()) && wanted.indexOf(f.mimeType) > -1;
    const items = [];

    async function walk(id, tower, names) {
      const children = await listAll(q(["'" + id + "' in parents", 'trashed=false']), 'id,name,mimeType,modifiedTime');
      for (const f of children) {
        if (f.mimeType === FOLDER) {
          await walk(f.id, tower || { id: f.id, name: f.name }, names.concat([f.name]));
        } else if (isInf(f)) {
          const t = tower || { id: root.id, name: root.name };
          items.push({ id: f.id, name: f.name, modifiedTime: f.modifiedTime, towerId: t.id, towerName: t.name, parentId: id, path: names.join(' / ') });
        }
      }
    }
    await walk(root.id, null, [root.name]);
    items.sort((a, b) => natural(a.towerName, b.towerName) || natural(a.name, b.name));
    return { root: root, items: items };
  }

  function isZip(buf) { return buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; }

  async function doc(id) {
    await guard(id);
    const meta = await json(API + '/' + id + '?fields=id,name,mimeType&supportsAllDrives=true');
    if (meta.mimeType === DOCX) {
      const res = await request(API + '/' + id + '?alt=media&supportsAllDrives=true');
      if (!res.ok) throw new Error('No se pudo descargar el archivo (' + res.status + ').');
      return Buffer.from(await res.arrayBuffer());
    }
    const res = await request(API + '/' + id + '/export?mimeType=' + encodeURIComponent(DOCX));
    if (res.ok) {
      const buf = Buffer.from(await res.arrayBuffer());
      if (isZip(buf)) return buf;
    }
    const alt = await request('https://docs.google.com/document/d/' + id + '/export?format=docx');
    if (alt.ok) {
      const buf = Buffer.from(await alt.arrayBuffer());
      if (isZip(buf)) return buf;
    }
    throw new Error('Drive no permite exportar este documento a Word (probablemente supera los 10 MB). Descárgalo como .docx y súbelo al grupo a mano.');
  }

  async function upload(o) {
    await guard(o.parentId);
    if (o.existingId) await guard(o.existingId);
    const boundary = 'hya' + Math.random().toString(16).slice(2);
    const meta = o.existingId ? { name: o.name } : { name: o.name, parents: [o.parentId] };
    const head = Buffer.from('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) + '\r\n--' + boundary + '\r\nContent-Type: ' + DOCX + '\r\n\r\n');
    const tail = Buffer.from('\r\n--' + boundary + '--');
    const body = Buffer.concat([head, o.buf, tail]);
    const base = 'https://www.googleapis.com/upload/drive/v3/files';
    const url = (o.existingId ? base + '/' + o.existingId : base) + '?uploadType=multipart&supportsAllDrives=true&fields=id,name';
    const res = await fetch(url, { method: o.existingId ? 'PATCH' : 'POST', headers: { Authorization: 'Bearer ' + await token(), 'Content-Type': 'multipart/related; boundary=' + boundary }, body: body });
    if (!res.ok) {
      let m = '';
      try { const j = await res.json(); m = (j.error && j.error.message) || ''; } catch (e) {}
      const err = new Error(res.status === 403 && /scope|insufficient/i.test(m)
        ? 'La credencial de Google es de solo lectura. Genera de nuevo token.json con npm run auth (o usa una cuenta de servicio con permiso de editor en la carpeta).'
        : res.status === 403 ? 'Sin permiso para escribir en esa carpeta de Drive. Comparte la carpeta con permiso de editor.'
        : 'Drive respondió ' + res.status + (m ? ': ' + m : ''));
      err.status = res.status;
      throw err;
    }
    return res.json();
  }

  return { folders: folders, folder: folder, scan: scan, doc: doc, upload: upload, getClient: getClient };
};
