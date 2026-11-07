const http = require('http');
const fs = require('fs');
const path = require('path');
const loadEnv = require('./env');

const root = __dirname;
loadEnv('.env');

const cfg = {
  port: Number(process.env.PORT) || 3000,
  host: process.env.HOST || '',
  credentials: process.env.GOOGLE_CREDENTIALS || 'credentials.json',
  credentialsJson: process.env.GOOGLE_CREDENTIALS_JSON || '',
  incluirDocx: /^(1|true|si|yes)$/i.test(process.env.INCLUIR_DOCX || ''),
  dataDir: process.env.DATA_DIR || 'data',
  driveRootFolderId: process.env.DRIVE_ROOT_FOLDER_ID || '',
  supabaseUrl: process.env.SUPABASE_URL || '',
  supabaseKey: process.env.SUPABASE_SERVICE_KEY || '',
  supabaseBucket: process.env.SUPABASE_BUCKET || 'assets',
  appPassword: process.env.APP_PASSWORD || '',
  sessionSecret: process.env.SESSION_SECRET || ''
};

const store = require('./store')(cfg, root);
const drive = process.env.HYA_FAKE_DRIVE ? require(path.resolve(process.env.HYA_FAKE_DRIVE))(cfg, root) : require('./drive')(cfg, root);
const session = require('./session')(cfg.appPassword, cfg.sessionSecret);
const sync = require('./sync');
const exporter = require('./exporter');
const { importDocx, extractFirma } = require('./importer');

if (store.kind === 'supabase' && !session.enabled) {
  console.error('Define APP_PASSWORD en el entorno. Con Supabase la aplicación no arranca sin contraseña de acceso.');
  process.exit(1);
}

const NAME = /^[A-Za-z0-9_-]{1,80}$/;
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const PUBLIC_FILES = new Set(['/login', '/login.html']);
const clients = new Set();

function send(res, code, body, type, extra) {
  const buf = Buffer.isBuffer(body) ? body : Buffer.from(typeof body === 'string' ? body : JSON.stringify(body));
  res.writeHead(code, Object.assign({ 'Content-Type': type || 'application/json; charset=utf-8', 'Content-Length': buf.length, 'Cache-Control': 'no-store' }, extra || {}));
  res.end(buf);
}

function fail(res, code, message) { send(res, code, { error: message }); }

function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) { reject(new Error('Contenido demasiado grande')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function notify(col) {
  const msg = 'data: ' + JSON.stringify({ col: col }) + '\n\n';
  clients.forEach((c) => { try { c.write(msg); } catch (e) {} });
}

function serveStatic(res, pathname) {
  const rel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  const base = path.resolve(root, 'public');
  const file = path.resolve(base, rel);
  if (!file.startsWith(base + path.sep) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) return fail(res, 404, 'No encontrado');
  let body = fs.readFileSync(file);
  if (path.basename(file) === 'index.html') body = Buffer.from(body.toString('utf8').replace('@@AUTH@@', session.enabled ? 'true' : 'false'));
  send(res, 200, body, MIME[path.extname(file)] || 'application/octet-stream');
}

function clientIp(req) { return String(req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim(); }

async function handle(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const p = url.pathname;
  const parts = p.split('/').filter(Boolean);
  const method = req.method;

  if (p === '/healthz') return send(res, 200, { ok: true });

  if (p === '/api/login' && method === 'POST') {
    if (!session.enabled) return send(res, 200, { ok: true });
    const ip = clientIp(req);
    if (session.throttled(ip)) return fail(res, 429, 'Demasiados intentos. Espera un minuto.');
    let pw = '';
    try { pw = JSON.parse((await readBody(req, 4096)).toString('utf8')).password || ''; } catch (e) {}
    if (!session.check(pw)) { session.fail(ip); return fail(res, 401, 'Contraseña incorrecta.'); }
    return send(res, 200, { ok: true }, null, { 'Set-Cookie': session.cookie(req, session.issue()) });
  }

  if (p === '/api/logout' && method === 'POST') return send(res, 200, { ok: true }, null, { 'Set-Cookie': session.cookie(req, '', true) });

  const authed = session.valid(session.readCookie(req));
  if (!authed) {
    if (method === 'GET' && PUBLIC_FILES.has(p)) return serveStatic(res, '/login.html');
    if (parts[0] === 'api' || parts[0] === '_blob') return fail(res, 401, 'Sesión no iniciada');
    res.writeHead(302, { Location: '/login' });
    return res.end();
  }
  if (p === '/login' || p === '/login.html') { res.writeHead(302, { Location: '/' }); return res.end(); }

  if (p === '/api/events') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' });
    res.write('retry: 2000\n\n');
    clients.add(res);
    const beat = setInterval(() => { try { res.write(': ping\n\n'); } catch (e) {} }, 25000);
    req.on('close', () => { clearInterval(beat); clients.delete(res); });
    return;
  }

  if (parts[0] === 'api' && parts[1] === 'col' && parts[2] && method === 'GET') {
    if (!NAME.test(parts[2])) return fail(res, 400, 'Colección no válida');
    return send(res, 200, await store.listCol(parts[2]));
  }

  if (parts[0] === 'api' && parts[1] === 'doc' && parts[2] && parts[3]) {
    if (!NAME.test(parts[2]) || !NAME.test(parts[3])) return fail(res, 400, 'Ruta no válida');
    if (method === 'GET') {
      const d = await store.getDoc(parts[2], parts[3]);
      return d ? send(res, 200, d) : fail(res, 404, 'No existe');
    }
    if (method === 'PUT') {
      const body = await readBody(req, 8 * 1024 * 1024);
      let data;
      try { data = JSON.parse(body.toString('utf8')); } catch (e) { return fail(res, 400, 'JSON no válido'); }
      await store.putDoc(parts[2], parts[3], data);
      notify(parts[2]);
      return send(res, 200, { ok: true });
    }
    if (method === 'DELETE') {
      await store.delDoc(parts[2], parts[3]);
      notify(parts[2]);
      return send(res, 200, { ok: true });
    }
  }

  if (p === '/api/asset' && method === 'POST') {
    const body = await readBody(req, 64 * 1024 * 1024);
    if (!body.length) return fail(res, 400, 'Archivo vacío');
    const id = await store.putAsset(body, String(req.headers['content-type'] || 'application/octet-stream'));
    return send(res, 200, { id: id });
  }

  if (parts[0] === 'api' && parts[1] === 'asset' && parts[2] && method === 'DELETE') {
    if (!/^[a-f0-9]{32}$/.test(parts[2])) return fail(res, 400, 'Identificador no válido');
    await store.delAsset(parts[2]);
    return send(res, 200, { ok: true });
  }

  if (parts[0] === '_blob' && parts[1] && method === 'GET') {
    if (!/^[a-f0-9]{32}$/.test(parts[1])) return fail(res, 400, 'Identificador no válido');
    const a = await store.getAsset(parts[1]);
    if (!a) return fail(res, 404, 'No existe');
    res.writeHead(200, { 'Content-Type': a.type, 'Content-Length': a.buf.length, 'Cache-Control': 'private, max-age=31536000, immutable' });
    return res.end(a.buf);
  }

  if (p === '/api/import' && method === 'POST') {
    const projectId = url.searchParams.get('projectId') || '';
    if (!NAME.test(projectId)) return fail(res, 400, 'Proyecto no válido');
    const buf = await readBody(req, 128 * 1024 * 1024);
    try {
      const r = await importDocx(store, buf, { projectId: projectId, groupId: url.searchParams.get('groupId') || '', name: url.searchParams.get('name') || '' });
      notify('informes2');
      notify('proyectos');
      notify('config');
      return send(res, 200, { id: r.id, warns: r.warns, infos: r.infos });
    } catch (e) {
      const m = String((e && e.message) || '');
      return fail(res, 422, /zip|central directory/i.test(m) ? 'El archivo no es un documento de Word válido.' : (m || 'No se pudo procesar el archivo.'));
    }
  }

  if (p === '/api/firmas/extract' && method === 'POST') {
    const buf = await readBody(req, 128 * 1024 * 1024);
    try { return send(res, 200, { firma: await extractFirma(store, buf) }); }
    catch (e) { return fail(res, 422, (e && e.message) || 'No se pudo leer el documento de Word.'); }
  }

  if (parts[0] === 'api' && parts[1] === 'drive') {
    try {
      if (parts[2] === 'folders') return send(res, 200, await drive.folders(url.searchParams.get('parent') || 'root'));
      if (parts[2] === 'folder') return send(res, 200, await drive.folder(url.searchParams.get('ref')));
      if (parts[2] === 'send' && method === 'POST') {
        const id = url.searchParams.get('id') || '';
        if (!NAME.test(id)) return fail(res, 400, 'Informe no válido');
        try { return send(res, 200, await exporter.exportOne({ store: store, drive: drive, notify: notify }, id)); }
        catch (e) { return fail(res, 422, (e && e.message) || 'No se pudo enviar a Drive'); }
      }
      if (parts[2] === 'export') {
        const projectId = url.searchParams.get('projectId') || '';
        if (!NAME.test(projectId)) return fail(res, 400, 'Proyecto no válido');
        if (method === 'POST') return send(res, 200, exporter.start({ store: store, drive: drive, notify: notify }, projectId, url.searchParams.get('groupId') || ''));
        return send(res, 200, exporter.status(projectId));
      }
      if (parts[2] === 'sync') {
        const projectId = url.searchParams.get('projectId') || '';
        if (!NAME.test(projectId)) return fail(res, 400, 'Proyecto no válido');
        if (method === 'POST') return send(res, 200, sync.start({ store: store, drive: drive, notify: notify }, projectId));
        return send(res, 200, sync.status(projectId));
      }
    } catch (e) {
      return fail(res, 502, e && e.message ? e.message : 'Error al leer Drive');
    }
  }

  if (method === 'GET') return serveStatic(res, p);
  fail(res, 404, 'No encontrado');
}

const server = http.createServer((req, res) => {
  handle(req, res).catch((e) => { try { fail(res, 500, e && e.message ? e.message : 'Error interno'); } catch (x) {} });
});

store.init().then(() => {
  const host = cfg.host || (session.enabled ? '0.0.0.0' : '127.0.0.1');
  server.listen(cfg.port, host, () => {
    console.log('Generador de informes HyA en http://localhost:' + cfg.port + ' (almacenamiento: ' + store.kind + ')');
  });
}).catch((e) => {
  console.error(e.message);
  process.exit(1);
});
