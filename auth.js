const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const { OAuth2Client } = require('google-auth-library');
const loadEnv = require('./env');

loadEnv('.env');

const PORT = 53682;
const file = path.resolve(__dirname, process.env.GOOGLE_CREDENTIALS || 'credentials.json');
const out = path.resolve(__dirname, 'token.json');

if (!fs.existsSync(file)) {
  console.error('No se encontró ' + file + '. Pon ahí el JSON de cliente OAuth descargado de Google Cloud.');
  process.exit(1);
}

let raw;
try { raw = JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) { console.error('El archivo de credenciales no es un JSON válido.'); process.exit(1); }
const c = raw.installed || raw.web;
if (!c) {
  console.error('Ese JSON no es un secreto de cliente OAuth. Si ya es un token o una cuenta de servicio no necesitas este paso.');
  process.exit(1);
}

const redirect = 'http://localhost:' + PORT + '/oauth2callback';
const client = new OAuth2Client(c.client_id, c.client_secret, redirect);
const url = client.generateAuthUrl({ access_type: 'offline', prompt: 'consent', scope: ['https://www.googleapis.com/auth/drive.readonly'] });

const server = http.createServer(async (req, res) => {
  const u = new URL(req.url, redirect);
  if (u.pathname !== '/oauth2callback') { res.writeHead(404); res.end(); return; }
  const code = u.searchParams.get('code');
  if (!code) { res.writeHead(400); res.end('Falta el código de autorización.'); return; }
  try {
    const { tokens } = await client.getToken(code);
    if (!tokens.refresh_token) throw new Error('Google no devolvió refresh_token. Revoca el acceso de la app en tu cuenta de Google y repite.');
    fs.writeFileSync(out, JSON.stringify({ type: 'authorized_user', client_id: c.client_id, client_secret: c.client_secret, refresh_token: tokens.refresh_token }, null, 2));
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end('<p>Listo. Ya puedes cerrar esta pestaña.</p>');
    console.log('Se guardó token.json. Pon GOOGLE_CREDENTIALS=token.json en el archivo .env.');
  } catch (e) {
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('No se pudo completar: ' + e.message);
    console.error(e.message);
  }
  server.close();
});

server.listen(PORT, '127.0.0.1', () => {
  console.log('Abre este enlace y autoriza el acceso de solo lectura a Drive:');
  console.log(url);
  const cmd = process.platform === 'win32' ? 'start "" "' + url + '"' : process.platform === 'darwin' ? 'open "' + url + '"' : 'xdg-open "' + url + '"';
  exec(cmd, () => {});
});
