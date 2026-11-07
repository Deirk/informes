const crypto = require('crypto');

const DAY = 24 * 60 * 60 * 1000;
const TTL = 30 * DAY;

function digest(v) { return crypto.createHash('sha256').update(String(v)).digest(); }

module.exports = function createSession(password, secret) {
  const enabled = !!password;
  const key = secret || crypto.createHash('sha256').update('hya:' + password).digest('hex');
  const attempts = {};

  function sign(value) { return crypto.createHmac('sha256', key).update(value).digest('hex'); }

  function issue() {
    const exp = String(Date.now() + TTL);
    return exp + '.' + sign(exp);
  }

  function valid(token) {
    if (!enabled) return true;
    const parts = String(token || '').split('.');
    if (parts.length !== 2 || !/^\d+$/.test(parts[0])) return false;
    if (Number(parts[0]) < Date.now()) return false;
    const a = Buffer.from(sign(parts[0]));
    const b = Buffer.from(parts[1]);
    return a.length === b.length && crypto.timingSafeEqual(a, b);
  }

  function check(candidate) {
    return crypto.timingSafeEqual(digest(candidate), digest(password));
  }

  function throttled(ip) {
    const now = Date.now();
    const list = (attempts[ip] || []).filter((t) => now - t < 60000);
    attempts[ip] = list;
    return list.length >= 10;
  }

  function fail(ip) { (attempts[ip] = attempts[ip] || []).push(Date.now()); }

  function readCookie(req) {
    const raw = req.headers.cookie || '';
    const m = /(?:^|;\s*)hya=([^;]+)/.exec(raw);
    return m ? decodeURIComponent(m[1]) : '';
  }

  function cookie(req, token, clear) {
    const secure = String(req.headers['x-forwarded-proto'] || '').split(',')[0] === 'https' ? '; Secure' : '';
    return 'hya=' + (clear ? '' : encodeURIComponent(token)) + '; Path=/; HttpOnly; SameSite=Lax; Max-Age=' + (clear ? 0 : Math.floor(TTL / 1000)) + secure;
  }

  return { enabled: enabled, issue: issue, valid: valid, check: check, throttled: throttled, fail: fail, readCookie: readCookie, cookie: cookie };
};
