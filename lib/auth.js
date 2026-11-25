// Đăng nhập: Google OAuth (nếu đã cấu hình) hoặc mã truy cập dùng chung.
import { createHmac, timingSafeEqual } from 'node:crypto';

const SESSION_COOKIE = 'session';
const SESSION_DAYS = 7;

export const googleEnabled = () => Boolean(process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET && process.env.SESSION_SECRET);

export function authMode() {
  if (googleEnabled()) return 'google';
  return process.env.APP_PASSWORD ? 'password' : 'none';
}

export function parseCookies(req) {
  const out = {};
  for (const part of String(req.headers.cookie || '').split(';')) {
    const at = part.indexOf('=');
    if (at > 0) out[part.slice(0, at).trim()] = decodeURIComponent(part.slice(at + 1).trim());
  }
  return out;
}

// Địa chỉ gốc của trang, dùng để tạo redirect URI gửi cho Google.
export function baseUrl(req) {
  if (process.env.APP_URL) return process.env.APP_URL.replace(/\/+$/, '');
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  const proto = req.headers['x-forwarded-proto'] || (String(host).startsWith('localhost') ? 'http' : 'https');
  return `${String(proto).split(',')[0]}://${host}`;
}

export const redirectUri = (req) => `${baseUrl(req)}/api/auth/callback`;

export function cookie(name, value, maxAgeSeconds, req) {
  const secure = baseUrl(req).startsWith('https') ? '; Secure' : '';
  return `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAgeSeconds}${secure}`;
}

const mac = (data) => createHmac('sha256', process.env.SESSION_SECRET).update(data).digest('base64url');

export function createSession(user) {
  const body = Buffer.from(JSON.stringify({ ...user, exp: Date.now() + SESSION_DAYS * 86_400_000 })).toString('base64url');
  return `${body}.${mac(body)}`;
}

export function readSession(req) {
  if (!process.env.SESSION_SECRET) return null;
  const [body, sig] = String(parseCookies(req)[SESSION_COOKIE] || '').split('.');
  if (!body || !sig) return null;
  const a = Buffer.from(sig);
  const b = Buffer.from(mac(body));
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null;
  try {
    const data = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    return data.exp > Date.now() && emailAllowed(data.email) ? data : null;
  } catch {
    return null;
  }
}

export const sessionCookie = (user, req) => cookie(SESSION_COOKIE, createSession(user), SESSION_DAYS * 86_400, req);
export const clearSessionCookie = (req) => cookie(SESSION_COOKIE, '', 0, req);

const list = (value) => String(value || '').toLowerCase().split(',').map((s) => s.trim()).filter(Boolean);

// ALLOWED_EMAIL_DOMAINS / ALLOWED_EMAILS: danh sách cách nhau bằng dấu phẩy.
// Không đặt biến nào thì mọi tài khoản Google đã xác minh đều vào được.
export function emailAllowed(email) {
  const value = String(email || '').toLowerCase();
  const domains = list(process.env.ALLOWED_EMAIL_DOMAINS);
  const emails = list(process.env.ALLOWED_EMAILS);
  if (!domains.length && !emails.length) return Boolean(value);
  return emails.includes(value) || domains.includes(value.split('@')[1]);
}

function passwordOk(given) {
  const expected = process.env.APP_PASSWORD;
  if (!expected) return true;
  const a = Buffer.from(String(given || ''));
  const b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

// Kiểm tra quyền gọi API. Trả về { ok, user } hoặc { ok: false, error }.
export function authorize(req) {
  if (googleEnabled()) {
    const user = readSession(req);
    return user ? { ok: true, user } : { ok: false, error: 'Phiên đăng nhập đã hết hạn. Hãy tải lại trang và đăng nhập bằng Google.' };
  }
  return passwordOk(req.headers['x-app-password']) ? { ok: true, user: null } : { ok: false, error: 'Mã truy cập không đúng.' };
}

export function redirect(res, location, cookies = []) {
  res.statusCode = 302;
  if (cookies.length) res.setHeader('Set-Cookie', cookies);
  res.setHeader('Location', location);
  res.setHeader('Cache-Control', 'no-store');
  res.end();
}
