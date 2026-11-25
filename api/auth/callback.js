import { googleEnabled, parseCookies, redirectUri, cookie, sessionCookie, emailAllowed, redirect } from '../../lib/auth.js';

export default async function handler(req, res) {
  const clearState = cookie('oauth_state', '', 0, req);
  const fail = (code) => redirect(res, `/?auth_error=${code}`, [clearState]);
  if (!googleEnabled()) return fail('not_configured');

  const query = new URL(req.url, 'http://localhost').searchParams;
  if (query.get('error')) return fail('cancelled');
  const state = query.get('state');
  const code = query.get('code');
  if (!code || !state || state !== parseCookies(req).oauth_state) return fail('state');

  try {
    const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        client_id: process.env.GOOGLE_CLIENT_ID,
        client_secret: process.env.GOOGLE_CLIENT_SECRET,
        redirect_uri: redirectUri(req),
        grant_type: 'authorization_code',
      }),
    });
    const tokens = await tokenRes.json();
    if (!tokenRes.ok || !tokens.id_token) {
      console.error('Google token exchange failed:', tokenRes.status, tokens.error, tokens.error_description);
      return fail('exchange');
    }
    // id_token nhận trực tiếp từ Google qua HTTPS nên chỉ cần kiểm tra các claim.
    const claims = JSON.parse(Buffer.from(tokens.id_token.split('.')[1], 'base64url').toString('utf8'));
    const issuerOk = claims.iss === 'https://accounts.google.com' || claims.iss === 'accounts.google.com';
    if (!issuerOk || claims.aud !== process.env.GOOGLE_CLIENT_ID || claims.exp * 1000 < Date.now()) return fail('token');
    if (!claims.email || claims.email_verified !== true) return fail('unverified');
    if (!emailAllowed(claims.email)) return fail('not_allowed');
    redirect(res, '/', [clearState, sessionCookie({ email: claims.email, name: claims.name || '' }, req)]);
  } catch (error) {
    console.error('Google login error:', error);
    fail('exchange');
  }
}
