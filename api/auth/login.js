import { randomBytes } from 'node:crypto';
import { googleEnabled, redirectUri, cookie, redirect } from '../../lib/auth.js';

export default function handler(req, res) {
  if (!googleEnabled()) return redirect(res, '/?auth_error=not_configured');
  const state = randomBytes(16).toString('base64url');
  const params = new URLSearchParams({
    client_id: process.env.GOOGLE_CLIENT_ID,
    redirect_uri: redirectUri(req),
    response_type: 'code',
    scope: 'openid email profile',
    state,
    prompt: 'select_account',
  });
  const domains = String(process.env.ALLOWED_EMAIL_DOMAINS || '').split(',').map((s) => s.trim()).filter(Boolean);
  if (domains.length === 1 && !process.env.ALLOWED_EMAILS) params.set('hd', domains[0]);
  redirect(res, `https://accounts.google.com/o/oauth2/v2/auth?${params}`, [cookie('oauth_state', state, 600, req)]);
}
