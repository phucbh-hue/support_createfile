import { clearSessionCookie, redirect } from '../../lib/auth.js';

export default function handler(req, res) {
  redirect(res, '/', [clearSessionCookie(req)]);
}
