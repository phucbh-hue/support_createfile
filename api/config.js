import { authMode, readSession } from '../lib/auth.js';

export default function handler(req, res) {
  const auth = authMode();
  const user = auth === 'google' ? readSession(req) : null;
  res.setHeader('Cache-Control', 'no-store');
  res.status(200).json({
    auth,
    user: user ? { email: user.email, name: user.name } : null,
    configured: Boolean(process.env.ANTHROPIC_API_KEY),
  });
}
