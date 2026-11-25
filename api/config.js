export default function handler(req, res) {
  res.status(200).json({
    passwordRequired: Boolean(process.env.APP_PASSWORD),
    configured: Boolean(process.env.ANTHROPIC_API_KEY),
  });
}
