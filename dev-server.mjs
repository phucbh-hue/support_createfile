// Máy chủ chạy thử ở máy cá nhân: phục vụ public/ và các hàm trong api/.
// Dùng: ANTHROPIC_API_KEY=... node dev-server.mjs
import http from 'node:http';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml' };
const port = Number(process.env.PORT || 3000);

http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  res.status = (code) => { res.statusCode = code; return res; };
  res.json = (obj) => { res.setHeader('Content-Type', 'application/json; charset=utf-8'); res.end(JSON.stringify(obj)); };
  try {
    if (url.pathname.startsWith('/api/')) {
      const name = url.pathname.slice(5).replace(/[^a-z0-9_-]/gi, '');
      const mod = await import(`./api/${name}.js`);
      if (req.method === 'POST') {
        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        req.body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
      }
      return await mod.default(req, res);
    }
    const rel = url.pathname === '/' ? 'index.html' : url.pathname.slice(1);
    const file = path.join(root, 'public', path.normalize(rel));
    if (!file.startsWith(path.join(root, 'public'))) return res.status(403).end();
    const data = await readFile(file);
    res.setHeader('Content-Type', types[path.extname(file)] || 'application/octet-stream');
    res.end(data);
  } catch (error) {
    if (error.code === 'ENOENT' || error.code === 'ERR_MODULE_NOT_FOUND') return res.status(404).end('Not found');
    console.error(error);
    if (!res.headersSent) res.status(500);
    res.end('Server error');
  }
}).listen(port, () => console.log(`http://localhost:${port}`));
