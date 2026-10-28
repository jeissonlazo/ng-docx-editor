// A tiny static server for the production build, used by scripts/verify-demo.mjs.
// No dependencies, no caching surprises: it serves dist/demo/browser over 127.0.0.1.
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';

const root = process.argv[2] ?? 'dist/demo/browser';
const port = Number(process.argv[3] ?? 4312);

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.ico': 'image/x-icon',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.wasm': 'application/wasm',
  '.ttf': 'font/ttf',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

createServer(async (request, response) => {
  const started = Date.now();
  const url = new URL(request.url ?? '/', `http://${request.headers.host}`);
  let path = normalize(join(root, decodeURIComponent(url.pathname)));
  if (!path.startsWith(normalize(root))) {
    response.writeHead(403).end('forbidden');
    return;
  }
  try {
    const info = await stat(path);
    if (info.isDirectory()) path = join(path, 'index.html');
    await stat(path);
  } catch {
    console.log(`${request.method} ${url.pathname} -> 404`);
    response.writeHead(404).end('not found');
    return;
  }
  response.writeHead(200, {
    'content-type': TYPES[extname(path)] ?? 'application/octet-stream',
    'cache-control': 'no-store',
  });
  const stream = createReadStream(path);
  stream.on('close', () => console.log(`${request.method} ${url.pathname} -> 200 (${Date.now() - started}ms)`));
  stream.pipe(response);
}).listen(port, '127.0.0.1', () => {
  console.log(`serving ${root} on http://127.0.0.1:${port}/`);
});
