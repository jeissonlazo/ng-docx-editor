// Runs the whole browser verification: serve the production demo build on a FREE port,
// drive Chrome against it, tear the server down.
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';

/** The first port in the range that can actually be bound right now. */
async function freePort(start = 4312, attempts = 20) {
  for (let port = start; port < start + attempts; port++) {
    const free = await new Promise((resolve) => {
      const probe = createServer();
      probe.once('error', () => resolve(false));
      probe.once('listening', () => probe.close(() => resolve(true)));
      probe.listen(port, '127.0.0.1');
    });
    if (free) return port;
  }
  throw new Error('no free port in range');
}

const port = await freePort();

const server = spawn(process.execPath, ['scripts/serve-demo.mjs', 'dist/demo/browser', String(port)], {
  stdio: ['ignore', 'pipe', 'pipe'],
});
server.stderr.on('data', (chunk) => process.stderr.write(chunk));

let ready = false;
server.stdout.on('data', (chunk) => {
  const text = chunk.toString();
  process.stdout.write(text);
  if (text.includes('serving')) ready = true;
});
server.on('error', (error) => {
  console.error('the demo server failed:', error.message);
  process.exit(1);
});

const stop = () => {
  if (!server.killed) server.kill();
};
process.on('exit', stop);

for (let attempt = 0; attempt < 50 && !ready; attempt++) {
  await new Promise((resolve) => setTimeout(resolve, 100));
}
if (!ready) {
  console.error('the demo server did not start');
  stop();
  process.exit(1);
}

const verify = spawn(process.execPath, ['scripts/verify-demo.mjs', `http://127.0.0.1:${port}/`], {
  stdio: 'inherit',
});

verify.on('exit', (code) => {
  stop();
  process.exit(code ?? 1);
});
