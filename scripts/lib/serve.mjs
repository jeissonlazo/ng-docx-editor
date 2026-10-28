// Static serving helpers for the verification scripts: pick a free port, serve a build
// directory, and tear it down.
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';

/** The first port in the range that can actually be bound right now. */
export async function freePort(start = 4312, attempts = 20) {
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

/** Serve `directory` on a free port. Resolves once the listener is up. */
export async function serve(directory) {
  const port = await freePort();
  const child = spawn(process.execPath, ['scripts/serve-demo.mjs', directory, String(port)], {
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  child.stderr.on('data', (chunk) => process.stderr.write(chunk));

  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`server for ${directory} did not start`)), 15_000);
    child.stdout.on('data', (chunk) => {
      if (!chunk.toString().includes('serving')) return;
      clearTimeout(timer);
      resolve(undefined);
    });
    child.once('error', reject);
  });

  return {
    port,
    url: `http://127.0.0.1:${port}/`,
    stop: () => {
      if (!child.killed) child.kill();
    },
  };
}
