// Distribution check: pack the built library, install it as a CONSUMER would, build a clean
// Angular app against it, and render that app in a browser.
//
// This is the guard the source-path demo cannot provide: it proves the published artifact —
// `exports` map, FESM bundle, declarations, `styles.css`, peer dependencies — is complete and
// consumable. Run `npm run build` first (the `check:consumer` script does).
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { launchChrome } from './lib/chrome.mjs';
import { serve } from './lib/serve.mjs';

const packageDir = 'dist/ng-docx-editor';
const cacheDir = '.cache';
const consumerProject = 'consumer-smoke';

const results = [];
let failures = 0;

function record(name, ok, detail = '') {
  results.push({ name, ok, detail });
  if (!ok) failures++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

/**
 * Run a command line through the platform shell.
 *
 * npm and npx are `.cmd` shims on Windows, which Node refuses to spawn without a shell — so the
 * whole command is passed as ONE string (a shell plus an argument array is the pattern Node
 * deprecates, because the array is concatenated, not escaped).
 */
function runShell(command, options = {}) {
  return execFileSync(command, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    shell: true,
    ...options,
  });
}

function quote(value) {
  return `"${value}"`;
}

if (!existsSync(join(packageDir, 'package.json'))) {
  console.error(`missing ${packageDir}; run "npm run build" first`);
  process.exit(1);
}

await mkdir(cacheDir, { recursive: true });

// ── 1. Pack ───────────────────────────────────────────────────────────────────────────
// An ABSOLUTE path: a bare relative spec like `dist/ng-docx-editor` is read by npm as the
// GitHub shorthand `owner/repo`, and it tries to clone it.
const packJson = runShell(
  `npm pack ${quote(resolve(packageDir))} --pack-destination ${quote(resolve(cacheDir))} --json`
);
const [packed] = JSON.parse(packJson);
const tarball = resolve(cacheDir, packed.filename);
const files = new Set(packed.files.map((file) => file.path));
const unpackedSize = packed.unpackedSize;
record('npm pack produces a tarball', existsSync(tarball), packed.filename);

// ── 2. The tarball carries everything the manifest promises ───────────────────────────
const required = [
  'package.json',
  'README.md',
  'styles.css',
  'fesm2022/ng-docx-editor.mjs',
  'types/ng-docx-editor.d.ts',
];
for (const requiredFile of required) {
  record(`tarball contains ${requiredFile}`, files.has(requiredFile));
}

const manifest = JSON.parse(await readFile(join(packageDir, 'package.json'), 'utf8'));
const exportTargets = Object.values(manifest.exports ?? {}).flatMap((entry) =>
  typeof entry === 'string'
    ? [entry]
    : Object.values(entry ?? {}).filter((value) => typeof value === 'string')
);
for (const target of exportTargets) {
  const relative = target.replace(/^\.\//, '');
  record(`exports target ${target} is packed`, files.has(relative));
}
record(
  'the styles entry re-exports the engine stylesheet',
  (await readFile(join(packageDir, 'styles.css'), 'utf8')).includes(
    '@docx-editor.dev/core/styles/editor.css'
  )
);
record(
  'peer dependencies name the engine and Angular',
  Object.keys(manifest.peerDependencies ?? {}).includes('@docx-editor.dev/core') &&
    Object.keys(manifest.peerDependencies ?? {}).includes('@angular/core')
);

// ── 3. Install it the way a consumer does ────────────────────────────────────────────
try {
  runShell(
    `npm install --no-save --no-package-lock --no-audit --no-fund ${quote(tarball)}`
  );
  record('npm install <tarball> succeeds', true);
} catch (error) {
  record('npm install <tarball> succeeds', false, String(error.stderr ?? error.message).slice(0, 400));
}

const resolveScript = join(cacheDir, 'resolve-check.mjs');
await writeFile(
  resolveScript,
  "process.stdout.write(import.meta.resolve('ng-docx-editor'));\n"
);
let resolved = '';
try {
  resolved = execFileSync(process.execPath, [resolveScript], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
} catch (error) {
  resolved = String(error.stdout ?? error.message);
}
record(
  'node resolves the package through its exports map',
  resolved.includes('ng-docx-editor'),
  resolved.replace('file:///', '').slice(-60)
);

// ── 4. A clean Angular app compiles against the packed artifact ───────────────────────
try {
  runShell(`npx ng build ${consumerProject}`);
  record(`ng build ${consumerProject} succeeds`, true);
} catch (error) {
  record(
    `ng build ${consumerProject} succeeds`,
    false,
    String(error.stdout ?? '') + String(error.stderr ?? '')
  );
}
const consumerOut = `dist/${consumerProject}/browser`;
record('the consumer build emits a bundle', existsSync(join(consumerOut, 'index.html')));

// ── 5. And it actually renders the editor ────────────────────────────────────────────
let server;
let browser;
try {
  server = await serve(consumerOut);
  browser = await launchChrome();
  const page = await browser.newPage({ viewport: { width: 1280, height: 820 } });
  const consoleErrors = [];
  const failedRequests = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', (error) => consoleErrors.push(String(error)));
  page.on('response', (response) => {
    if (response.status() >= 400) failedRequests.push(`${response.status()} ${response.url()}`);
  });

  await page.goto(server.url, { waitUntil: 'load', timeout: 60_000 });
  await page.waitForSelector('.docx-pages .docx-page', { timeout: 60_000 });
  await page.waitForFunction(
    () => document.querySelectorAll('.docx-pages .docx-page').length >= 3,
    undefined,
    { timeout: 60_000 }
  );
  const pages = await page.locator('.docx-pages .docx-page').count();
  record('the consumer app paints pages', pages >= 3, `${pages} pages`);

  const toolbarSlots = await page.locator('.docx-toolbar [data-slot]').count();
  record('the consumer app renders the chrome', toolbarSlots >= 15, `${toolbarSlots} controls`);

  // The packaged stylesheet must actually reach the page: the engine's bar is a tinted pill,
  // not the transparent default a button would have.
  const barBackground = await page
    .locator('.docx-toolbar')
    .evaluate((node) => getComputedStyle(node).backgroundColor);
  record(
    'the packaged stylesheet is applied',
    barBackground !== 'rgba(0, 0, 0, 0)' && barBackground !== 'transparent',
    barBackground
  );

  record(
    'the consumer app requests only present assets',
    failedRequests.length === 0,
    failedRequests[0] ?? ''
  );
  record('the consumer app logs no errors', consoleErrors.length === 0, consoleErrors[0] ?? '');
  await page.screenshot({ path: 'docs/screenshots/consumer-smoke.png' });
} catch (error) {
  record('the consumer app renders headlessly', false, error.message);
} finally {
  await browser?.close();
  server?.stop();
}

await writeFile(
  join(cacheDir, 'consumer-report.json'),
  JSON.stringify(
    { when: new Date().toISOString(), tarball: packed.filename, unpackedSize, failures, results },
    null,
    2
  )
);
console.log(`\n${results.length - failures}/${results.length} checks passed`);
process.exit(failures === 0 ? 0 : 1);
