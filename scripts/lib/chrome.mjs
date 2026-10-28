// Resolve a Chrome/Chromium binary for the verification scripts.
//
// Order: DOCX_CHROME_PATH (CI sets this), then the usual Windows install locations, then the
// names Playwright and Linux CI images expose. Verification never downloads a browser: it drives
// whatever Chrome the machine already has, so the suite stays a few megabytes of tooling.
import { existsSync } from 'node:fs';
import { chromium } from 'playwright-core';

const WINDOWS_CANDIDATES = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
];

const UNIX_CANDIDATES = [
  '/usr/bin/google-chrome',
  '/usr/bin/google-chrome-stable',
  '/usr/bin/chromium',
  '/usr/bin/chromium-browser',
  '/opt/google/chrome/chrome',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
];

/** The Chrome executable this machine will actually run. Throws when none is found. */
export function resolveChrome() {
  const explicit = process.env['DOCX_CHROME_PATH'];
  if (explicit) {
    if (!existsSync(explicit)) {
      throw new Error(`DOCX_CHROME_PATH points at a missing file: ${explicit}`);
    }
    return explicit;
  }
  const candidates = process.platform === 'win32' ? WINDOWS_CANDIDATES : UNIX_CANDIDATES;
  for (const candidate of candidates) {
    if (existsSync(candidate)) return candidate;
  }
  throw new Error(
    `No Chrome found. Set DOCX_CHROME_PATH to a Chrome/Chromium executable. Looked at:\n  ${candidates.join('\n  ')}`
  );
}

/** Launch headless Chrome with the flags headless containers need. */
export function launchChrome() {
  return chromium.launch({
    executablePath: resolveChrome(),
    headless: true,
    args: ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'],
  });
}
