// Runtime verification of the demo in a real browser.
//
// Drives the installed Chrome through playwright-core (no browser download): loads the demo,
// waits for the engine to paint pages, then exercises the chrome — typing, bold, the menu bar,
// zoom, the outline pane, undo, the locale and colour-mode bindings, and the provider-first
// composition — and writes screenshots plus docs/verification.json.
//
// Usage: node scripts/verify-demo.mjs [url]
import { chromium } from 'playwright-core';
import { mkdir, writeFile } from 'node:fs/promises';

const url = process.argv[2] ?? 'http://127.0.0.1:4312/';
const executablePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

const results = [];
const consoleErrors = [];
const failedRequests = [];
let failures = 0;

function record(name, ok, detail = '') {
  results.push({ name, ok, detail });
  if (!ok) failures++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

const browser = await chromium.launch({
  executablePath,
  headless: true,
  args: ['--no-sandbox', '--disable-gpu'],
});
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
page.on('console', (message) => {
  if (message.type() === 'error') consoleErrors.push(message.text());
});
page.on('pageerror', (error) => consoleErrors.push(`pageerror: ${error.message}`));
page.on('response', (response) => {
  if (response.status() >= 400) failedRequests.push(`${response.status()} ${response.url()}`);
});

try {
  await page.goto(url, { waitUntil: 'load', timeout: 60_000 });

  // 1. The engine paints pages.
  await page.waitForSelector('.docx-pages .docx-page', { timeout: 60_000 });
  await page.waitForFunction(
    () => document.querySelectorAll('.docx-pages .docx-page').length >= 3,
    undefined,
    { timeout: 60_000 }
  );
  const pageCount = await page.locator('.docx-pages .docx-page').count();
  record('engine paints pages', pageCount >= 3, `${pageCount} pages`);

  const paintedText = (await page.locator('.docx-pages').innerText()).slice(0, 4000);
  record('document text is painted', paintedText.trim().length > 200, `${paintedText.trim().length} chars`);

  // 2. The toolbar is derived from the registry.
  const boldButton = page.locator('.docx-toolbar__button[data-slot="text.bold"]');
  record('toolbar renders the bold slot', (await boldButton.count()) === 1);
  record('bold is enabled for an editable document', await boldButton.isEnabled());
  const slots = await page
    .locator('.docx-toolbar [data-slot]')
    .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('data-slot')));
  record('toolbar exposes registry slots', new Set(slots).size >= 15, `${new Set(slots).size} slots`);

  // 3. Typing lands in the document.
  const surface = page.locator('.docx-page-content').first();
  await surface.click({ position: { x: 300, y: 300 } });
  await page.keyboard.press('Control+End');
  await page.keyboard.type(' Angular typing works.');
  await page.waitForTimeout(400);
  const afterTyping = await page.locator('.docx-pages').innerText();
  record('typing reaches the document', afterTyping.includes('Angular typing works.'));

  // 4. Bold is applied through the engine and reflected on the control.
  await page.keyboard.down('Shift');
  for (let i = 0; i < 18; i++) await page.keyboard.press('ArrowLeft');
  await page.keyboard.up('Shift');
  await boldButton.click();
  await page.waitForTimeout(400);
  const boldActive = await boldButton.getAttribute('data-active');
  record('bold becomes active after the click', boldActive !== null);

  // 5. Undo takes the edits back out. Two steps: one reverts the bold, the next the typing —
  // the engine groups them separately, so the assertion walks until the text is gone.
  const markupBeforeUndo = await page.locator('.docx-pages').innerHTML();
  await page.locator('.docx-toolbar__button[data-slot="history.undo"]').click();
  await page.waitForTimeout(400);
  const markupAfterUndo = await page.locator('.docx-pages').innerHTML();
  record('undo repaints the document', markupBeforeUndo !== markupAfterUndo);
  if ((await page.locator('.docx-pages').innerText()).includes('Angular typing works.')) {
    await page.locator('.docx-toolbar__button[data-slot="history.undo"]').click();
    await page.waitForTimeout(400);
  }
  const afterUndo = await page.locator('.docx-pages').innerText();
  record('undo reverts the typed edit', !afterUndo.includes('Angular typing works.'));

  // 5b. The imperative handle serializes the document (the demo's Save calls ref.save()).
  const downloadPromise = page.waitForEvent('download', { timeout: 20_000 });
  await page.locator('.app-actions button', { hasText: 'Save' }).click();
  const download = await downloadPromise;
  record(
    'the ref handle serializes a .docx download',
    download.suggestedFilename() === 'sample.docx',
    download.suggestedFilename()
  );

  // 6. The menu bar is derived from CHROME_MENUS and opens.
  const fileTrigger = page.locator('.docx-menubar__trigger', { hasText: 'File' });
  record('menu bar renders File', (await fileTrigger.count()) === 1);
  await fileTrigger.click();
  await page.waitForTimeout(200);
  const menuRows = await page.locator('.docx-menubar__menu .docx-menubar__item').count();
  record('File menu lists rows', menuRows >= 4, `${menuRows} rows`);
  const saveEnabled = await page
    .locator('.docx-menubar__menu .docx-menubar__item[data-slot="file.save"]')
    .getAttribute('aria-disabled');
  record('File › Save is available', saveEnabled === null);
  await page.keyboard.press('Escape');

  // 7. Zoom is engine state and the readout follows it.
  const zoomValue = page.locator('.docx-toolbar__zoom .docx-toolbar__stepper-value');
  const before = (await zoomValue.innerText()).trim();
  await page.locator('.docx-toolbar__zoom .docx-toolbar__stepper-button').first().click();
  await page.waitForTimeout(300);
  const after = (await zoomValue.innerText()).trim();
  record('zoom stepper changes the scale', before !== after, `${before} -> ${after}`);
  await page.locator('.docx-toolbar__zoom .docx-toolbar__stepper-button').last().click();
  await page.waitForTimeout(300);
  await page.locator('.docx-editor__scroll-container').evaluate((node) => (node.scrollTop = 0));
  await page.waitForTimeout(300);
  await page.screenshot({ path: 'docs/screenshots/editor-packaged.png' });

  // 8. The outline pane lists the document's headings and scrolls to one.
  await page.locator('.docx-nav__toggle').first().click();
  await page.waitForTimeout(300);
  const headingRows = await page.locator('button.docx-nav__heading').count();
  record('outline pane lists headings', headingRows >= 1, `${headingRows} headings`);
  if (headingRows > 3) {
    await page.locator('button.docx-nav__heading').nth(3).click();
    await page.waitForTimeout(500);
    const scrolled = await page
      .locator('.docx-editor__scroll-container')
      .evaluate((node) => node.scrollTop);
    record('outline row scrolls the document', scrolled > 0, `scrollTop=${Math.round(scrolled)}`);
  }
  await page.screenshot({ path: 'docs/screenshots/editor-outline.png' });

  // 9. The locale binding re-labels the chrome.
  await page.locator('.docx-nav__close').first().click();
  await page.waitForTimeout(200);
  await page.locator('.locale-select').selectOption('es');
  await page.waitForTimeout(500);
  const spanishMenu = await page.locator('.docx-menubar__trigger', { hasText: 'Archivo' }).count();
  record('locale switch re-labels the chrome', spanishMenu === 1);
  await page.locator('.locale-select').selectOption('en');
  await page.waitForTimeout(400);

  // 10. Colour mode themes the chrome only.
  await page.locator('.app-actions button', { hasText: 'Theme' }).click();
  await page.waitForTimeout(400);
  const dark = await page.locator('docx-editor .docx-editor.dark').count();
  record('colour mode applies the dark chrome class', dark === 1);
  await page.screenshot({ path: 'docs/screenshots/editor-dark.png' });
  await page.locator('.app-actions button', { hasText: 'Theme' }).click();
  await page.locator('.app-actions button', { hasText: 'Theme' }).click();
  await page.waitForTimeout(300);

  // 11. The provider-first composition renders the same chrome from the same store.
  await page.locator('.toggle input').check();
  await page.waitForSelector('demo-composed-editor .docx-pages .docx-page', { timeout: 60_000 });
  await page.waitForFunction(
    () => document.querySelectorAll('demo-composed-editor .docx-page').length >= 3,
    undefined,
    { timeout: 60_000 }
  );
  const composedPages = await page.locator('demo-composed-editor .docx-page').count();
  record('provider-first composition paints pages', composedPages >= 3, `${composedPages} pages`);
  const composedSlots = await page.locator('demo-composed-editor .docx-toolbar [data-slot]').count();
  record('provider-first composition renders the toolbar', composedSlots >= 15, `${composedSlots} controls`);
  const composedMenu = await page.locator('demo-composed-editor .docx-menubar__trigger').count();
  record('provider-first composition renders the menu bar', composedMenu === 5, `${composedMenu} menus`);
  await page.screenshot({ path: 'docs/screenshots/editor-composed.png' });

  // 12. Nothing 404'd and nothing threw.
  record('no failed requests', failedRequests.length === 0, failedRequests.slice(0, 3).join(' | '));
  record('no console errors', consoleErrors.length === 0, consoleErrors.slice(0, 2).join(' | '));
} catch (error) {
  record('verification script ran to completion', false, error.message);
} finally {
  await mkdir('docs/screenshots', { recursive: true });
  const report = {
    url,
    when: new Date().toISOString(),
    failures,
    results,
    consoleErrors: consoleErrors.slice(0, 25),
    failedRequests: failedRequests.slice(0, 25),
  };
  await writeFile('docs/verification.json', JSON.stringify(report, null, 2));
  await browser.close();
  console.log(`\n${results.length - failures}/${results.length} checks passed`);
  process.exit(failures === 0 ? 0 : 1);
}
