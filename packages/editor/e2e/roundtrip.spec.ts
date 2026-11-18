// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.5 test 1: the engine in the browser. Open a fixture, save without editing, and
// the download is byte-identical part by part (the E1.a oracle, through the browser build).
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages } from './helpers.js';

for (const name of ['loadAndSave.docx', 'HelloWordOnline.docx', 'invoice.docx', 'sections-multiple.docx']) {
  test(`untouched round trip in the browser: ${name}`, async ({ page }) => {
    await openFixture(page, name);
    await expect(page.getByTestId('name')).toHaveText(name);
    const saved = await download(page);
    await expect(page.getByTestId('status')).toContainText('Saved, unchanged');
    expect(comparePackages(await fixture(name), saved)).toEqual([]);
  });
}

test('typing changes the main part only, and app.xml', async ({ page }) => {
  const name = 'HelloWordOnline.docx';
  await openFixture(page, name);
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await editor.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' (edited in the browser)');
  await expect(editor).toContainText('Hello World from Word Online (edited in the browser)');
  const saved = await download(page);
  await expect(page.getByTestId('status')).toContainText('Saved with changes');
  const diffs = comparePackages(await fixture(name), saved, ['word/document22.xml', 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
  expect(diffs).toEqual([]);
});
