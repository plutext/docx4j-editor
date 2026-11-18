// SPDX-License-Identifier: GPL-3.0-or-later
//
// Page numbers from the ribbon (ED-005 section 12.95): a document with no footers gets one with a
// centred PAGE field, shown at every page boundary and at the end with the page's number; a second
// press changes nothing; the download carries the footer parts and the references.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages } from './helpers.js';

test('the button makes a footer with the page number, shown at the boundaries, idle the second time; the saved file has the parts', async ({ page }) => {
  await openFixture(page, 'headers.docx');   // five sections, headers, no footers, odd and even pages
  await expect(page.locator('.docx-editor .page-end .boundary-footer .story-view')).toHaveCount(0);
  await page.getByTestId('page-numbers').click();
  const end = page.locator('.docx-editor .page-end .boundary-footer .story-view');
  await expect(end).toHaveCount(1);
  await expect(end.locator('.field')).toHaveText(/^\d+$/);
  const last = await end.locator('.field').textContent();
  const first = page.locator('.docx-editor .page-boundary').first().locator('.boundary-footer .story-view .field');
  await expect(first).toHaveText('1');
  expect(Number(last)).toBeGreaterThan(1);
  await page.getByTestId('page-numbers').click();
  await expect(page.locator('.docx-editor .page-end .boundary-footer .story-view')).toHaveCount(1);
  await expect(end.locator('.field')).toHaveCount(1);
  const saved = await download(page);
  const diffs = comparePackages(await fixture('headers.docx'), saved, ['word/document.xml', 'word/styles.xml', 'docProps/app.xml']);
  expect(diffs.map((d) => `${d.name}:${d.reason}`).sort()).toEqual(['word/footer1.xml:added', 'word/footer2.xml:added']);
});
