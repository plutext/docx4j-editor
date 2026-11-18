// SPDX-License-Identifier: GPL-3.0-or-later
//
// Proofing (view/proofing.ts; Jason, 2026-11-11): the browser's spell checker told what it needs. The document's language
// on the surface; a run's own language (direct, by character style, and East Asian text by w:lang's eastAsia) and Word's
// "Do not check spelling" (direct, and by paragraph style) on its text; the option turning spell checking off, kept per
// browser. proofing/proofing.docx is built by fixtures/build/proofing.mjs.
import { test, expect } from '@playwright/test';
import { openFixture, download, comparePackages, fixture } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

test('the document\'s language on the surface, a run\'s own language and Word\'s noProof on its text; the file saves as it came', async ({ page }) => {
  await openFixture(page, 'proofing/proofing.docx');
  const surface = page.locator(body);
  await expect(surface).toHaveAttribute('lang', 'en-AU');
  await expect(surface).toHaveAttribute('spellcheck', 'true');
  await expect(surface.locator('[lang="fr-FR"]')).toHaveText('Une phrase en français.');
  await expect(surface.locator('[lang="de-DE"]')).toHaveText('Ein deutscher Satz.');
  await expect(surface.locator('[lang="ja-JP"]')).toHaveText('日本語の文章です。');
  // Word's "Do not check spelling or grammar": a run's own, and a paragraph style's (Code) for all its text.
  await expect(surface.locator('[spellcheck="false"]')).toHaveText(['marked Do not check spelling', 'const gizmo = frobnicate(800);']);
  // Text in the document's own language carries nothing of its own: the surface's says it.
  await expect(surface.locator('[lang="en-AU"]')).toHaveCount(0);
  const saved = await download(page);
  expect(comparePackages(await fixture('proofing/proofing.docx'), saved)).toEqual([]);
});

test('the option turns spell checking off on every editing surface at once, and is kept per browser', async ({ page }) => {
  await openFixture(page, 'proofing/proofing.docx');
  const surface = page.locator(body);
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-spellcheck')).toBeChecked();
  await page.getByTestId('option-spellcheck').uncheck();
  await expect(surface).toHaveAttribute('spellcheck', 'false');
  // Kept: after a reload the surface opens with it off.
  await page.reload();
  await page.getByTestId('open').setInputFiles((await import('./helpers.js')).fixturePath('proofing/proofing.docx'));
  await expect(page.locator(body)).toHaveAttribute('spellcheck', 'false');
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-spellcheck')).not.toBeChecked();
  await page.getByTestId('option-spellcheck').check();
  await expect(page.locator(body)).toHaveAttribute('spellcheck', 'true');
});
