// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-009: a kept mc:AlternateContent whose selected branch holds text is shown as that text, editable;
// a character typed into it is saved into the Choice, the Fallback as it was.
import { test, expect } from '@playwright/test';
import { ZipPartStore } from '@docx4j/core-ts';
import { openFixture, caretInBlock, download } from './helpers.js';

test('the selected branch\'s text is shown and edited in place; the save writes the Choice, the Fallback unchanged', async ({ page }) => {
  await openFixture(page, 'mc/mc-branches.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(editor).toContainText('The Earth is flat');
  await expect(editor).not.toContainText('The Earth is round');
  await expect(editor).toContainText('I prefer w14');
  await expect(editor).toContainText('bold ying plain ying');
  // One opaque glyph: the element with no Choice the engine understands and no Fallback selects nothing (core-ts CR-029) and
  // is kept whole; every other element shows its branch (the picture is an image).
  await expect(editor.locator('.opaque-inline.kind-AlternateContent')).toHaveCount(1);
  await caretInBlock(page, 2, 4);
  await page.keyboard.type('XX');
  await expect(editor).toContainText('The XXEarth is flat');
  const saved = await download(page);
  await expect(page.getByTestId('status')).toContainText('Saved with changes');
  const xml = new TextDecoder().decode(new ZipPartStore(saved).loadSync('word/document.xml'));
  expect(xml).toMatch(/<mc:Choice Requires="wps"><w:t[^>]*>The XXEarth is flat<\/w:t><\/mc:Choice><mc:Fallback><w:t xml:space="preserve">The Earth is round<\/w:t><\/mc:Fallback>/);
  expect(xml.split('<mc:AlternateContent').length - 1).toBe(12);
});
