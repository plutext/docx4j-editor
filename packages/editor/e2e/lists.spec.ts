// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.8 item 6 (step C2): bullets on and off from the toolbar, Enter continuing and
// Enter on an empty item ending, restart from the palette, a download passing the oracle with only
// the main part, the numbering part and the relationships changed.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages, caretInBlock, caretAtBlockEnd } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

test('Bullets from the toolbar on a document with no numbering part: on, Enter continues, Enter on an empty item ends, off; the download carries the new part', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  const bullets = page.getByRole('button', { name: 'Bullets' });
  const labels = page.locator(`${body} .list-label:not(.story-view .list-label)`);
  await caretInBlock(page, 0);
  await expect(bullets).toHaveAttribute('aria-pressed', 'false');
  await bullets.click();
  await expect(bullets).toHaveAttribute('aria-pressed', 'true');
  await expect(labels).toHaveCount(1);
  await expect(labels.first()).toHaveClass(/bullet/);
  // Enter at the end continues the list; typing fills the item; Enter twice ends it.
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  await page.keyboard.type('second item');
  await expect(labels).toHaveCount(2);
  await page.keyboard.press('Enter');
  await expect(labels).toHaveCount(3);
  await page.keyboard.press('Enter');
  await expect(labels).toHaveCount(2);
  await expect(bullets).toHaveAttribute('aria-pressed', 'false');
  // Bullets off on the first item.
  await caretInBlock(page, 0);
  await bullets.click();
  await expect(labels).toHaveCount(1);
  await expect(bullets).toHaveAttribute('aria-pressed', 'false');
  const saved = await download(page);
  const original = await fixture('HelloWordOnline.docx');
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const main = (await WordprocessingMLPackage.load(original)).getMainDocumentPart().partName.storeName;   // Word Online names it document22.xml
  const diffs = comparePackages(original, saved, [main, 'docProps/app.xml']);
  expect(diffs).toEqual([{ name: 'word/numbering.xml', reason: 'added' }]);
});

test('Numbering from the toolbar, the indent buttons changing the level, Ctrl+Shift+L for bullets', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const numbering = page.getByRole('button', { name: 'Numbering' });
  const labels = page.locator(`${body} .list-label:not(.story-view .list-label)`);
  const before = await labels.count();
  await caretInBlock(page, 4);   // "Some text."
  await numbering.click();
  await expect(numbering).toHaveAttribute('aria-pressed', 'true');
  await expect(labels).toHaveCount(before + 1);
  const mine = page.locator(`${body} > p`).nth(4).locator('.list-label');
  await expect(mine).toHaveText('1.');
  await page.getByRole('button', { name: 'Increase indent' }).click();
  await expect(mine).not.toHaveText('1.');
  await page.getByRole('button', { name: 'Decrease indent' }).click();
  await expect(mine).toHaveText('1.');
  await page.keyboard.press('Control+Shift+l');
  await expect(mine).toHaveClass(/bullet/);
  await expect(page.getByRole('button', { name: 'Bullets' })).toHaveAttribute('aria-pressed', 'true');
});

test('Restart numbering from the palette: the item counts from 1 again and the later items follow; the round trip touches only the main part and the numbering part', async ({ page }) => {
  await openFixture(page, 'numbering-stories.docx');
  const labels = page.locator(`${body} .list-label:not(.story-view .list-label)`);
  await expect(labels).toHaveText(['1.', '2.', '3.', '4.', '5.', '6.']);
  await caretInBlock(page, 3);   // the third body item
  await page.getByTestId('palette-button').click();
  await page.keyboard.type('Restart numbering');
  await page.keyboard.press('Enter');
  await expect(labels).toHaveText(['1.', '2.', '1.', '2.', '3.', '4.']);
  const saved = await download(page);
  expect(comparePackages(await fixture('numbering-stories.docx'), saved, ['word/document.xml', 'word/numbering.xml', 'docProps/app.xml'])).toEqual([]);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const numbering = await pkg.getMainDocumentPart().numberingDefinitionsPart!.readContents();
  expect(numbering.num!.some((n) => n.lvlOverride?.[0]?.startOverride?.val === 1 && String(n.abstractNumId.val) === '70')).toBe(true);
});
