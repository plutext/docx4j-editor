// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.4: a document's embedded fonts load as faces for the session and the text set
// in them renders with them; nothing is written, so the download is byte for byte.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages } from './helpers.js';

test('fonts-embedded.docx: the four Noto Sans faces load, the paragraphs set in it use it, the status line says so, the download is untouched', async ({ page }) => {
  await openFixture(page, 'fonts-embedded.docx');
  await expect(page.getByTestId('status')).toContainText('Embedded fonts loaded: Noto Sans (4 faces)');
  const faces = await page.evaluate(() => {
    const out: string[] = [];
    document.fonts.forEach((f) => { if (f.family.replace(/"/g, '') === 'Noto Sans') out.push(`${f.weight}/${f.style}:${f.status}`); });
    return out.sort();
  });
  expect(faces).toEqual(['bold/italic:loaded', 'bold/normal:loaded', 'normal/italic:loaded', 'normal/normal:loaded']);
  expect(await page.evaluate(() => document.fonts.check('16px "Noto Sans"'))).toBe(true);
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  const regular = paragraphs.filter({ hasText: 'set in Noto Sans, regular' }).locator('span').first();
  expect(await regular.evaluate((el) => getComputedStyle(el).fontFamily)).toMatch(/^"?Noto Sans"?/);
  const plain = paragraphs.filter({ hasText: "document's default font" });
  expect(await plain.evaluate((el) => getComputedStyle(el).fontFamily)).not.toMatch(/Noto Sans/);
  // The diagnostics tab lists them.
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-diagnostics').click();
  await expect(page.getByTestId('diag-fonts')).toHaveText('4 loaded');
  // Untouched: every part byte for byte, the font parts included.
  const saved = await download(page);
  expect(comparePackages(await fixture('fonts-embedded.docx'), saved)).toEqual([]);
  // Closing the document removes the faces.
  await page.getByTestId('close').click();
  expect(await page.evaluate(() => { let n = 0; document.fonts.forEach((f) => { if (f.family.replace(/"/g, '') === 'Noto Sans') n++; }); return n; })).toBe(0);
});

// The same source saved with the characters used only by Word 2010 and by Microsoft 365 Word (its Aptos embeddings stripped); the latter's Noto Sans regular and bold are a version without a format 12 cmap subtable, so only its italics need the repair.
for (const [name, repaired] of [['fonts-embedded-subset.docx', 4], ['fonts-embedded-subset-m365.docx', 2]] as const) test(`${name}: the subset faces load and are said to be a subset, a letter the subset holds draws in it, a letter it lacks falls back`, async ({ page }) => {
  await openFixture(page, name);
  await expect(page.getByTestId('status')).toContainText('Embedded fonts loaded: Noto Sans (4 faces, subset)');
  const faces = await page.evaluate(() => {
    const out: string[] = [];
    document.fonts.forEach((f) => { if (f.family.replace(/"/g, '') === 'Noto Sans') out.push(`${f.weight}/${f.style}:${f.status}`); });
    return out.sort();
  });
  expect(faces).toEqual(['bold/italic:loaded', 'bold/normal:loaded', 'normal/italic:loaded', 'normal/normal:loaded']);
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  const regular = paragraphs.filter({ hasText: 'is set in Noto Sans, regular' }).locator('span').first();
  await expect(regular).toHaveText('regular');
  expect(await regular.evaluate((el) => getComputedStyle(el).fontFamily)).toMatch(/^"?Noto Sans"?/);
  // The subset holds the letters of "regular" and "abc" and no `x` (the model test reads the cmap): drawn in the stack "Noto Sans" then monospace, `x` measures as monospace does and `a` does not.
  const widths = await page.evaluate(() => {
    const canvas = document.createElement('canvas').getContext('2d')!;
    const measure = (font: string, text: string) => { canvas.font = font; return canvas.measureText(text).width; };
    return {
      aInStack: measure('32px "Noto Sans", monospace', 'a'), aMonospace: measure('32px monospace', 'a'),
      xInStack: measure('32px "Noto Sans", monospace', 'x'), xMonospace: measure('32px monospace', 'x'),
    };
  });
  expect(widths.xInStack).toBeCloseTo(widths.xMonospace, 3);
  expect(Math.abs(widths.aInStack - widths.aMonospace)).toBeGreaterThan(0.5);
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-diagnostics').click();
  // UPSTREAM-5: the browser took the faces only because the model zeroed the cmap language field Word wrote; the diagnostics say so.
  await expect(page.getByTestId('diag-fonts')).toHaveText(`4 loaded, ${repaired} repaired for the browser (cmap (0,4) format 12 language field zeroed; cmap (3,10) format 12 language field zeroed)`);
  const saved = await download(page);
  expect(comparePackages(await fixture(name), saved)).toEqual([]);
});
