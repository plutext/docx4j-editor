// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.5 (step F4): the Paragraph dialog's Tabs section writes the paragraph's own
// stops and the tabs follow at once; the default interval is a settings-part write from Options;
// the download carries both and nothing else changes.
import { test, expect } from '@playwright/test';
import { ZipPartStore } from '@docx4j/core-ts';
import { openFixture, caretInBlock, download, fixture, comparePackages } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

test('a stop set in the dialog moves the tab to it, removing it restores the default, Options sets the default interval, the download carries the paragraph and the settings part', async ({ page }) => {
  await openFixture(page, 'headers-footers.docx');
  // Block 3 holds two tabs and no custom stops (render.spec.ts measures it against the defaults).
  const tabs = page.locator(`${body} > p`).nth(3).locator('.tab');
  await expect(tabs).toHaveCount(2);
  const edges = async (): Promise<number[]> => tabs.evaluateAll((els) => {
    const column = els[0]!.closest('.ProseMirror')!.getBoundingClientRect().left;
    return els.map((el) => el.getBoundingClientRect().right - column);
  });
  await caretInBlock(page, 3);
  await page.getByRole('button', { name: 'Paragraph settings' }).click();
  const dialog = page.getByRole('dialog', { name: 'Paragraph' });
  await expect(dialog.getByTestId('tab-stop')).toHaveCount(0);
  await dialog.getByTestId('tab-add').click();
  const row = dialog.getByTestId('tab-stop').first();
  await expect(row).toHaveAttribute('data-origin', 'own');
  await row.getByLabel('Position').fill('100');
  await row.getByLabel('Alignment').selectOption('left');
  await row.getByLabel('Leader').selectOption('dot');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);
  // 100pt is 133.33px from the column's edge: the first tab ends there, with its leader.
  await expect.poll(async () => (await edges())[0]).toBeCloseTo(133.33, 0);
  await expect(tabs.first()).toHaveAttribute('data-leader', 'dot');
  // The default interval from Options: 72pt (96px). The second tab, past the custom stop, ends at a multiple of it.
  await page.getByLabel('Options').click();
  const interval = page.getByTestId('option-default-tab-stop');
  await expect(interval).toHaveValue('36');
  await interval.fill('72');
  await interval.press('Enter');
  await expect(page.getByTestId('status')).toContainText('Default tab stop: 72pt');
  await expect.poll(async () => { const e = (await edges())[1]!; return Math.abs(e / 96 - Math.round(e / 96)); }).toBeLessThan(0.03);
  await expect.poll(async () => (await edges())[0]).toBeCloseTo(133.33, 0);
  await page.getByLabel('Options').click();   // the pane closed again, off the toolbar
  // Removing the stop: the first tab goes back to the defaults, now 96px.
  await caretInBlock(page, 3);
  await page.getByRole('button', { name: 'Paragraph settings' }).click();
  await expect(dialog.getByTestId('tab-stop')).toHaveCount(1);
  await expect(dialog.getByTestId('tab-stop').first().getByLabel('Position')).toHaveValue('100');
  await dialog.getByTestId('tab-stop').first().getByRole('button', { name: 'Remove this stop' }).click();
  await expect(dialog.getByTestId('tab-stop')).toHaveCount(0);
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(async () => { const e = (await edges())[0]!; return Math.abs(e / 96 - Math.round(e / 96)); }).toBeLessThan(0.03);
  // The download: the main part (the paragraph's stops came and went, so it was rebuilt) and the settings part, nothing else.
  const saved = await download(page);
  expect(comparePackages(await fixture('headers-footers.docx'), saved, ['word/document.xml', 'word/settings.xml', 'docProps/app.xml'])).toEqual([]);
  const settings = new TextDecoder().decode(new ZipPartStore(saved).loadSync('word/settings.xml'));
  expect(settings).toMatch(/<w:defaultTabStop w:val="1440"\/>/);
});

test('a right stop ends the following text at the stop and typing after the tab moves the text left; a centre stop centres it; a decimal stop puts the separator there (ED-003 section 11.20)', async ({ page }) => {
  await openFixture(page, 'headers-footers.docx');
  const paragraph = page.locator(`${body} > p`).nth(3);
  const tabs = paragraph.locator('.tab');
  await expect(tabs).toHaveCount(2);
  const column = async (): Promise<number> => paragraph.evaluate((el) => el.closest('.ProseMirror')!.getBoundingClientRect().left);
  const firstRight = async (): Promise<number> => (await tabs.first().evaluate((el) => el.getBoundingClientRect().right)) - (await column());
  const secondLeft = async (): Promise<number> => (await tabs.nth(1).evaluate((el) => el.getBoundingClientRect().left)) - (await column());
  const setStops = (stops: unknown) => page.evaluate((s) => (window as unknown as { docx4jEditor: { run(n: string, a: unknown): boolean } }).docx4jEditor.run('paragraph.tabStops', { stops: s }), stops);
  await caretInBlock(page, 3);
  // The block is "<tab>private List<SectionWrapper> sections; <tab>": the text after the first tab is about 313px wide. A right stop at 200pt (267px) has less room than that, so the tab collapses to its minimum as Word's does; at 300pt (400px) the text fits and ends there, so the second tab starts there.
  expect(await setStops([{ position: 200, alignment: 'right', leader: 'dot' }])).toBe(true);
  await expect.poll(firstRight).toBeLessThan(2);
  expect(await setStops([{ position: 300, alignment: 'right', leader: 'dot' }])).toBe(true);
  await expect.poll(secondLeft).toBeCloseTo(400, 0);
  const before = await firstRight();
  expect(before).toBeGreaterThan(20);
  // Typing right after the tab makes the text longer: it still ends at the stop, so the tab shrinks and the text moves left.
  await caretInBlock(page, 3, 1);
  await page.keyboard.type('zz');
  await expect.poll(firstRight).toBeLessThan(before - 5);
  await expect.poll(secondLeft).toBeCloseTo(400, 0);
  // A centre stop: the text is centred on it.
  await caretInBlock(page, 3);
  expect(await setStops([{ position: 300, alignment: 'center' }])).toBe(true);
  await expect.poll(async () => ((await firstRight()) + (await secondLeft())) / 2).toBeCloseTo(400, 0);
  // A decimal stop, with "1.5" typed after the tab: the full stop sits at the stop.
  await caretInBlock(page, 3, 1);
  await page.keyboard.type('1.5');
  await caretInBlock(page, 3);
  expect(await setStops([{ position: 300, alignment: 'decimal' }])).toBe(true);
  /** Where the first full stop after the first tab is, from the column's edge. */
  const separatorAt = async (): Promise<number> => paragraph.evaluate((el) => {
    const columnLeft = el.closest('.ProseMirror')!.getBoundingClientRect().left;
    const tab = el.querySelector('.tab')!;
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let node: Text | null;
    while ((node = walker.nextNode() as Text | null)) {
      if (!(tab.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)) continue;
      const i = node.data.indexOf('.');
      if (i < 0) continue;
      const range = document.createRange();
      range.setStart(node, i);
      range.setEnd(node, i + 1);
      return range.getBoundingClientRect().left - columnLeft;
    }
    return -1;
  });
  await expect.poll(separatorAt).toBeCloseTo(400, 0);
  // Text after the separator that wraps: the separator stays at the stop, as in Word. The wrap was read as
  // "no separator", which collapsed the tab (ED-005 section 12.42: CI's fonts showed it).
  await caretInBlock(page, 3, 4);
  await page.keyboard.type(' and then enough words after the separator to run well past the end of this line');
  await expect.poll(separatorAt).toBeCloseTo(400, 0);
});
