// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8's "New documents and templates" in the browser: New asks the page size once
// and keeps it as an option; a template opened is offered for editing or as the basis of a new
// document; the Options entries show as one pane.
import { test, expect, type Page } from '@playwright/test';
import { fixture, download } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

/** The Options pane shown, whether or not a click left it open. */
async function showOptions(page: Page): Promise<void> {
  const open = await page.locator('details.options-menu').evaluate((el) => (el as HTMLDetailsElement).open);
  if (!open) await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-margins')).toBeVisible();
}

test('New asks the page size the first time, makes a document to type in, and remembers the size', async ({ page }) => {
  await page.goto('/');
  await page.getByTestId('new').click();
  const dialog = page.getByRole('dialog', { name: 'New document' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('page-size').locator('option').nth(0)).toHaveText(/^A4/);
  await expect(dialog.getByTestId('page-size').locator('option').nth(1)).toHaveText(/^Letter/);
  await dialog.getByTestId('page-size').selectOption('LETTER');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await page.locator(body).waitFor();
  await expect(page.getByTestId('name')).toHaveText('Document1.docx');
  await page.locator(body).click();
  await page.keyboard.type('Hello, new document');
  await expect(page.locator(`${body} > p`).first()).toHaveText('Hello, new document');
  await page.getByTestId('style-picker').selectOption('Heading1');
  await expect(page.locator(`${body} > p`).first()).toHaveClass(/s-Heading1/);
  // The style renders, though the document had never used it: the stylesheet gained its rule (Heading 1 is 14 pt in the default styles, the body 11 pt).
  expect(parseFloat(await page.locator(`${body} > p`).first().evaluate((el) => getComputedStyle(el).fontSize))).toBeGreaterThan(15);
  const saved = await download(page);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  expect(session.doc.child(0).textContent).toBe('Hello, new document');
  expect((session.doc.attrs['finalSectPr'] as { pgSz?: { w?: number } }).pgSz?.w).toBe(12240);
  // The size is kept as an option and the next New does not ask.
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-page-size')).toHaveValue('LETTER');
  await page.keyboard.press('Escape');
  await page.getByTestId('new').click();
  await expect(page.getByRole('dialog', { name: 'New document' })).toHaveCount(0);
  await expect(page.getByTestId('name')).toHaveText('Document2.docx');
  // The margins option (ED-007 section 8.1 item 15): Normal until chosen; Narrow gives the next New Word's 1.27 cm sides.
  await showOptions(page);
  await expect(page.getByTestId('option-margins')).toHaveValue('normal');
  await page.getByTestId('option-margins').selectOption('narrow');
  await page.keyboard.press('Escape');
  await page.getByTestId('new').click();
  await expect(page.getByTestId('name')).toHaveText('Document3.docx');
  const pgMar = await page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { attrs: { finalSectPr: { pgMar: Record<string, number> } } } } } }).docx4jEditor.state.doc.attrs.finalSectPr.pgMar);
  expect(pgMar).toMatchObject({ top: 720, right: 720, bottom: 720, left: 720, header: 708 });
  // The Options entries sit in one pane, not on top of one another.
  await page.getByLabel('Options').click();
  const boxes = await page.locator('.options-menu .menu-pane label').evaluateAll((els) => els.map((el) => Math.round(el.getBoundingClientRect().top)));
  expect(new Set(boxes).size).toBe(boxes.length);
});

test('a .dotx opened is offered for editing or as a new document; the new document keeps the parts with the document content type', async ({ page }) => {
  const { WordprocessingMLPackage, ContentTypes } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  const main = pkg.getMainDocumentPart();
  main.contentType = ContentTypes.WORDPROCESSINGML_TEMPLATE;
  pkg.contentTypeManager.addOverrideContentType(main.partName, ContentTypes.WORDPROCESSINGML_TEMPLATE);
  const dotx = await pkg.save();
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name: 'letterhead.dotx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template', buffer: Buffer.from(dotx) });
  await page.locator(body).waitFor();
  const offer = page.getByTestId('template-offer');
  await expect(offer).toContainText('dotx');
  await page.getByTestId('template-new').click();
  await page.locator(body).waitFor();
  await expect(page.getByTestId('name')).toHaveText('Document1.docx');
  await expect(page.getByTestId('template-offer')).toHaveCount(0);
  await expect(page.locator(`${body} > p`).first()).toContainText('Hello');
  const saved = await download(page);
  const made = await WordprocessingMLPackage.load(saved);
  expect(made.getMainDocumentPart().contentType).toBe(ContentTypes.WORDPROCESSINGML_DOCUMENT);
  // Opened again and edited as the template itself: the offer dismissed, the name kept.
  await page.getByTestId('open').setInputFiles({ name: 'letterhead.dotx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template', buffer: Buffer.from(dotx) });
  await expect(page.getByTestId('template-offer')).toBeVisible();
  await page.getByTestId('template-edit').click();
  await expect(page.getByTestId('template-offer')).toHaveCount(0);
  await expect(page.getByTestId('name')).toHaveText('letterhead.dotx');
});
