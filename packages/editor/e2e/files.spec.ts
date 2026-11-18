// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.5 tests 2 and 6: an edit session end to end, dirty state, and recovery.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, fixturePath, comparePackages, caretInBlock } from './helpers.js';

test('edit, save, reopen: the download passes the oracle and carries the edits', async ({ page }) => {
  const name = 'tracked-changes.docx';
  await openFixture(page, name);
  await expect(page.getByTestId('name')).toHaveText(name);
  await caretInBlock(page, 4, 10);   // end of "Some text."
  await page.keyboard.type(' and more');
  await page.keyboard.press('Enter');
  await page.keyboard.type('A new paragraph');
  await page.getByTestId('style-picker').selectOption('Heading1');
  await expect(page.getByTestId('name')).toHaveText(`${name} •`);
  const saved = await download(page);
  await expect(page.getByTestId('status')).toContainText('Saved with changes');
  await expect(page.getByTestId('name')).toHaveText(name);
  expect(comparePackages(await fixture(name), saved, ['word/document.xml', 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml')).toEqual([]);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  expect(session.doc.child(4).textContent).toBe('Some text. and more');
  expect(session.doc.child(5).textContent).toBe('A new paragraph');
  expect(session.doc.child(5).attrs['pStyle']).toBe('Heading1');
  expect(session.doc.childCount).toBe(65);
  // Reopen the download in the app.
  await page.getByTestId('close').click();
  await page.getByTestId('open').setInputFiles({ name: 'reopened.docx', mimeType: 'application/octet-stream', buffer: Buffer.from(saved) });
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(5)).toHaveText('A new paragraph');
});

test('the app.xml option off leaves the properties untouched', async ({ page }) => {
  const name = 'HelloWordOnline.docx';
  await openFixture(page, name);
  await page.getByLabel('Options').click();
  await page.getByTestId('option-app').uncheck();
  await caretInBlock(page, 0, 5);
  await page.keyboard.type('X');
  const saved = await download(page);
  expect(comparePackages(await fixture(name), saved, ['word/document22.xml'])).toEqual([]);
});

test('the status line stops saying "Saved" once the document is edited again (Jason, 2026-11-13)', async ({ page }) => {
  const name = 'tracked-changes.docx';
  await openFixture(page, name);
  await caretInBlock(page, 4, 10);
  await page.keyboard.type('!');
  await download(page);
  const status = page.getByTestId('status');
  await expect(status).toContainText('Saved with changes');
  await expect(page.getByTestId('name')).toHaveText(name);
  await caretInBlock(page, 4, 11);
  await page.keyboard.type('?');
  await expect(page.getByTestId('name')).toHaveText(`${name} •`);
  await expect(status).toHaveText('');
  // A second save says so again, and the edit after it takes that away too.
  await download(page);
  await expect(status).toContainText('Saved with changes');
  await caretInBlock(page, 4, 12);
  await page.keyboard.type('.');
  await expect(page.getByTestId('name')).toHaveText(`${name} •`);
  await expect(status).toHaveText('');
});

test('recovery: an unsaved edit survives a reload and can be recovered or discarded', async ({ page }) => {
  const name = 'tracked-changes.docx';
  await openFixture(page, name);
  await caretInBlock(page, 4, 10);
  await page.keyboard.type(' (unsaved)');
  await expect(page.getByTestId('name')).toHaveText(`${name} •`);
  await page.waitForTimeout(600);   // the session record is written 300 ms after the last update
  await page.reload();
  const recovery = page.getByTestId('recovery');
  await expect(recovery).toContainText(name);
  await page.getByTestId('recover').click();
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).toHaveText('Some text. (unsaved)');
  await expect(page.getByTestId('name')).toHaveText(`${name} •`);
  const saved = await download(page);
  await expect(page.getByTestId('status')).toContainText('Saved with changes');   // the session record is clean once this shows
  await expect(page.getByTestId('name')).toHaveText(name);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  expect(session.doc.child(4).textContent).toBe('Some text. (unsaved)');
  // Saved: no recovery offered after another reload.
  await page.reload();
  await expect(page.getByTestId('recovery')).toHaveCount(0);
  // Discard path.
  await page.getByTestId('open').setInputFiles(fixturePath(name));
  await page.getByTestId('toolbar').waitFor();
  await caretInBlock(page, 4, 10);
  await page.keyboard.type('!');
  await page.waitForTimeout(600);
  await page.reload();
  await page.getByTestId('discard').click();
  await expect(page.getByTestId('recovery')).toHaveCount(0);
  await page.reload();
  await expect(page.getByTestId('recovery')).toHaveCount(0);
});

test('Close: by default it discards at once; with the option on it asks first, for a document with unsaved changes only', async ({ page }) => {
  const name = 'tracked-changes.docx';
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const dialog = page.getByRole('dialog', { name: 'Close without saving?' });
  // Under another name each time: React reports no change for a file input whose value is the same path as before.
  const reopen = async (as: string): Promise<void> => {
    await page.getByTestId('open').setInputFiles({ name: as, mimeType: 'application/octet-stream', buffer: Buffer.from(await fixture(name)) });
    await body.waitFor();
    await page.getByTestId('toolbar').waitFor();
  };
  // The default, as Close has always been: an unsaved edit is discarded with no question.
  await openFixture(page, name);
  await caretInBlock(page, 4, 10);
  await page.keyboard.type('!');
  await expect(page.getByTestId('name')).toHaveText(`${name} •`);
  await page.getByTestId('close').click();
  await expect(body).toHaveCount(0);
  await expect(dialog).toHaveCount(0);
  // The option on: a document with nothing unsaved still closes at once.
  await page.getByLabel('Options').click();
  await page.getByTestId('option-ask-before-close').check();
  await page.keyboard.press('Escape');
  await reopen('clean.docx');
  await page.getByTestId('close').click();
  await expect(body).toHaveCount(0);
  await expect(dialog).toHaveCount(0);
  // With an unsaved edit it asks; Cancel keeps the document and the edit, and the question's OK closes it.
  await reopen('edited.docx');
  await caretInBlock(page, 4, 10);
  await page.keyboard.type('!');
  await expect(page.getByTestId('name')).toHaveText('edited.docx •');
  await page.getByTestId('close').click();
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('edited.docx has unsaved changes.');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  await expect(body.locator('> p').nth(4)).toHaveText('Some text.!');
  await page.getByTestId('close').click();
  await dialog.getByRole('button', { name: 'Close without saving' }).click();
  await expect(body).toHaveCount(0);
  // Kept per browser, and closed without saving means nothing to recover.
  await page.reload();
  await expect(page.getByTestId('recovery')).toHaveCount(0);
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-ask-before-close')).toBeChecked();
});
