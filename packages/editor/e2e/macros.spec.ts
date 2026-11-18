// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.4: macros in the browser and in the document.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, download, caretInBlock, setConsoleScript } from './helpers.js';
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

/** Open the Macros menu (a details element: a click on its summary toggles it, so open it only when it is closed). */
async function openMacrosMenu(page: Page) {
  const details = page.locator('details.macros-menu');
  if (!(await details.evaluate((d) => (d as HTMLDetailsElement).open))) await page.getByTestId('macros-menu').click();
}

const SCRIPT = `for (const p of selection.paragraphs) p.insertText(' [macro]', 'End');
return selection.paragraphs.length;`;

async function saveFromConsole(page: Page, name: string, home: 'browser' | 'document') {
  await page.getByTestId('perspective-developer').click();
  const console = page.getByTestId('panel-console');
  await expect(console.locator('.console-editor .monaco-editor')).toBeVisible();
  await setConsoleScript(page, SCRIPT);
  await page.getByTestId('console-save-macro').click();
  const dialog = page.getByTestId('save-macro-dialog');
  await dialog.getByTestId('macro-name').fill(name);
  if (home === 'document') await dialog.getByTestId('macro-home-document').check();
  await dialog.getByTestId('macro-save').click();
  await expect(dialog).toBeHidden();
}

test('a console script saved as a browser macro runs from the menu and from the palette as one undo step, and can be deleted', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await saveFromConsole(page, 'Mark it', 'browser');
  await expect(page.getByTestId('status')).toContainText('saved in this browser');
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  // From the menu.
  await caretInBlock(page, 4);
  await openMacrosMenu(page);
  const entry = page.locator('[data-testid="macro-entry"][data-home="browser"]', { hasText: 'Mark it' });
  await entry.getByTestId('macro-run').click();
  await expect(page.getByTestId('status')).toContainText('Macro “Mark it” ran: 1');
  await expect(paragraphs.nth(4)).toHaveText('Some text. [macro]');
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
  // From the palette.
  await caretInBlock(page, 6);
  await page.keyboard.press('Control+Shift+p');
  const palette = page.getByTestId('palette');
  await palette.getByLabel('Command', { exact: true }).fill('Macro: Mark');
  await expect(palette.locator('li[data-macro="Mark it"] .badge')).toHaveText('browser');
  await page.keyboard.press('Enter');
  await expect(palette).toBeHidden();
  await expect(paragraphs.nth(6)).toHaveText(/ \[macro\]$/);
  // It survives a reload, and Delete removes it.
  await page.reload();
  await openMacrosMenu(page);
  await expect(page.locator('[data-testid="macro-entry"]', { hasText: 'Mark it' })).toBeVisible();
  await page.locator('[data-testid="macro-entry"]', { hasText: 'Mark it' }).getByTestId('macro-delete').click();
  await expect(page.locator('[data-testid="macro-entry"]', { hasText: 'Mark it' })).toHaveCount(0);
});

test('a macro saved in the document travels with the file: the reopened document lists it and shows its source before its first run', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await saveFromConsole(page, 'Shipped', 'document');
  await expect(page.getByTestId('status')).toContainText('saved in the document');
  await openMacrosMenu(page);
  await expect(page.locator('[data-testid="macro-entry"][data-home="document"]', { hasText: 'Shipped' })).toBeVisible();
  // The download carries the part; a fresh page (no trust record) reopens it.
  const saved = await download(page);
  const { EditorDocument, readDocumentMacros } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  expect((await readDocumentMacros(session)).map((m) => m.name)).toEqual(['Shipped']);
  const path = join(tmpdir(), `macros-${Date.now()}.docx`);
  await writeFile(path, saved);
  await page.evaluate(() => localStorage.clear());
  await page.goto('/');
  await page.getByTestId('open').setInputFiles(path);
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await openMacrosMenu(page);
  const entry = page.locator('[data-testid="macro-entry"][data-home="document"]', { hasText: 'Shipped' });
  await expect(entry).toBeVisible();
  await caretInBlock(page, 4);
  await entry.getByTestId('macro-run').click();
  // Shown first: the source, Cancel changes nothing.
  const dialog = page.getByTestId('macro-source-dialog');
  await expect(dialog.getByTestId('macro-source')).toContainText('[macro]');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
  // Run: it runs, and the next run asks nothing.
  await openMacrosMenu(page);
  await entry.getByTestId('macro-run').click();
  await dialog.getByTestId('macro-run-anyway').click();
  await expect(paragraphs.nth(4)).toHaveText('Some text. [macro]');
  await openMacrosMenu(page);
  await entry.getByTestId('macro-run').click();
  await expect(dialog).toBeHidden();
  await expect(paragraphs.nth(4)).toHaveText('Some text. [macro] [macro]');
  // Deleted from the document: the menu no longer lists it.
  await openMacrosMenu(page);
  await entry.getByTestId('macro-delete').click();
  await expect(entry).toHaveCount(0);
});
