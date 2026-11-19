// SPDX-License-Identifier: GPL-3.0-or-later
//
// The site's sample document (the release plan's phase C, step 2): with nothing open, the drop zone offers it; it opens
// as a file without a handle, and an unedited save is the file the site serves, part by part.
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import { download, comparePackages } from './helpers.js';

const sample = join(dirname(fileURLToPath(import.meta.url)), '..', 'public', 'samples', 'hello-world.docx');

test('the drop zone opens the sample document: "Hello World" in Heading 1, then "from docx4j-editor"', async ({ page }) => {
  await page.goto('/');
  await page.getByTestId('open-sample').click();
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(editor.locator('p')).toHaveText(['Hello World', 'from docx4j-editor']);
  await expect(page.getByTestId('name')).toHaveText('hello-world.docx');
  const styles = await page.evaluate(() => {
    const view = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { attrs: Record<string, unknown> } } } } }).docx4jEditor;
    const doc = view.state.doc;
    return Array.from({ length: doc.childCount }, (_, i) => doc.child(i).attrs['pStyle'] ?? null);
  });
  expect(styles).toEqual(['Heading1', null]);
  const saved = await download(page);
  expect(comparePackages(new Uint8Array(await readFile(sample)), saved)).toEqual([]);
});
