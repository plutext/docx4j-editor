// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.3: pasted Office JS in the console through the `Word` shim.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, download, fixture, comparePackages, caretInBlock, setConsoleScript } from './helpers.js';

async function consoleReady(page: Page, name: string) {
  await openFixture(page, name);
  await page.getByTestId('perspective-developer').click();
  const console = page.getByTestId('panel-console');
  await expect(console.locator('.console-editor .monaco-editor')).toBeVisible();
  await expect(page.getByTestId('console-types')).toContainText('types loaded');
  return console;
}

async function setScript(page: Page, _console: ReturnType<Page['getByTestId']>, script: string) {
  await setConsoleScript(page, script);
}

test('an add-in script runs unchanged through Word.run: it edits the document as one undo step, and the download carries it', async ({ page }) => {
  const console = await consoleReady(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  await setScript(page, console, `await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text,items/style');
  await context.sync();
  for (const p of paragraphs.items) if (p.style === 'Heading 1') p.insertText(' (reviewed)', 'End');
  const sel = context.document.getSelection();
  sel.insertText('Selected: ', 'Start');
  await context.sync();
});
return Word.InsertLocation.end;`);
  await page.getByTestId('console-run').click();
  const output = page.getByTestId('console-output');
  await expect(output).toContainText('End');
  await expect(output).not.toContainText('fails in Word');
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  const heading = paragraphs.filter({ hasText: 'This is style Heading 1' }).first();
  await expect(heading).toHaveText(/\(reviewed\)$/);
  await expect(paragraphs.nth(4)).toHaveText('Selected: Some text.');
  const saved = await download(page);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  expect(session.doc.child(4).textContent).toBe('Selected: Some text.');
  expect(comparePackages(await fixture('tracked-changes.docx'), saved, ['word/document.xml', 'docProps/app.xml', 'docProps/core.xml', 'word/settings.xml']).filter((d) => d.name !== 'docProps/app.xml')).toEqual([]);
  // One undo step for the whole batch.
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
  await expect(heading).not.toHaveText(/\(reviewed\)/);
});

test('a member the subset lacks is reported at its line before the run and thrown by name at the run; the lint offers the load calls', async ({ page }) => {
  const console = await consoleReady(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  await setScript(page, console, `await Word.run(async (context) => {
  const body = context.document.body;
  for (const p of body.paragraphs.items) console.log(p.text);
  body.insertFileFromBase64('AAAA', 'End');
});`);
  // Check: the static pass names the member and the line, the lint the unloaded read.
  await page.getByTestId('console-check').click();
  const output = page.getByTestId('console-output');
  await expect(output).toContainText('line 4: not in the subset');
  await expect(output).toContainText('insertFileFromBase64');
  await expect(output).toContainText('line 3: `p.text` is read without `load(\'items/text\')`');
  await expect(page.getByTestId('console-insert-loads')).toBeVisible();
  // The run: the reads print, then the proxy throws NotSupportedError naming the member.
  await page.getByTestId('console-run').click();
  await expect(output).toContainText('Some text.');
  await expect(output).toContainText('NotSupportedError');
  await expect(output).toContainText('Body.insertFileFromBase64');
  // Insert load calls: the load and the sync land before the loop, and the lint is satisfied.
  await page.getByTestId('console-insert-loads').click();
  const lines = console.locator('.console-editor .view-lines');
  await expect(lines).toContainText("paragraphs.load('items/text')");
  await expect(lines).toContainText('await context.sync();');
  await page.getByTestId('console-check').click();
  await expect(output).not.toContainText('fails in Word');
  // A name used outside its callback is a type error, said as such (not "nothing to report").
  await setScript(page, console, `await Word.run(async (context) => { await context.sync(); });\ncontext.document.body.paragraphs;`);
  await page.getByTestId('console-check').click();
  await expect(output).toContainText("line 2: type error: Cannot find name 'context'");
});

test('the console offers the portable selection: range and font', async ({ page }) => {
  const console = await consoleReady(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  // A caret is an empty range: an insertion at its start does not widen it (the engine's semantics, as Office JS's), so the paragraph is what shows the text.
  await setScript(page, console, `selection.range.insertText('Range: ', 'Start');
selection.font.italic = true;
return [selection.range.text, selection.paragraphs[0].text, selection.paragraphs.length];`);
  await page.getByTestId('console-run').click();
  await expect(page.getByTestId('console-output')).toContainText('"Range: Some text."');
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).toHaveText('Range: Some text.');
});
