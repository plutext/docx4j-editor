// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.116: in a paragraph holding a comment from the file, Enter inside the commented text, a paste of
// paragraphs, a paste inside the comment and Backspace at the paragraph's start left every save refused by the
// self-check ("Block 0 differs"). Each case saves: the download is the check, since a refused save offers none.
import { test, expect, type Page } from '@playwright/test';
import { ZipPartStore } from '@docx4j/core-ts';
import { openFixture, caretInBlock, fixture } from './helpers.js';

const body = (page: Page) => page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');

/** A paste on the editor, as a browser dispatches one; Firefox's constructor takes one flavour as `dataType` and `data`. */
async function paste(page: Page, type: string, data: string): Promise<void> {
  await page.evaluate(([t, d]) => {
    const dt = new DataTransfer();
    dt.setData(t!, d!);
    let event = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
    if (event.clipboardData?.getData(t!) !== d) event = new ClipboardEvent('paste', { dataType: t, data: d, bubbles: true, cancelable: true } as ClipboardEventInit);
    document.querySelector('.ProseMirror[aria-label="Document"]')!.dispatchEvent(event);
  }, [type, data] as const);
}

/** The saved main part's text and comment markers: `[0` a start, `0]` an end, `^0` a reference, `|` a paragraph's end. */
function shapeOf(xml: string): string {
  return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:commentRangeStart w:id="(\d+)"\/>|<w:commentRangeEnd w:id="(\d+)"\/>|<w:commentReference w:id="(\d+)"\/>|<\/w:p>/g)]
    .map((m) => (m[1] !== undefined ? `"${m[1]}"` : m[2] !== undefined ? `[${m[2]}` : m[3] !== undefined ? `${m[3]}]` : m[4] !== undefined ? `^${m[4]}` : '|')).join(' ');
}

/** Download, which a refused save never offers: the main part, or the status line's words when it was refused. */
async function saved(page: Page): Promise<string> {
  const download = page.waitForEvent('download', { timeout: 10_000 }).catch(() => null);
  await page.getByTestId('download').click();
  const d = await download;
  if (!d) throw new Error(`not saved: ${await page.getByTestId('status').innerText()}`);
  return new TextDecoder().decode(new ZipPartStore(new Uint8Array(await (await import('node:fs/promises')).readFile((await d.path())!))).loadSync('word/document.xml'));
}

async function open(page: Page): Promise<void> {
  await openFixture(page, 'comments-two.docx');
  await body(page).focus();
}

// comments-two.docx: "One [0 paragraph 0] ^0" then "[1 blsgh  1] ^1".
test('Enter inside a comment from the file: the document saves, the comment over both halves (ED-005 section 12.116)', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 0, 'One par'.length);
  await page.keyboard.press('Enter');
  expect(shapeOf(await saved(page))).toBe('"One " [0 "par" | "agraph" 0] ^0 | [1 "blsgh " 1] ^1 |');
});

test('Undo after Enter inside the comment saves the file\'s own part', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 0, 'One par'.length);
  await page.keyboard.press('Enter');
  await page.keyboard.press('ControlOrMeta+z');
  const original = new TextDecoder().decode(new ZipPartStore(await fixture('comments-two.docx')).loadSync('word/document.xml'));
  expect(await saved(page)).toBe(original);
});

test('Backspace at the start of a commented paragraph joins it, its comment with it', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 1, 0);
  await page.keyboard.press('Backspace');
  expect(shapeOf(await saved(page))).toBe('"One " [0 "paragraph" 0] ^0 [1 "blsgh " 1] ^1 |');
});

test('two plain paragraphs pasted inside the comment are in it', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 0, 'One par'.length);
  await paste(page, 'text/plain', 'first\n\nsecond');
  expect(shapeOf(await saved(page))).toMatch(/^"One " \[0 "par" "first" \| "second" "agraph" 0\] \^0 \|/);
});

test('a word pasted inside the comment is in it', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 0, 'One par'.length);
  await paste(page, 'text/plain', 'x');
  expect(shapeOf(await saved(page))).toMatch(/^"One " \[0 "par" "x" "agraph" 0\] \^0 \|/);
});

test('two HTML paragraphs pasted before the comment: it goes on after them', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 0, 'One'.length);
  await paste(page, 'text/html', '<p>first</p><p>second</p>');
  expect(shapeOf(await saved(page))).toMatch(/\| "second" " " \[0 "paragraph" 0\] \^0 \| \[1 "blsgh " 1\] \^1 \|$/);
});

test('two paragraphs pasted at the commented paragraph\'s start', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 1, 0);
  await paste(page, 'text/plain', 'first\n\nsecond');
  expect(shapeOf(await saved(page))).toMatch(/\| "first" \| "second" \[1 "blsgh " 1\] \^1 \|$/);
});

test('Format as Markdown inside the comment', async ({ page }) => {
  await open(page);
  await caretInBlock(page, 0, 'One par'.length);
  await paste(page, 'text/plain', '# A heading\n\n- an item\n');
  await page.getByTestId('format-markdown').click();
  await expect(page.getByTestId('status')).toContainText('Pasted');
  const shape = shapeOf(await saved(page));
  expect(shape.startsWith('"One " [0 "par"')).toBe(true);
  expect(shape).toContain('"agraph" 0] ^0 |');
});
