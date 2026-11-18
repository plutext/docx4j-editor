// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.114: a character XML 1.0 excludes, brought in on any way in, leaves a document that saves. Each case
// downloads, since the save is what such a character stopped, and the file is opened again, since a browser refuses one.
import { test, expect, type Locator, type Page } from '@playwright/test';
import { openFixture, openBytes, download, caretInBlock } from './helpers.js';

const body = (page: Page) => page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');

/**
 * Dispatch a paste on the focused surface, as a browser would. Firefox's constructor ignores `clipboardData`: it takes one
 * flavour as `dataType` and `data`, and the event's own data takes the rest.
 */
async function paste(page: Page, data: Record<string, string>): Promise<void> {
  await page.evaluate((d) => {
    const flavours = Object.entries(d);
    const [firstType, firstValue] = flavours[0]!;
    const dt = new DataTransfer();
    for (const [type, value] of flavours) dt.setData(type, value);
    let event = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
    if (event.clipboardData?.getData(firstType) !== firstValue) {
      event = new ClipboardEvent('paste', { dataType: firstType, data: firstValue, bubbles: true, cancelable: true } as ClipboardEventInit);
      for (const [type, value] of flavours.slice(1)) event.clipboardData!.setData(type, value);
    }
    const target = document.activeElement?.closest('.ProseMirror') ?? document.querySelector('.ProseMirror[aria-label="Document"]')!;
    target.dispatchEvent(event);
  }, data);
}

/** A field's value as a paste leaves it. Playwright's `fill` takes these characters out in Firefox, where a field keeps them. */
async function setField(field: Locator, text: string): Promise<void> {
  await field.evaluate((element, value) => {
    const prototype = element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(prototype, 'value')!.set!.call(element, value);
    element.dispatchEvent(new Event('input', { bubbles: true }));
  }, text);
}

interface Held { excluded: string[]; breaks: string[]; text: string }

/** What the body holds: any excluded character, each break's kind, and its text. */
function held(page: Page): Promise<Held> {
  return page.evaluate(() => {
    const { doc } = (window as unknown as { docx4jEditor: { state: { doc: { textContent: string; descendants(f: (n: { type: { name: string }; attrs: Record<string, unknown> }) => void): void } } } }).docx4jEditor.state;
    const breaks: string[] = [];
    doc.descendants((n) => { if (n.type.name === 'hardBreak') breaks.push(n.attrs['kind'] as string); });
    const excluded = [...new Set([...doc.textContent].filter((c) => /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(c)))];
    return { excluded, breaks, text: doc.textContent };
  });
}

/** The save goes through and the file opens again, holding what the page held. */
async function savesAndOpens(page: Page): Promise<Held> {
  const before = await held(page);
  expect(before.excluded).toEqual([]);
  const bytes = await download(page);
  await openBytes(page, 'saved.docx', bytes);
  const after = await held(page);
  expect(after.breaks).toEqual(before.breaks);
  expect(after.text).toBe(before.text);
  return after;
}

for (const [name, character, kind] of [['U+000B', '\u000B', 'line'], ['U+000C', '\u000C', 'page'], ['U+000E', '\u000E', 'column']] as const) {
  test(`a plain paste of ${name} is a ${kind} break, and the document saves`, async ({ page }) => {
    await openFixture(page, 'comments-two.docx');
    await body(page).focus();
    await caretInBlock(page, 0, 3);
    await paste(page, { 'text/plain': `one${character}two` });
    await expect.poll(async () => (await held(page)).text).toContain('onetwo');
    expect((await savesAndOpens(page)).breaks).toEqual([kind]);
  });
}

test('a plain paste of U+0002 is dropped, and the document saves', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await body(page).focus();
  await caretInBlock(page, 0, 3);
  await paste(page, { 'text/plain': 'one\u0002two' });
  await expect.poll(async () => (await held(page)).text).toContain('onetwo');
  expect((await savesAndOpens(page)).breaks).toEqual([]);
});

test('a plain paste of several paragraphs', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await body(page).focus();
  await caretInBlock(page, 0, 3);
  await paste(page, { 'text/plain': 'one\u000Btwo\n\nthree\u000Cfour\u0002\n\nfive' });
  await expect.poll(async () => (await held(page)).text).toContain('five');
  expect((await savesAndOpens(page)).breaks).toEqual(['line', 'page']);
});

test('an HTML paste', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await body(page).focus();
  await caretInBlock(page, 0, 3);
  await paste(page, { 'text/html': '<p>one\u000Btwo\u0002</p>', 'text/plain': 'one\u000Btwo\u0002' });
  await expect.poll(async () => (await held(page)).text).toContain('onetwo');
  expect((await savesAndOpens(page)).breaks).toEqual(['line']);
});

test('a paste with tracking on is a tracked change that saves', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await page.getByTestId('tracking-switch').click();
  const name = page.getByTestId('tracking-name');
  await name.fill('Reviewer');
  await name.press('Enter');
  await body(page).focus();
  await caretInBlock(page, 0, 3);
  await paste(page, { 'text/plain': 'one\u000Ctwo' });
  await expect.poll(async () => (await held(page)).text).toContain('onetwo');
  expect((await savesAndOpens(page)).breaks).toEqual(['page']);
});

test('dropped text', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await page.evaluate((text) => {
    const target = document.querySelector('.ProseMirror[aria-label="Document"]')!;
    const box = target.querySelector('p')!.getBoundingClientRect();
    const dt = new DataTransfer();
    dt.setData('text/plain', text);
    target.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, clientX: box.left + 2, clientY: box.top + 2, bubbles: true, cancelable: true }));
  }, 'one\u000Btwo');
  await expect.poll(async () => (await held(page)).text).toContain('onetwo');
  expect((await savesAndOpens(page)).breaks).toEqual(['line']);
});

test('the Replace box', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await body(page).focus();
  await page.keyboard.press('Control+h');
  await page.getByTestId('find-query').fill('One');
  await setField(page.getByTestId('find-replacement'), 'one\u000Btwo');
  await page.getByTestId('find-replace-all').click();
  await expect.poll(async () => (await held(page)).text).toContain('onetwo');
  expect((await savesAndOpens(page)).breaks).toEqual(['line']);
});

test('a comment: a reply from its box, and a paste into its text', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await page.getByTestId('tracking-switch').click();
  const name = page.getByTestId('tracking-name');
  await name.fill('Reviewer');
  await name.press('Enter');
  await page.getByTestId('tracking-switch').click();
  const first = page.getByTestId('comment-bubble').first();
  await first.click();
  await first.getByTestId('comment-reply').click();
  await setField(first.getByTestId('comment-reply-text'), 'a reply\u000Bon two lines\u0007');
  await first.getByTestId('comment-post').click();
  const mine = first.locator('.comment-reply [data-testid="comment-body"][data-editable="true"]');
  await expect(mine).toContainText('a reply');
  await mine.locator('.ProseMirror').focus();
  await paste(page, { 'text/plain': 'pasted\u000Bhere' });
  await expect(mine).toContainText('pasted');
  const bytes = await download(page);
  await openBytes(page, 'saved.docx', bytes);
  await expect(page.getByTestId('comment-bubble').first()).toContainText('a reply');
});

test('a Markdown file opens, and saves, where it would not open', async ({ page }) => {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name: 'notes.md', mimeType: 'text/markdown', buffer: Buffer.from('# Title\n\nbefore\u000Cafter\n\nfirst\u000Bsecond\u0002\n') });
  await body(page).waitFor();
  await page.getByTestId('toolbar').waitFor();
  expect((await savesAndOpens(page)).breaks).toEqual(['page', 'line']);
});

for (const [name, text, breaks] of [
  ['U+000B', '# A heading\u000Bon two lines\n\n- an item\n', ['line']],
  ['U+000E', '# A heading\u000Eon two\n\n- an item\n', ['column']],
  ['U+0002', '# A heading\n\n- an item\u0002\n', []],
] as const) {
  test(`Format as Markdown still formats text that held ${name}`, async ({ page }) => {
    await openFixture(page, 'HelloWordOnline.docx');
    await body(page).focus();
    await caretInBlock(page, 0, 3);
    await paste(page, { 'text/plain': text });
    await page.getByTestId('format-markdown').click();
    await expect(page.getByTestId('status')).toContainText('Pasted');
    expect((await savesAndOpens(page)).breaks).toEqual(breaks);
  });
}

test('what a save still refuses says why: the character, and where it stands', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await page.getByTestId('tracking-switch').click();
  const name = page.getByTestId('tracking-name');
  await setField(name, 'Some\u000BOne');
  await name.press('Enter');
  await body(page).focus();
  await caretInBlock(page, 0, 3);
  await page.keyboard.type('x');
  await page.getByTestId('download').click();
  const status = page.getByTestId('status');
  await expect(status).toContainText('Not saved: Problem saving part /word/document.xml');
  await expect(status).toContainText('U+000B');
  await expect(status).toBeVisible();
});
