// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.3: paste from Word in the browser, the clipboard event carrying a capture of
// the corpus, the status line's report, one undo step, the download passing the oracle.
import { test, expect, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { openFixture, download, fixture, fixturePath, comparePackages, caretInBlock, caretAtBlockEnd, capture } from './helpers.js';

/** Dispatch a paste event with `text/html` on the focused editor, as a browser would. */
async function pasteHtml(page: Page, html: string): Promise<void> {
  await page.evaluate((h) => {
    const dt = new DataTransfer();
    dt.setData('text/html', h);
    dt.setData('text/plain', 'plain text stand-in');
    const target = document.activeElement?.closest('.ProseMirror') ?? document.querySelector('.ProseMirror[aria-label="Document"]')!;
    target.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, html);
}

test('a capture of the whole sample document pastes: headings, marks, both lists with labels, the table; the status line reports; one undo step; the download passes the oracle', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  await pasteHtml(page, await capture('all'));
  const status = page.getByTestId('status');
  await expect(status).toContainText('Pasted');
  await expect(status).toContainText('2 lists');
  await expect(status).toContainText('1 table');
  await expect(status).toContainText('dropped');
  await expect(status).toContainText('added the Heading2');
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.filter({ hasText: '1. Headings and body text' })).toHaveClass(/s-Heading1/);
  await expect(paragraphs.filter({ hasText: 'A second-level heading' })).toHaveClass(/s-Heading2/);
  // The caret sat at the start of "Some text.": the first pasted paragraph's text takes the caret's paragraph (its own style goes with the join) and the last pasted paragraph takes "Some text.", as any paste fits.
  await expect(paragraphs.filter({ hasText: 'Paste samples for the docx4j editor' })).toHaveCount(1);
  await expect(paragraphs.filter({ hasText: 'highlighted word.Some text.' })).toHaveCount(1);
  const inline = paragraphs.filter({ hasText: 'This paragraph has' });
  await expect(inline.locator('strong')).toHaveCount(2);   // "bold", and "bold italic together"
  await expect(inline.locator('u')).toHaveText('underlined');
  await expect(inline.locator('s')).toHaveText('struck through');
  await expect(inline.locator('sup')).toHaveText('script');
  await expect(inline.locator('sub')).toHaveText('script');
  await expect(paragraphs.filter({ hasText: 'A line break follows' }).locator('br')).toHaveCount(1);
  await expect(paragraphs.filter({ hasText: 'A line break follows' }).locator('.tab')).toHaveCount(1);
  // The direct formatting kept (ED-003 section 8.6, F5): the centred, indented paragraph in 14-point red Arial, the highlighted word.
  const centred = paragraphs.filter({ hasText: 'keeps the text and drops the rest' });
  await expect(centred).toHaveAttribute('style', /text-align: center/);
  expect(await centred.evaluate((el) => getComputedStyle(el).textAlign)).toBe('center');
  expect(await centred.locator('span').first().evaluate((el) => [getComputedStyle(el).fontSize, getComputedStyle(el).color])).toEqual(['18.6667px', 'rgb(255, 0, 0)']);
  expect(await paragraphs.filter({ hasText: 'highlighted word' }).locator('span').filter({ hasText: /^highlighted word\./ }).first().evaluate((el) => getComputedStyle(el).backgroundColor)).toBe('rgb(255, 255, 0)');
  await expect(status).toContainText('dropped spacing (9)');   // margin-bottom and line-height on the cell paragraphs, still counted
  // The lists carry live labels: bullets on the bulleted items, numbers on the numbered ones, nested levels indented.
  await expect(paragraphs.filter({ hasText: 'First bullet' }).locator('.list-label')).toHaveCount(1);
  await expect(paragraphs.filter({ hasText: 'Second item' }).locator('.list-label')).toHaveText(/2\.?/);
  await expect(paragraphs.filter({ hasText: 'A nested item under the second' }).locator('.list-label')).toHaveText(/a\.?/);
  // The fixture has a table of its own; the pasted one is the second.
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] table')).toHaveCount(2);
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] table td').filter({ hasText: 'Name' })).toHaveCount(1);
  // One undo step takes it all back.
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+z');
  await expect(paragraphs.filter({ hasText: '1. Headings and body text' })).toHaveCount(0);
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] table')).toHaveCount(1);
  await page.keyboard.press('Control+y');
  await expect(paragraphs.filter({ hasText: '1. Headings and body text' })).toHaveCount(1);
  // The download: the main part, the numbering part (two definitions made) and app.xml changed, nothing else.
  const saved = await download(page);
  expect(comparePackages(await fixture('tracked-changes.docx'), saved, ['word/document.xml', 'word/numbering.xml', 'word/styles.xml', 'docProps/app.xml']).filter((d) => d.reason !== 'added')).toEqual([]);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  let tables = 0;
  session.doc.forEach((n) => { if (n.type.name === 'table') tables++; });
  expect(tables).toBe(2);
});

test("the editor's own clipboard is pasted exactly, not converted", async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
  await pasteHtml(page, '<p data-pm-slice="1 1 []">from the editor</p>');
  await expect(paragraphs.nth(4)).toHaveText('from the editorSome text.');
  await expect(page.getByTestId('status')).not.toContainText('Pasted');
});

/** Dispatch a paste event carrying the given flavours on the focused editor. */
async function pasteFlavours(page: Page, data: Record<string, string>): Promise<void> {
  await page.evaluate((d) => {
    const dt = new DataTransfer();
    for (const [type, value] of Object.entries(d)) dt.setData(type, value);
    const target = document.activeElement?.closest('.ProseMirror') ?? document.querySelector('.ProseMirror[aria-label="Document"]')!;
    target.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, data);
}

test("a terminal's HTML (Konsole's capture) pastes as its plain text, a paragraph per line (ED-005 section 12.67)", async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  const html = await readFile(fixturePath('paste/terminal/konsole-26.04.html'), 'utf8');
  const text = await readFile(fixturePath('paste/terminal/konsole-26.04.txt'), 'utf8');
  await pasteFlavours(page, { 'text/html': html, 'text/plain': text });
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.filter({ hasText: /^wl-paste --list-types$/ })).toHaveCount(1);
  await expect(paragraphs.filter({ hasText: /^UTF8_STRING$/ })).toHaveCount(1);
  await expect(paragraphs.filter({ hasText: /^--list-types$/ })).toHaveCount(0);
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] span[style*="color"]')).toHaveCount(0);
});

test('a link in pasted HTML is kept in Word\'s look, and the download carries it (ED-005 section 12.67)', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  await pasteHtml(page, '<p>Read <a href="https://example.com/guide">the guide</a> or <a href="#_Toc1">the heading</a>.</p>');
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.locator('a.link', { hasText: 'the guide' })).toHaveCount(1);
  await expect(body.locator('a.link')).toHaveCount(1);
  await expect(page.getByTestId('status')).toContainText('dropped 1 link');
  const saved = await download(page);
  const { EditorDocument, hyperlinkTarget } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  const links: [string, string][] = [];
  session.doc.descendants((n) => {
    const link = n.marks.find((m) => m.type.name === 'link');
    if (n.isText && link) links.push([n.text!, hyperlinkTarget(session, link.attrs['rId'] as string)]);
    return true;
  });
  expect(links).toEqual([['the guide', 'https://example.com/guide']]);
});

test('a web address pasted as plain text is a link, its text the address; what is typed after it is not; the download carries it (ED-005 section 12.121)', async ({ page }) => {
  const address = 'https://example.com/guide?a=1&b=2';
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  await page.evaluate((text) => {
    const dt = new DataTransfer();
    dt.setData('text/plain', text);
    const target = document.activeElement?.closest('.ProseMirror') ?? document.querySelector('.ProseMirror[aria-label="Document"]')!;
    target.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, address);
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.locator('a.link')).toHaveText([address]);
  await page.keyboard.type(' next');
  await expect(body.locator('p', { hasText: `${address} next` })).toHaveCount(1);
  await expect(body.locator('a.link')).toHaveText([address]);
  const saved = await download(page);
  const { EditorDocument, hyperlinkTarget } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  const links: [string, string, string | undefined][] = [];
  session.doc.descendants((n) => {
    const link = n.marks.find((m) => m.type.name === 'link');
    if (n.isText && link) links.push([n.text!, hyperlinkTarget(session, link.attrs['rId'] as string), n.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'] as string | undefined]);
    return true;
  });
  expect(links).toEqual([[address, address, 'Hyperlink']]);
  // what was typed after the link has the paragraph's formatting, not the link's look
  const runs: [string, string | undefined][] = [];
  session.doc.child(1).forEach((n) => { runs.push([n.text ?? n.type.name, n.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'] as string | undefined]); });
  expect(runs).toEqual([[address, 'Hyperlink'], [' next', undefined]]);
});
