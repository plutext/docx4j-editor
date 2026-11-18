// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.115, from Jason's save of docx4j's Getting Started guide: Shift+Down from the start of the paragraph
// before a heading ended after the heading's leading bookmark starts, which are drawn as nothing, and Delete took them and
// left their ends, so docx4j's TOC update refused the file. Each case downloads, since the save is where the loss showed.
import { test, expect, type Page } from '@playwright/test';
import { WordprocessingMLPackage, ZipPartStore } from '@docx4j/core-ts';
import { openBytes, download, caretInBlock } from './helpers.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (s: string): string => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const start = (id: number): string => `<w:bookmarkStart w:id="${id}" w:name="_Toc${id}"/>`;
const end = (id: number): string => `<w:bookmarkEnd w:id="${id}"/>`;

/** The guide's shape: a paragraph, then a heading whose two TOC bookmarks open before its text and close after it, then a paragraph. */
async function openGuide(page: Page): Promise<void> {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}"><w:body>`
    + `<w:p>${run('One source change: a call Docx4J.load(file, null) no longer compiles.')}</w:p>`
    + `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${start(1)}${start(2)}${run('OpenXML concepts')}${end(1)}${end(2)}</w:p>`
    + `<w:p>${run('After the heading.')}</w:p><w:sectPr/></w:body></w:document>`);
  await openBytes(page, 'guide.docx', await pkg.save());
}

/** The saved main part's markers in order, `[1` a start and `1]` an end, `|` a paragraph's end. */
async function savedMarkers(page: Page): Promise<string> {
  const xml = new TextDecoder().decode(new ZipPartStore(await download(page)).loadSync('word/document.xml'));
  expect(xml).toContain('OpenXML concepts');
  return [...xml.matchAll(/<w:bookmarkStart [^>]*w:id="(\d+)"[^>]*\/>|<w:bookmarkEnd [^>]*w:id="(\d+)"[^>]*\/>|<\/w:p>/g)]
    .map((m) => (m[1] !== undefined ? `[${m[1]}` : m[2] !== undefined ? `${m[2]}]` : '|')).join(' ');
}

/** The heading's start, and its start after the atoms it opens with: where the caret is drawn. */
function headingPositions(page: Page): Promise<{ start: number; visible: number; end: number; selection: [number, number] }> {
  return page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: import('prosemirror-model').Node; selection: { from: number; to: number } } } }).docx4jEditor;
    const doc = v.state.doc;
    let pos = 0;
    for (let k = 0; k < doc.childCount; k++) {
      const block = doc.child(k);
      if (block.textContent === 'OpenXML concepts') {
        let visible = pos + 1;
        for (let i = 0; i < block.childCount && block.child(i).type.name === 'opaqueInline'; i++) visible += block.child(i).nodeSize;
        return { start: pos + 1, visible, end: pos + block.nodeSize - 1, selection: [v.state.selection.from, v.state.selection.to] as [number, number] };
      }
      pos += block.nodeSize;
    }
    throw new Error('no heading');
  });
}

const select = (page: Page, from: number, to: number): Promise<void> => page.evaluate(([a, b]) => {
  (window as unknown as { docx4jEditor: { caretAt(pos: number, to?: number): void } }).docx4jEditor.caretAt(a!, b);
}, [from, to] as const);

test('Shift+Down from the paragraph before a heading, then Delete: the heading keeps its bookmarks (ED-005 section 12.115)', async ({ page }) => {
  await openGuide(page);
  await caretInBlock(page, 0, 0);
  const { start: heading, end } = await headingPositions(page);
  // Each press's selection reaches the state when the browser reports it: wait for it before deciding on another.
  for (let n = 0; n < 6; n++) {
    const to = (await headingPositions(page)).selection[1];
    if (to >= heading) break;
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForFunction((was) => (window as unknown as { docx4jEditor: { state: { selection: { to: number } } } }).docx4jEditor.state.selection.to !== was, to);
  }
  const reached = (await headingPositions(page)).selection[1];
  expect(reached).toBeGreaterThanOrEqual(heading);
  expect(reached).toBeLessThan(end);
  await page.keyboard.press('Delete');
  await expect(page.locator('.ProseMirror[aria-label="Document"]')).not.toContainText('One source change');
  expect(await savedMarkers(page)).toBe('[1 [2 1] 2] | |');
});

test('a selection that ends at the heading\'s text, deleted: the starts stay and the caret is at the text; Undo is one step', async ({ page }) => {
  await openGuide(page);
  const { start, visible } = await headingPositions(page);
  await select(page, 1, visible);
  await page.keyboard.press('Delete');
  const after = await headingPositions(page);
  expect(after.visible).toBe(after.start + 2);
  expect(after.selection).toEqual([after.visible, after.visible]);
  expect(await savedMarkers(page)).toBe('[1 [2 1] 2] | |');
  // The save left the focus on Download.
  await page.locator('.ProseMirror[aria-label="Document"]').focus();
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.ProseMirror[aria-label="Document"]')).toContainText('One source change');
  expect((await headingPositions(page)).start).toBe(start);
  expect(await savedMarkers(page)).toBe('| [1 [2 1] 2] | |');
});

test('typing over such a selection: the starts stay after what was typed', async ({ page }) => {
  await openGuide(page);
  const { visible } = await headingPositions(page);
  await select(page, 1, visible);
  await page.keyboard.type('X');
  expect(await savedMarkers(page)).toBe('[1 [2 1] 2] | |');
});

test('cut and pasted at the end: the starts stay in the heading and the paste brings no second copy', async ({ page }) => {
  await openGuide(page);
  const { visible } = await headingPositions(page);
  await select(page, 1, visible);
  await page.evaluate(() => {
    const w = window as unknown as { __cut?: string };
    // On the document, so after ProseMirror's handler on the editor has written the slice.
    document.addEventListener('cut', (e) => { w.__cut = e.clipboardData?.getData('text/html') ?? ''; }, { once: true });
  });
  await page.keyboard.press('ControlOrMeta+x');
  await page.waitForFunction(() => (window as unknown as { __cut?: string }).__cut !== undefined);
  const html = await page.evaluate(() => (window as unknown as { __cut?: string }).__cut ?? '');
  expect(html).toContain('One source change');
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { content: { size: number } } }; caretAt(pos: number): void } }).docx4jEditor;
    v.caretAt(v.state.doc.content.size - 1);
  });
  await page.evaluate((h) => {
    const dt = new DataTransfer();
    dt.setData('text/html', h);
    let event = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
    // Firefox's constructor ignores `clipboardData`, taking one flavour as `dataType` and `data`.
    if (event.clipboardData?.getData('text/html') !== h) event = new ClipboardEvent('paste', { dataType: 'text/html', data: h, bubbles: true, cancelable: true } as ClipboardEventInit);
    document.querySelector('.ProseMirror[aria-label="Document"]')!.dispatchEvent(event);
  }, html);
  await expect(page.locator('.ProseMirror[aria-label="Document"]')).toContainText('One source change');
  const markers = await savedMarkers(page);
  expect(markers.startsWith('[1 [2 1] 2] |')).toBe(true);
  expect(markers.match(/\[\d/g)).toHaveLength(2);
});
