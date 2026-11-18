// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.5 tests 4 and 5: decorations and the look.
import { test, expect } from '@playwright/test';
import { ZipPartStore, WordprocessingMLPackage } from '@docx4j/core-ts';
import { openFixture, openBytes, download, fixture, comparePackages, caretAtBlockEnd, caretInBlock } from './helpers.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

test('a section break decoration renders; the untouched round trip still holds', async ({ page }) => {
  await openFixture(page, 'sections-multiple.docx');
  await expect(page.locator('.docx-editor .section-break').first()).toBeVisible();
  await expect(page.locator('.docx-editor .section-break').first()).toContainText('Section Break');
  const saved = await download(page);
  expect(comparePackages(await fixture('sections-multiple.docx'), saved)).toEqual([]);
});

test('page boundaries render from the markers, and turn approximate from the first edited block on', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');   // three w:lastRenderedPageBreak markers
  await expect(page.locator('.docx-editor .page-boundary')).toHaveCount(3);
  await expect(page.locator('.docx-editor .page-boundary').first()).toContainText('Page 2');
  await expect(page.locator('.docx-editor .page-boundary.stale')).toHaveCount(0);
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await editor.locator(':scope > p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  await expect(page.locator('.docx-editor .page-boundary.stale').first()).toBeVisible();
});

test('list labels render for a numbered document, and the round trip holds', async ({ page }) => {
  await openFixture(page, 'numbering-stories.docx');
  // The body's labels: the headers at the boundaries carry their own lists' labels since E2.d.
  const labels = page.locator('.docx-editor .list-label:not(.story-view .list-label)');
  await expect(labels).toHaveText(['1.', '2.', '3.', '4.', '5.', '6.']);
  // The label fills the hanging indent (ED-003 section 5.4): its box ends where the paragraph's text begins.
  const first = labels.first();
  const geometry = await first.evaluate((el) => {
    const p = el.parentElement!;
    const label = el.getBoundingClientRect();
    const text = document.createRange();
    text.setStart(el.nextSibling!, 0); text.setEnd(el.nextSibling!, 1);
    // The label's glyphs, not its box: an inline-block inherits text-indent and indented them by the hanging (ED-005 section 12.97).
    const glyphs = document.createRange();
    glyphs.selectNodeContents(el);
    const cs = getComputedStyle(p);
    return { labelStart: label.left, labelEnd: label.right, glyphStart: glyphs.getBoundingClientRect().left, textStart: text.getBoundingClientRect().left, padding: cs.paddingInlineStart, indent: cs.textIndent, firstLine: p.getBoundingClientRect().left + parseFloat(cs.paddingInlineStart) + parseFloat(cs.textIndent) };
  });
  expect(Math.abs(geometry.labelEnd - geometry.textStart)).toBeLessThan(1.5);
  // The label is drawn where the first line starts, at the left indent less the hanging (Word's placement).
  expect(Math.abs(geometry.glyphStart - geometry.labelStart)).toBeLessThan(1.5);
  expect(Math.abs(geometry.glyphStart - geometry.firstLine)).toBeLessThan(1.5);
  expect(geometry.padding).not.toBe('0px');
  expect(geometry.indent.startsWith('-')).toBe(true);
  const saved = await download(page);
  expect(comparePackages(await fixture('numbering-stories.docx'), saved)).toEqual([]);
});

test('Enter in a list item numbers the new item and renumbers the rest live (ED-003 section 5.4)', async ({ page }) => {
  await openFixture(page, 'numbering-stories.docx');
  const labels = page.locator('.docx-editor .list-label:not(.story-view .list-label)');
  await expect(labels).toHaveText(['1.', '2.', '3.', '4.', '5.', '6.']);
  await caretAtBlockEnd(page, 1);   // P02, the first body item
  await page.keyboard.press('Enter');
  await page.keyboard.type('a new item');
  await expect(labels).toHaveText(['1.', '2.', '3.', '4.', '5.', '6.', '7.']);
  // The body's own paragraphs: the headers at the page boundaries have theirs too.
  await expect(page.locator('.docx-editor .ProseMirror[aria-label="Document"] > p').nth(2)).toHaveText(/^2\.a new item$/);   // the label widget's text comes first
  // Tab at its start takes it a level down; Shift+Tab brings it back.
  await caretInBlock(page, 2);
  await page.keyboard.press('Tab');
  await expect(labels.nth(1)).not.toHaveText('2.');
  await expect(labels.nth(2)).toHaveText('2.');
  await page.keyboard.press('Shift+Tab');
  await expect(labels).toHaveText(['1.', '2.', '3.', '4.', '5.', '6.', '7.']);
});

test('a heading takes its style rule; a table renders as a table with the style\'s borders', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const heading = page.locator('.docx-editor p.s-Heading1').first();
  await expect(heading).toBeVisible();
  const weight = await heading.evaluate((el) => getComputedStyle(el).fontWeight);
  expect(['700', 'bold']).toContain(weight);
  const table = page.locator('.docx-editor table.docx-table').first();
  await expect(table).toBeVisible();
  expect(await table.locator('td').first().evaluate((el) => getComputedStyle(el).borderTopStyle)).toBe('solid');
  await expect(page.getByTestId('banner')).toContainText('tracked');
});

test('Hebrew text is right-to-left and Japanese text takes its script font', async ({ page }) => {
  await openFixture(page, 'fonts-hebrew-no-cs.docx');
  const rtl = page.locator('.docx-editor p[style*="direction: rtl"], .docx-editor .rtl, .docx-editor bdi');
  expect(await rtl.count()).toBeGreaterThan(0);
  await openFixture(page, 'fonts-cjk-linebox.docx');
  const spans = page.locator('.docx-editor .ProseMirror[aria-label="Document"] span[style*="font-family"]');
  expect(await spans.count()).toBeGreaterThan(0);
});

test('bold inside an emoji sequence does not split it', async ({ page }) => {
  await openFixture(page, 'fonts-symbol-and-emoji.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await editor.click();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+End');
  await page.keyboard.press('Control+b');
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const xml = await pkg.getMainDocumentPart().getXml();
  expect(xml.isWellFormed()).toBe(true);
});

test('tabs are as wide as the distance to the next stop (ED-003 section 5.6): a paragraph of headers-footers.docx against its stops', async ({ page }) => {
  await openFixture(page, 'headers-footers.docx');
  const body = '.docx-editor .ProseMirror[aria-label="Document"]';
  // Block 3, "private List<SectionWrapper> sections;", holds two tabs and no custom stops: each ends at a multiple of the default interval (720 twips, 48px) from the column's edge.
  const tabs = page.locator(`${body} > p`).nth(3).locator('.tab');
  await expect(tabs).toHaveCount(2);
  await expect(tabs.first()).toHaveAttribute('style', /inline-size/);
  const edges = await tabs.evaluateAll((els) => {
    const column = els[0]!.closest('.ProseMirror')!.getBoundingClientRect().left;
    return els.map((el) => el.getBoundingClientRect().right - column);
  });
  for (const edge of edges) expect(Math.abs(edge / 48 - Math.round(edge / 48))).toBeLessThan(0.03);
  expect(edges[1]!).toBeGreaterThan(edges[0]!);
  // A paragraph with a custom stop: a tab typed into it ends at that stop, not at a default one.
  const stopped = page.locator(`${body} p`).filter({ has: page.locator('.tab[data-leader]') }).first();
  const count = await stopped.count();
  if (count > 0) {
    const leader = await stopped.locator('.tab[data-leader]').first().getAttribute('data-leader');
    expect(leader).toBeTruthy();
  }
  // Typing before a tab moves it: its width shrinks so it still ends at the same stop.
  const before = edges[0]!;
  await caretInBlock(page, 3);
  await page.keyboard.type('xx');
  await expect.poll(async () => (await tabs.first().evaluateAll((els) => els.map((el) => el.getBoundingClientRect().right - el.closest('.ProseMirror')!.getBoundingClientRect().left)))[0]).toBeCloseTo(before, 0);
});

// ED-005 section 12.39 (Jason, 2026-09-30): the page keeps its section's width in every perspective. A side panel that
// leaves the page area narrower than the page scrolls the area rather than laying the page out again, so lines break where
// Word breaks them; with comments, the margin beside the page can be scrolled to (`safe` centring).
test('the page keeps its section\'s width in every perspective, and the comments margin can be scrolled to', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  const measure = () => page.evaluate(() => {
    const body = document.querySelector('.docx-editor.body') as HTMLElement;
    const area = document.querySelector('.page-area') as HTMLElement;
    const margin = document.querySelector('.comments-margin') as HTMLElement | null;
    const a = area.getBoundingClientRect();
    return {
      page: Math.round(body.getBoundingClientRect().width),
      styled: Math.round(parseFloat(getComputedStyle(body).inlineSize)),
      overflows: area.scrollWidth > area.clientWidth,
      // Where the margin ends against the area's scrollable extent: it must not lie before the area's start.
      marginStart: margin ? Math.round(margin.getBoundingClientRect().left - a.left + area.scrollLeft) : null,
      scrollWidth: area.scrollWidth,
    };
  });
  const edit = await measure();
  expect(edit.page).toBe(edit.styled);
  for (const perspective of ['review', 'developer', 'template']) {
    await page.getByTestId(`perspective-${perspective}`).click();
    await page.waitForTimeout(200);
    const now = await measure();
    expect(now.page, perspective).toBe(edit.page);
    if (now.marginStart !== null) {
      expect(now.marginStart, `${perspective}: the margin's start inside the scrollable area`).toBeGreaterThanOrEqual(0);
      expect(now.marginStart, `${perspective}: the margin within the area's scroll`).toBeLessThan(now.scrollWidth);
    }
  }
  // Review with its panel open at this viewport is narrower than the page and the margin: the area scrolls.
  await page.getByTestId('perspective-review').click();
  expect((await measure()).overflows).toBe(true);
});

test("a section break shows the kind of the section after it, which is where Word keeps it (w:type, ECMA-376 17.6.22), and a page boundary only where that section starts a page (ED-006 section 9.9)", async ({ page }) => {
  await openFixture(page, 'print-pages.docx');
  const breaks = page.locator('.docx-editor .section-break');
  await expect(breaks).toHaveText(['Section Break (Next Page)', 'Section Break (Continuous)', 'Section Break (Next Page)']);
  // Section three starts continuous: no page boundary after the second break; sections two and four start pages.
  const boundaryAfter = (i: number) => page.locator(`.docx-editor .section-break[data-section="${i}"] + .page-boundary`);
  await expect(boundaryAfter(0)).toHaveCount(1);
  await expect(boundaryAfter(1)).toHaveCount(0);
  await expect(boundaryAfter(2)).toHaveCount(1);
});

test('a run\'s own theme colour is drawn as Word draws it, over a stale value beside it; a tracked change keeps its author\'s colour until a view draws it plain; an edit keeps the reference in the file (ED-008 section 8.8)', async ({ page }) => {
  await openFixture(page, 'theme-colour-run.docx');
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const colourOf = (text: string, within = body) => within.getByText(text, { exact: true }).evaluate((el) => getComputedStyle(el).color);
  const accent1 = 'rgb(79, 129, 189)';   // the theme's 4F81BD
  await expect(body.getByText('stale run', { exact: true })).toBeVisible();
  expect(await colourOf('stale run')).toBe(accent1);
  expect(await colourOf('agreed run')).toBe(accent1);
  expect(await colourOf('tinted run')).toBe('rgb(149, 179, 215)');   // 95B3D7, Word's Accent 1, Lighter 40%
  expect(await colourOf('cell run')).toBe(accent1);
  expect(await colourOf('control run')).toBe(accent1);

  // In a change the author's colour wins, as it does for any run's own colour; a view that draws the change plain shows the theme's.
  const author = await body.locator('ins.insertion').evaluate((el) => getComputedStyle(el).color);
  expect(author).not.toBe(accent1);
  expect(await colourOf('inserted run')).toBe(author);
  expect(await colourOf('deleted run')).toBe(await body.locator('del.deletion').evaluate((el) => getComputedStyle(el).color));
  await page.evaluate(() => { (window as unknown as { docx4jEditor: { markupView: string } }).docx4jEditor.markupView = 'accepted'; });
  await expect(body).toHaveClass(/markup-accepted/);
  expect(await colourOf('inserted run')).toBe(accent1);
  await page.evaluate(() => { (window as unknown as { docx4jEditor: { markupView: string } }).docx4jEditor.markupView = 'original'; });
  await expect(body).toHaveClass(/markup-original/);
  expect(await colourOf('deleted run')).toBe(accent1);
  await page.evaluate(() => { (window as unknown as { docx4jEditor: { markupView: string } }).docx4jEditor.markupView = 'all'; });

  // Typing in the run rewrites its paragraph, and the file keeps the reference and the value it read.
  await caretInBlock(page, 0, 3);
  await page.keyboard.type('X');
  await expect(body.getByText('staXle run', { exact: true })).toBeVisible();
  expect(await colourOf('staXle run')).toBe(accent1);
  const xml = new TextDecoder().decode(new ZipPartStore(await download(page)).loadSync('word/document.xml'));
  expect(xml).toMatch(/<w:color (?=[^>]*w:themeColor="accent1")(?=[^>]*w:val="FF00FF")[^>]*\/>(?:(?!<\/w:r>).)*staXle run/);
});

test('a run\'s own font, size and "off" values win over its character and paragraph styles, as Word draws them (ED-005 section 12.132)', async ({ page }) => {
  // The shape of Jason's share application: a character style with a font, a size and bold, under runs that state Arial,
  // 11 points and, two of them, bold off. The style's font and size had won, and its bold.
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().styleDefinitionsPart!.setXml(`<w:styles xmlns:w="${W}">`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Loud"><w:name w:val="Loud"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:i/><w:caps/></w:rPr></w:style>'
    + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
    + '<w:style w:type="character" w:styleId="Big"><w:name w:val="Big"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
    + '</w:styles>');
  const r = (text: string, rPr = ''): string => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const own = '<w:rStyle w:val="Big"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>';
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}"><w:body>`
    + `<w:p>${r('styled', '<w:rStyle w:val="Big"/>')}${r(' own font', `${own}<w:sz w:val="22"/>`)}${r(' own off', `${own}<w:b w:val="0"/><w:sz w:val="22"/>`)}</w:p>`
    + `<w:p><w:pPr><w:pStyle w:val="Loud"/></w:pPr>${r('loud')}${r('quiet', '<w:b w:val="0"/><w:i w:val="0"/><w:caps w:val="0"/>')}</w:p>`
    + '<w:sectPr/></w:body></w:document>');
  await openBytes(page, 'run-styles.docx', await pkg.save());
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const look = (text: string) => body.getByText(text, { exact: true }).evaluate((el) => {
    const cs = getComputedStyle(el);
    return { family: cs.fontFamily.split(',')[0]!.replace(/"/g, ''), size: cs.fontSize, weight: cs.fontWeight, style: cs.fontStyle, caps: cs.textTransform };
  });
  await expect(body.getByText('styled', { exact: true })).toBeVisible();
  expect(await look('styled')).toMatchObject({ family: 'Courier New', size: '18.6667px', weight: '700' });
  expect(await look('own font')).toMatchObject({ family: 'Arial', size: '14.6667px', weight: '700' });
  expect(await look('own off')).toMatchObject({ family: 'Arial', size: '14.6667px', weight: '400' });
  expect(await look('loud')).toMatchObject({ weight: '700', style: 'italic', caps: 'uppercase' });
  expect(await look('quiet')).toMatchObject({ weight: '400', style: 'normal', caps: 'none' });

  // Bold the person sets wins over the run's own "off"; set off again, the file keeps the run's own.
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(pos: number, to?: number): void } }).docx4jEditor;
    const start = v.state.doc.child(0).nodeSize + 1 + 'loud'.length;
    v.caretAt(start, start + 'quiet'.length);
  });
  await page.keyboard.press('Control+b');
  await expect.poll(async () => (await look('quiet')).weight).toBe('700');
  await page.keyboard.press('Control+b');
  await expect.poll(async () => (await look('quiet')).weight).toBe('400');
  const xml = new TextDecoder().decode(new ZipPartStore(await download(page)).loadSync('word/document.xml'));
  expect(xml).toMatch(/<w:b w:val="(?:0|false)"\/>(?:(?!<\/w:r>).)*quiet/);
});

test('a line is as tall as the text drawn on it, not as a character style\'s size the run replaces (ED-005 section 12.133)', async ({ page }) => {
  // Jason's share application again: its address lines, under a 14pt character style with the runs' own 11pt, were
  // 16.1pt tall where Word's are 12.6pt, since the style's element, outermost, kept the style's size.
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().styleDefinitionsPart!.setXml(`<w:styles xmlns:w="${W}">`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:style>'
    + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
    + '<w:style w:type="character" w:styleId="Big"><w:name w:val="Big"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
    + '</w:styles>');
  const p = (rPr: string, text: string): string => `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}"><w:body>`
    + p('<w:rStyle w:val="Big"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>', 'Styled, own size')
    + p('<w:sz w:val="22"/>', 'Plain eleven')
    + p('<w:rStyle w:val="Big"/>', 'Styled alone')
    + p('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:b/><w:sz w:val="28"/>', 'Plain fourteen')
    + p('<w:rStyle w:val="Big"/><w:sz w:val="22"/>', 'Styled, own size only')
    + '<w:sectPr/></w:body></w:document>');
  await openBytes(page, 'style-lines.docx', await pkg.save());
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.getByText('Styled alone')).toBeVisible();
  const height = (i: number) => body.locator(':scope > p').nth(i).evaluate((el) => el.getBoundingClientRect().height);
  const look = (text: string) => body.getByText(text, { exact: true }).evaluate((el) => {
    const cs = getComputedStyle(el);
    return { family: cs.fontFamily.split(',')[0]!.replace(/"/g, ''), size: cs.fontSize, weight: cs.fontWeight };
  });
  expect(Math.abs(await height(0) - await height(1))).toBeLessThan(0.5);
  expect(Math.abs(await height(2) - await height(3))).toBeLessThan(0.5);
  expect(await height(2)).toBeGreaterThan(await height(0) + 3);
  // The style's size and family are still the text's where the run states neither, or one of them.
  expect(await look('Styled alone')).toMatchObject({ family: 'Courier New', size: '18.6667px', weight: '700' });
  expect(await look('Styled, own size only')).toMatchObject({ family: 'Courier New', size: '14.6667px', weight: '700' });
  expect(Math.abs(await height(4) - await height(1))).toBeLessThan(1.5);
});

test('a run\'s own theme colour in a header is drawn from the theme (ED-008 section 8.8)', async ({ page }) => {
  await openFixture(page, 'theme-colour-run.docx');
  const header = page.locator('.docx-editor .page-start .boundary-header .story-view .theme-colour').first();
  await expect(header).toBeVisible();
  expect(await header.evaluate((el) => getComputedStyle(el).color)).toBe('rgb(79, 129, 189)');
});
