// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 4.6, step B3: the parts panel made editable. A style's colour changed in the
// styles part and seen on the page; a numbering format changed and the labels following; a data
// part edited and a control's value following; an unknown element refused with its path; a
// projected part read only, saying where it is edited; two pages in a room, one applying a style
// the other then shows.
import { test, expect, type Page, type BrowserContext } from '@playwright/test';
import { openFixture, caretInBlock, xmlEditorValue, setXmlEditor, download, fixture, comparePackages } from './helpers.js';

async function parts(page: Page, name: string): Promise<void> {
  await openFixture(page, name);
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-parts').click();
  await page.getByTestId('panel-parts').waitFor();
}

async function openPart(page: Page, partName: string): Promise<void> {
  await page.getByTestId('parts-tree').getByRole('button', { name: partName, exact: true }).first().click();
}

/** A style's colour set in the part's text: its own `w:color` replaced, or one added to its run properties. */
function withColor(xml: string, id: string, color: string): string {
  const style = new RegExp(`<w:style [^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(xml)![0];
  const own = /<w:color [^>]*\/>/;
  return xml.replace(style, own.test(style) ? style.replace(own, `<w:color w:val="${color}"/>`) : style.replace('<w:rPr>', `<w:rPr><w:color w:val="${color}"/>`));
}

const heading = (page: Page) => page.locator('.docx-editor .ProseMirror[aria-label="Document"] > p.s-Heading1').first();

test('a theme changed in the theme part is seen on the page: a style\'s colour and a run\'s own (ED-008 section 8.8)', async ({ page }) => {
  await parts(page, 'theme-colour-run.docx');
  const body = page.locator('.docx-editor .ProseMirror[aria-label="Document"]');
  const colourOf = (text: string) => body.getByText(text, { exact: true }).evaluate((el) => getComputedStyle(el).color);
  expect(await colourOf('stale run')).toBe('rgb(79, 129, 189)');
  await openPart(page, '/word/theme/theme1.xml');
  const text = await xmlEditorValue(page, 'part-xml');
  await setXmlEditor(page, text.replace(/(<a:accent1>\s*<a:srgbClr val=")4F81BD/, '$1123456'), 'part-xml');
  await page.getByTestId('part-apply').click();
  await expect.poll(() => colourOf('stale run')).toBe('rgb(18, 52, 86)');
  // The run whose value agreed with the old theme follows it too: both are drawn from the reference.
  expect(await colourOf('agreed run')).toBe('rgb(18, 52, 86)');
});

test('a style changed in the styles part is seen on the page and saved; an unknown element is refused with its path; Revert puts the part back', async ({ page }) => {
  await parts(page, 'tracked-changes.docx');
  await openPart(page, '/word/styles.xml');
  const text = await xmlEditorValue(page, 'part-xml');
  expect(text).toContain('w:styleId="Heading1"');
  await expect(page.getByTestId('part-apply')).toBeDisabled();
  expect(await heading(page).evaluate((el) => getComputedStyle(el).color)).not.toBe('rgb(255, 0, 0)');

  // An element the object model does not know: refused, with where it is, and nothing changes.
  await setXmlEditor(page, text.replace('<w:name w:val="heading 1"/>', '<w:name w:val="heading 1"/>\n<w:bogus/>'), 'part-xml');
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-problems')).toContainText(/w:bogus in \/w:styles\[1\]\/w:style\[\d+\] would not be saved \(line \d+\)/);
  await expect(page.getByTestId('part-status')).toContainText('Not applied');
  await expect(page.getByTestId('part-xml')).toHaveAttribute('aria-invalid', 'true');

  // The colour changed: applied, the page follows.
  await setXmlEditor(page, withColor(text, 'Heading1', 'FF0000'), 'part-xml');
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-status')).toHaveText('Applied.');
  await expect.poll(() => heading(page).evaluate((el) => getComputedStyle(el).color)).toBe('rgb(255, 0, 0)');
  await expect(page.getByTestId('name')).toContainText('•');
  const saved = await download(page);
  expect(comparePackages(await fixture('tracked-changes.docx'), saved, ['word/styles.xml', 'docProps/app.xml'])).toEqual([]);

  // Revert: the part as the pane opened it, applied.
  await page.getByTestId('part-revert').click();
  await expect(page.getByTestId('part-status')).toHaveText('Applied.');
  await expect.poll(() => heading(page).evaluate((el) => getComputedStyle(el).color)).not.toBe('rgb(255, 0, 0)');
});

test('a numbering format changed: the list labels follow', async ({ page }) => {
  await parts(page, 'tracked-changes.docx');
  const labels = page.locator('.docx-editor .ProseMirror[aria-label="Document"] .list-label');
  await expect(labels.filter({ hasText: /^1\)$/ })).toHaveCount(1);
  await openPart(page, '/word/numbering.xml');
  const text = await xmlEditorValue(page, 'part-xml');
  await setXmlEditor(page, text.replace(/w:numFmt w:val="decimal"/g, 'w:numFmt w:val="upperRoman"'), 'part-xml');
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-status')).toHaveText('Applied.');
  await expect(labels.filter({ hasText: /^I\)$/ })).toHaveCount(1);
  await expect(labels.filter({ hasText: /^1\)$/ })).toHaveCount(0);
});

test('a data part edited: the bound control\'s value follows in the Template perspective', async ({ page }) => {
  await parts(page, 'opendope/binding-simple.docx');
  await openPart(page, '/customXml/item1.xml');
  await expect(page.getByTestId('part-editor')).toContainText('kept as XML');
  const text = await xmlEditorValue(page, 'part-xml');
  expect(text).toContain('the UPDATED element 1 contents');
  await setXmlEditor(page, text.replace('the UPDATED element 1 contents', 'Changed in the parts panel'), 'part-xml');
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-status')).toHaveText('Applied.');
  await page.getByTestId('perspective-template').click();
  await caretInBlock(page, 0, 26);
  await page.getByTestId('tab-control').click();
  await expect(page.getByTestId('cc-verdict')).toContainText('Changed in the parts panel');
});

test('the main document part is read only and says where it is edited; a malformed text says its line', async ({ page }) => {
  await parts(page, 'tracked-changes.docx');
  await openPart(page, '/word/document.xml');
  await expect(page.getByTestId('part-readonly')).toContainText('Reveal codes');
  await expect(page.getByTestId('part-apply')).toHaveCount(0);
  // Shown in a read-only XML editor, coloured as the editable parts are (ED-005 section 12.63); typing changes nothing.
  const shown = await xmlEditorValue(page, 'part-readonly-xml');
  expect(shown).toContain('<w:document');
  await expect(page.getByTestId('part-readonly-xml').locator('.monaco-editor .view-line span[class^="mtk"]:not(.mtk1)').first()).toBeVisible();
  await page.getByTestId('part-readonly-xml').locator('.monaco-editor textarea').focus();
  await page.keyboard.type('x');
  expect(await xmlEditorValue(page, 'part-readonly-xml')).toBe(shown);
  // Its relationships too, in a smaller one.
  expect(await xmlEditorValue(page, 'part-rels-xml')).toContain('<Relationship ');
  // Straight from the read-only main part to the settings: the editor holds the settings, never the part shown before.
  await openPart(page, '/word/settings.xml');
  const text = await xmlEditorValue(page, 'part-xml');
  expect(text).toContain('<w:settings');
  expect(text).not.toContain('<w:document');
  await setXmlEditor(page, text.replace('<w:defaultTabStop', '<w:defaultTabStop <'), 'part-xml');
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-problems')).toContainText('does not parse');
  await expect(page.getByTestId('part-problems')).toContainText(/\(line \d+\)/);
  await expect(page.getByTestId('part-status')).toContainText('Not applied');
});

/** A context whose pages have the relay option set to the preview server's mount. */
async function withRelay(context: BrowserContext): Promise<void> {
  await context.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
}

test('two pages in a room: a style applied on one is seen on the other', async ({ browser }) => {
  const contextA = await browser.newContext();
  const contextB = await browser.newContext();
  await withRelay(contextA);
  await withRelay(contextB);
  const a = await contextA.newPage();
  const b = await contextB.newPage();
  await openFixture(a, 'tracked-changes.docx');
  await a.getByTestId('share').click();
  const link = await a.getByTestId('share-link').inputValue();
  await a.keyboard.press('Escape');
  await b.goto(link);
  await b.getByTestId('join-accept').click();
  await expect(b.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  expect(await heading(b).evaluate((el) => getComputedStyle(el).color)).not.toBe('rgb(0, 128, 0)');
  await a.getByTestId('perspective-developer').click();
  await a.getByTestId('tab-parts').click();
  await openPart(a, '/word/styles.xml');
  await setXmlEditor(a, withColor(await xmlEditorValue(a, 'part-xml'), 'Heading1', '008000'), 'part-xml');
  await a.getByTestId('part-apply').click();
  await expect(a.getByTestId('part-status')).toHaveText('Applied.');
  await expect.poll(() => heading(b).evaluate((el) => getComputedStyle(el).color), { timeout: 20_000 }).toBe('rgb(0, 128, 0)');
  await contextA.close();
  await contextB.close();
});

test('Format indents a part Word wrote on one line, as one edit Ctrl+Z undoes; Apply writes it', async ({ page }) => {
  await parts(page, 'tracked-changes.docx');
  await openPart(page, '/word/stylesWithEffects.xml');
  await expect(page.getByTestId('part-editor')).toContainText('kept as XML');
  const original = await xmlEditorValue(page, 'part-xml');
  expect(original.split('\n').length).toBeLessThanOrEqual(2);
  await page.getByTestId('part-format').click();
  await expect(page.getByTestId('part-status')).toHaveText('Formatted. Ctrl+Z undoes it; Apply writes it');
  const formatted = await xmlEditorValue(page, 'part-xml');
  expect(formatted.split('\n').length).toBeGreaterThan(100);
  expect(formatted).toContain('\n  <w:docDefaults>');
  // The same XML: elements, attributes (in any order: the browser's serializer may reorder the namespace declarations)
  // and text, whitespace between elements aside.
  const canonical = (xml: string): Promise<string> => page.evaluate((text) => {
    const walk = (n: Node): string => {
      if (n.nodeType === Node.TEXT_NODE) return (n.nodeValue ?? '').trim() === '' ? '' : JSON.stringify(n.nodeValue);
      if (n.nodeType !== Node.ELEMENT_NODE) return '';
      const el = n as globalThis.Element;
      const attrs = Array.from(el.attributes).map((a) => `${a.name}=${a.value}`).sort().join(' ');
      return `<${el.tagName} ${attrs}>${Array.from(el.childNodes).map(walk).join('')}</>`;
    };
    return walk(new DOMParser().parseFromString(text, 'application/xml').documentElement);
  }, xml);
  expect(await canonical(formatted)).toBe(await canonical(original));
  expect(formatted.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
  await expect(page.getByTestId('part-apply')).toBeEnabled();
  // Ctrl+Z in the editor: the text as it was.
  await page.getByTestId('part-xml').locator('.monaco-editor textarea').focus();
  await page.keyboard.press('Control+z');
  await expect.poll(() => xmlEditorValue(page, 'part-xml')).toBe(original);
  await expect(page.getByTestId('part-apply')).toBeDisabled();
  // Formatted again and applied.
  await page.getByTestId('part-format').click();
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-status')).toHaveText('Applied.');
});
