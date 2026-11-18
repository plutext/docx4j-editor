// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.76: a picture inserted from the ribbon lands at the caret, drawn from its new part,
// and the download holds the drawing and the media part.
import { test, expect } from '@playwright/test';
import { deflateSync } from 'node:zlib';
import { openFixture, openBytes, caretInBlock, download, comparePackages, fixture } from './helpers.js';

function png(w: number, h: number): Buffer {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf: Buffer): number => { let c = 0xffffffff; for (const b of buf) c = crcTable[(c ^ b) & 0xff]! ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type: string, data: Buffer): Buffer => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type, 'ascii'), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((1 + w * 3) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = y * (1 + w * 3) + 1 + x * 3; raw[o + 2] = 255; }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

test('Insert picture: the file chosen lands at the caret as an image drawn from its part; the download holds the drawing and the media part', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4, 5);
  await expect(page.getByTestId('insert-picture')).toBeEnabled();
  await page.getByTestId('insert-picture-file').setInputFiles({ name: 'blue.png', mimeType: 'image/png', buffer: png(120, 60) });
  const image = page.locator('.docx-editor .ProseMirror[aria-label="Document"] > p').nth(4).locator('img.image');
  await expect(image).toHaveCount(1);
  await expect(image).toHaveAttribute('src', /^blob:/);
  await expect.poll(() => image.evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(120);
  expect(await image.evaluate((el) => el.classList.contains('missing'))).toBe(false);
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const xml = await pkg.getMainDocumentPart().getXml();
  expect(xml).toMatch(/<wp:inline[\s\S]*wp:docPr [^>]*name="blue"[\s\S]*<a:blip r:embed="rId\d+"/);
  expect(pkg.getPart('/word/media/image1.png')).toBeTruthy();
});

// ED-005 section 12.88: a picture control bound from the Data tab shows its node's image at once.
test('a picture control inserted and bound from the Data tab shows the node\'s image', async ({ page }) => {
  await openFixture(page, 'opendope/xpaths-qualified-attributes.docx');
  await page.getByTestId('perspective-template').click();
  // An empty paragraph at the end of the body, a picture control in it.
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { content: { size: number } } }; caretAt(a: number): void; run(name: string, args: unknown): boolean } }).docx4jEditor;
    v.caretAt(v.state.doc.content.size - 1);
    v.run('insertParagraph', undefined);
    v.run('contentControl.insert', { kind: 'Picture' });
  });
  await expect(page.locator('.docx-editor.body .sdt-inline[data-kind="Picture"], .docx-editor.body .sdt-block[data-kind="Picture"]').first()).toBeAttached();
  // Word's form (section 12.126): the placeholder picture, docx4j's 200 by 200 pixels, until the bind puts the node's in its place.
  const picture = page.locator('.docx-editor.body [data-kind="Picture"] img').first();
  await expect.poll(() => picture.evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(200);
  const placeholder = await picture.getAttribute('src');
  await page.getByTestId('tab-data').click();
  const logo = '/data[1]/logo[1]';
  await page.locator(`[data-xpath="${logo}"] > .tree-row`).hover();
  await page.getByRole('button', { name: `Bind to ${logo}: the control at the caret, or a new picture control` }).click();
  await expect.poll(() => page.locator('.docx-editor.body [data-kind="Picture"] img').first().getAttribute('src')).not.toBe(placeholder);
});

// ED-005 section 12.89: at open, bound controls whose content is older than their data are offered a refresh.
test('the stale-content banner: Refresh fills the picture control; Keep leaves the file as it was', async ({ page }) => {
  await openFixture(page, 'pictures/38a-picture-template.docx');
  await expect(page.getByTestId('stale-offer')).toBeVisible();
  await expect(page.getByTestId('stale-offer')).toContainText('One bound control shows content older than its data.');
  // The control shows its red placeholder picture; Refresh puts the node's image (another part) in its place.
  const picture = page.locator('.docx-editor.body [data-kind="Picture"] img').first();
  await expect(picture).toBeAttached();
  const placeholder = await picture.getAttribute('src');
  await page.getByTestId('stale-refresh').click();
  await expect(page.getByTestId('stale-offer')).toHaveCount(0);
  await expect(page.getByTestId('status')).toHaveText('Refreshed from the data');
  await expect.poll(() => page.locator('.docx-editor.body [data-kind="Picture"] img').first().getAttribute('src')).not.toBe(placeholder);
  // Keep: nothing changes, and the file comes back as it went in.
  await openFixture(page, 'pictures/38a-picture-template.docx');
  await expect(page.getByTestId('stale-offer')).toBeVisible();
  await page.getByTestId('stale-keep').click();
  await expect(page.getByTestId('stale-offer')).toHaveCount(0);
  const saved = await download(page);
  expect(comparePackages(await fixture('pictures/38a-picture-template.docx'), saved)).toEqual([]);
});

/** The caret inside the control whose title (`w:alias`) is `label`. */
async function caretInControl(page: import('@playwright/test').Page, label: string): Promise<void> {
  await page.evaluate((label) => {
    const v = (window as unknown as { docx4jEditor: { view: { posAtDOM(n: Node, o: number): number }; caretAt(a: number): void } }).docx4jEditor;
    const el = document.querySelector(`.docx-editor.body [data-label="${label}"]`);
    if (!el) throw new Error(`no control ${label}`);
    const text = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
    v.caretAt(v.view.posAtDOM(text ?? el, 0) + (text ? 1 : 0));
  }, label);
}

// ED-005 section 12.117: the Control tab's picture facts, and a handler's width set there, the picture placed again at once.
test('the Control tab shows a picture binding\'s facts; a handler\'s width set to auto is written to the tag and the picture placed again', async ({ page }) => {
  await openFixture(page, 'pictures/38b-handler-bound.docx');
  await page.getByTestId('perspective-template').click();
  await caretInControl(page, 'width=4500');
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  const facts = page.getByTestId('control-picture');
  await expect(facts).toBeVisible();
  await expect(page.getByTestId('cc-picture-kind')).toHaveText("OpenDoPE's picture handler (od:Handler=picture), bound by its od:xpath entry");
  await expect(page.getByTestId('cc-picture-data')).toHaveText('1 KB of image data');
  await expect(page.getByTestId('cc-picture-image')).toContainText('PNG, 800 × 200 pixels');
  await expect.poll(() => page.getByTestId('cc-picture-image').locator('img').evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(800);
  await expect(page.getByTestId('cc-picture-width-mode')).toHaveValue('points');
  await expect(page.getByTestId('cc-picture-width')).toHaveValue('225');
  // The picture the file holds, then auto: the tag says so, and the control holds a new picture (another part, so another URL).
  const picture = page.locator('.docx-editor.body [data-label="width=4500"] img').first();
  const before = await picture.getAttribute('src');
  await page.getByTestId('cc-picture-width-mode').selectOption('auto');
  await page.getByTestId('cc-picture-width-apply').click();
  await expect(page.getByTestId('cc-tag')).toHaveValue('od:xpath=x2&od:Handler=picture&width=auto');
  await expect.poll(() => page.locator('.docx-editor.body img').evaluateAll((els, b) => els.filter((e) => (e as HTMLImageElement).src !== b).length, before)).toBeGreaterThan(0);
  // A width of no points is refused, saying so.
  await page.getByTestId('cc-picture-width-mode').selectOption('points');
  await page.getByTestId('cc-picture-width').fill('0');
  await page.getByTestId('cc-picture-width-apply').click();
  await expect(page.getByTestId('status')).toHaveText('Enter a width in points greater than 0');
  // Word's picture control: its kind, no width field.
  await openFixture(page, 'pictures/38a-picture-template.docx');
  await page.getByTestId('perspective-template').click();
  await caretInControl(page, 'Logo');
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('cc-picture-kind')).toHaveText("Word's picture control, mapped by w:dataBinding");
  await expect(page.getByTestId('cc-picture-image')).toContainText('PNG, 300 × 150 pixels');
  await expect(page.getByTestId('cc-picture-width-mode')).toHaveCount(0);
});

// ED-005 section 12.117: the Data tab shows a node holding an image as a thumbnail; its bind makes Word's picture control there,
// showing the image; an SVG node, which the editor cannot place as a new picture, is refused, saying why.
test('the Data tab: an image node\'s thumbnail; its bind makes a picture control showing the image; an SVG node\'s bind refused', async ({ page }) => {
  await openFixture(page, 'pictures/38b-handler-bound.docx');
  await page.getByTestId('perspective-template').click();
  await page.getByTestId('tab-data').click();
  const logo = '/data[1]/logo[1]';
  const thumbnail = page.locator(`[data-xpath="${logo}"] > .tree-row [data-testid="node-thumbnail"]`);
  await expect(thumbnail).toHaveAttribute('title', 'PNG, 300 × 150 pixels');
  await expect.poll(() => thumbnail.evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(300);
  await expect(page.locator('[data-xpath="/data[1]/svg[1]"] > .tree-row [data-testid="node-thumbnail"]')).toHaveAttribute('title', 'SVG, which states no size in pixels');
  // The caret after the last paragraph's text, outside every control.
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { descendants(f: (n: { type: { name: string }; textContent: string; content: { size: number } }, pos: number) => boolean): void } }; caretAt(a: number): void } }).docx4jEditor;
    let end = -1;
    v.state.doc.descendants((n, pos) => { if (n.type.name === 'paragraph' && n.textContent === 'After the controls.') end = pos + 1 + n.content.size; return true; });
    v.caretAt(end);
  });
  const pictures = page.locator('.docx-editor.body [data-kind="Picture"]');
  await expect(pictures).toHaveCount(0);
  await page.locator('[data-xpath="/data[1]/svg[1]"] > .tree-row').hover();
  await page.getByRole('button', { name: 'Bind to /data[1]/svg[1]: the control at the caret, or a new picture control' }).click();
  await expect(page.getByTestId('status')).toHaveText('The editor places PNG, JPEG, GIF and BMP pictures: this node holds another kind of image');
  await expect(pictures).toHaveCount(0);
  await page.locator(`[data-xpath="${logo}"] > .tree-row`).hover();
  await page.getByRole('button', { name: `Bind to ${logo}: the control at the caret, or a new picture control` }).click();
  await expect(page.getByTestId('status')).toHaveText(`Inserted a picture control bound to ${logo}`);
  await expect(pictures).toHaveCount(1);
  await expect.poll(() => pictures.locator('img').first().evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(300);
});

// Section 12.125 (Jason, 2026-11-12, tmp/testcase_sdt_pic.docx): a node holding a data URL, not bare base64, as a person's or a
// program's data often does. The Data tab's bind had made a plain text control showing the URL, and a picture control bound by
// a drag had kept its placeholder.
const JASONS_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAgMAAAAP2OW3AAAADFBMVEUDAP//AAAA/wb//AAD4Tw1AAAACXBIWXMAAAsTAAALEwEAmpwYAAAADElEQVQI12NwYNgAAAF0APHJnpmVAAAAAElFTkSuQmCC';
const PIC = '/invoice[1]/pic[1]';

/** HelloWordOnline with a data part whose `pic` holds Jason's PNG as a data URL, open in the Template perspective's Data tab. */
async function withDataUrl(page: import('@playwright/test').Page): Promise<void> {
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  pkg.customXmlParts.add(`<invoice><pic>${JASONS_PNG}</pic><name>Joe Bloggs</name></invoice>`);
  await openBytes(page, 'testcase_sdt_pic.docx', await pkg.save());
  await page.getByTestId('perspective-template').click();
  await page.getByTestId('tab-data').click();
  await expect(page.locator(`[data-xpath="${PIC}"] > .tree-row [data-testid="node-thumbnail"]`)).toHaveAttribute('title', 'PNG, 2 × 2 pixels');
}

test('a node holding a data URL: the Data tab\'s bind makes a picture control showing the image', async ({ page }) => {
  await withDataUrl(page);
  await caretInBlock(page, 0, 3);
  await page.locator(`[data-xpath="${PIC}"] > .tree-row`).hover();
  await page.getByRole('button', { name: `Bind to ${PIC}: the control at the caret, or a new picture control` }).click();
  await expect(page.getByTestId('status')).toHaveText(`Inserted a picture control bound to ${PIC}`);
  const pictures = page.locator('.docx-editor.body [data-kind="Picture"]');
  await expect(pictures).toHaveCount(1);
  await expect.poll(() => pictures.locator('img').first().evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(2);
  await expect(pictures).not.toContainText('data:image');
});

test('a node holding a data URL dragged onto a picture control: the image in place of the placeholder', async ({ page }) => {
  await withDataUrl(page);
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { content: { size: number } } }; caretAt(a: number): void; run(name: string, args?: unknown): boolean } }).docx4jEditor;
    v.caretAt(v.state.doc.content.size - 1);
    v.run('insertParagraph');
    v.run('contentControl.insert', { kind: 'Picture' });
  });
  const control = page.locator('.docx-editor.body [data-kind="Picture"]');
  await expect(control).toHaveCount(1);
  // Word's form (section 12.126): the placeholder picture, 200 pixels wide, where the node's (2 pixels) goes.
  await expect.poll(() => control.locator('img').first().evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(200);
  await page.locator(`[data-xpath="${PIC}"] .node-name`).first().dragTo(control);
  await expect(page.getByTestId('status')).toContainText(`Bound to ${PIC}`);
  await expect.poll(() => control.locator('img').first().evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(2);
});

// ED-005 section 12.127: a picture on the clipboard goes in at the caret as Insert picture puts one.
/** A paste of one picture file, with `html` beside it; a WebP is drawn and encoded in the page, as a browser holds one. */
async function pastePicture(page: import('@playwright/test').Page, picture: { png: Buffer } | { webp: { width: number; height: number } }, html = ''): Promise<void> {
  await page.evaluate(async ([pic, markup]) => {
    let file: File;
    if ('png' in pic) {
      file = new File([Uint8Array.from(atob(pic.png), (c) => c.charCodeAt(0))], 'image.png', { type: 'image/png' });
    } else {
      const canvas = document.createElement('canvas');
      canvas.width = pic.webp.width;
      canvas.height = pic.webp.height;
      const g = canvas.getContext('2d')!;
      g.fillStyle = '#00f';
      g.fillRect(0, 0, canvas.width, canvas.height);
      const blob = await new Promise<Blob>((resolve) => canvas.toBlob((b) => resolve(b!), 'image/webp'));
      if (blob.type !== 'image/webp') throw new Error(`the browser encoded ${blob.type}`);
      file = new File([blob], 'image.webp', { type: 'image/webp' });
    }
    const dt = new DataTransfer();
    dt.items.add(file);
    if (markup !== '') dt.setData('text/html', markup);
    document.querySelector('.docx-editor .ProseMirror[aria-label="Document"]')!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, ['png' in picture ? { png: picture.png.toString('base64') } : picture, html] as const);
}

test('a picture pasted from the clipboard (a screenshot: the file alone) lands at the caret, drawn from its part; the download holds it', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4, 5);
  await pastePicture(page, { png: png(120, 60) });
  const image = page.locator('.docx-editor .ProseMirror[aria-label="Document"] > p').nth(4).locator('img.image');
  await expect(image).toHaveCount(1);
  await expect.poll(() => image.evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(120);
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  expect(await pkg.getMainDocumentPart().getXml()).toMatch(/<wp:inline[\s\S]*<wp:extent cx="1143000" cy="571500"\/>[\s\S]*<a:blip r:embed="rId\d+"/);
  expect(pkg.getPart('/word/media/image1.png')).toBeTruthy();
});

test("a browser's Copy image of a WebP (the file and an <img>'s markup): the picture goes in as PNG, which Word reads", async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4, 5);
  await pastePicture(page, { webp: { width: 40, height: 20 } }, "<meta charset='utf-8'><img src=\"https://example.org/blue.webp\" alt=\"blue\"/>");
  const image = page.locator('.docx-editor .ProseMirror[aria-label="Document"] > p').nth(4).locator('img.image');
  await expect(image).toHaveCount(1);
  await expect.poll(() => image.evaluate((el) => (el as HTMLImageElement).naturalWidth)).toBe(40);
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const media = pkg.getPart('/word/media/image1.png');
  expect(media).toBeTruthy();
  expect(pkg.getPart('/word/media/image1.webp')).toBeFalsy();
  expect(await pkg.getMainDocumentPart().getXml()).toMatch(/<wp:extent cx="381000" cy="190500"\/>/);
});
