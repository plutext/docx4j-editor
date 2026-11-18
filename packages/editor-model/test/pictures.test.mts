// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 proposal 45 part 2 (section 12.76): a picture made in the editor.
import { describe, test, expect } from 'vitest';
import { deflateSync } from 'node:zlib';
import { unzipSync } from 'fflate';
import { EditorState, TextSelection } from 'prosemirror-state';
import { createHash } from 'node:crypto';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, pictureFactsOf, base64Bytes, imageBase64Of, looksLikeImageBase64, withPictureWidth, dataBindRefusal, bindsPicture, pictureControlsIn, sessionPlugin, stalePictureControls, readControl, fittedInBox } from '../src/index.mjs';
import { fixture, mainXml } from './helpers.mjs';

/** The node in Jason's tmp/testcase_sdt_pic.docx (2026-11-12): a 2 by 2 PNG as a data URL. */
const JASONS_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAgMAAAAP2OW3AAAADFBMVEUDAP//AAAA/wb//AAD4Tw1AAAACXBIWXMAAAsTAAALEwEAmpwYAAAADElEQVQI12NwYNgAAAF0APHJnpmVAAAAAElFTkSuQmCC';

/** A PNG of `w` by `h` red pixels, made here (no fixture needed). */
export function pngBase64(w: number, h: number): string {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf: Buffer): number => { let c = 0xffffffff; for (const b of buf) c = crcTable[(c ^ b) & 0xff]! ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type: string, data: Buffer): Buffer => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type, 'ascii'), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  const raw = Buffer.alloc((1 + w * 3) * h);
  for (let y = 0; y < h; y++) { raw[y * (1 + w * 3)] = 0; for (let x = 0; x < w; x++) { const o = y * (1 + w * 3) + 1 + x * 3; raw[o] = 255; raw[o + 1] = 0; raw[o + 2] = 0; } }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]).toString('base64');
}

describe('range.insertInlinePicture', () => {
  test('an image node at the caret, its part and relationship made in prepare; saved, the main part holds the inline drawing and the package the media part; the natural size, or the width asked', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    const r = await runCommandAsync(session, 'range.insertInlinePicture', { base64: pngBase64(96, 48) }, { from: 3 });
    expect(r.ok).toBe(true);
    const first = session.doc.child(0);
    let image: { attrs: Record<string, unknown> } | undefined;
    first.forEach((n) => { if (n.type.name === 'image') image = n; });
    expect(image).toBeDefined();
    expect(image!.attrs['sourceId']).toBeTruthy();
    expect(image!.attrs['width']).toBe(72);   // 96 px at 96 dpi = one inch
    expect(image!.attrs['height']).toBe(36);
    // A second one, 100 pt wide: the height follows the ratio.
    const r2 = await runCommandAsync(session, 'range.insertInlinePicture', { base64: pngBase64(200, 100), width: 100, name: 'Logo' }, { from: 6 });
    expect(r2.ok).toBe(true);
    const saved = (await session.save()).bytes;
    const xml = await mainXml(saved);
    expect((xml.match(/<wp:inline/g) ?? []).length).toBe(2);
    expect(xml).toMatch(/wp:docPr [^>]*name="Logo"/);
    expect(xml).toMatch(/<wp:extent cx="1270000" cy="635000"\/>/);
    const again = await EditorDocument.open(saved);
    expect(again.pkg.getPart('/word/media/image1.png')).toBeTruthy();
    expect(again.pkg.getPart('/word/media/image2.png')).toBeTruthy();
    expect(again.pkg.getPart('/word/media/image3.png')).toBeFalsy();
    let count = 0;
    again.doc.child(0).forEach((n) => { if (n.type.name === 'image') count++; });
    expect(count).toBe(2);
  });

  test('refused in a note, and with no base64', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    const r = await runCommandAsync(session, 'range.insertInlinePicture', { base64: '' }, { from: 3 });
    expect(r.ok).toBe(false);
  });
});

// ED-005 section 12.117 (proposal 45 part 3's tabs): what a node's image data is, a handler's width written to its tag, and the
// Data tab's bind of an image node making Word's picture control.
describe('picture facts, the width parameter, and a picture control bound from the data', () => {
  test('pictureFactsOf reads a PNG and an SVG from their headers; text, markup and an unread format are no image', async () => {
    const png = pngBase64(30, 20);
    expect(pictureFactsOf(png)).toEqual({ format: 'PNG', contentType: 'image/png', widthPx: 30, heightPx: 20, bytes: base64Bytes(png).length, placeable: true });
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>').toString('base64');
    expect(pictureFactsOf(svg)).toMatchObject({ format: 'SVG', placeable: false });
    expect(pictureFactsOf('Joe Bloggs')).toBeUndefined();
    expect(pictureFactsOf('<p>markup</p>')).toBeUndefined();
    expect(pictureFactsOf('SUkqAAgAAAA=')).toBeUndefined();   // a TIFF's head: an image, not one the engine reads
  });

  test('a data URL is read as its base64, as the engine reads a node (section 12.125): the facts, the bytes, and the Data tab\'s bind making a picture control', async () => {
    const png = pngBase64(30, 20);
    const url = `data:image/png;base64,${png}`;
    expect(imageBase64Of(` ${url} `)).toBe(png);
    expect(looksLikeImageBase64(url)).toBe(true);
    expect(pictureFactsOf(url)).toEqual(pictureFactsOf(png));
    expect(base64Bytes(url)).toEqual(base64Bytes(png));
    expect(pictureFactsOf('data:text/plain,Joe Bloggs')).toBeUndefined();
    // Jason's node (tmp/testcase_sdt_pic.docx), a 2 by 2 PNG as a data URL: the bind had made a plain text control showing it.
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    session.pkg.customXmlParts.add(`<invoice><pic>${JASONS_PNG}</pic></invoice>`);
    await session.template.load();
    let end = 0;
    session.doc.descendants((n, pos) => { if (end === 0 && n.type.name === 'paragraph' && n.textContent.length > 10) end = pos + 1 + n.content.size; return end === 0; });
    const pic = { xpath: '/invoice[1]/pic[1]' };
    expect(pictureFactsOf(JASONS_PNG)).toMatchObject({ format: 'PNG', widthPx: 2, heightPx: 2, placeable: true });
    expect(bindsPicture(stateAt(session, end), pic)).toBe(true);
    expect(runCommand(session, 'contentControl.insertBound', pic, { from: end }).ok).toBe(true);
    expect(pictureControlsIn(session.doc, session.template).map((c) => c.picture.shape)).toEqual(['picture']);
  });

  test('a picture made from a data URL holds the image\'s bytes, never its prefix read as base64', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    const png = pngBase64(30, 20);
    const made = session.addInlinePicture(`data:image/png;base64,${png}`);
    expect(made).toMatchObject({ width: 22.5, height: 15 });
    const zip = unzipSync((await session.save({ selfCheck: false })).bytes);
    const media = Object.keys(zip).filter((name) => name.startsWith('word/media/'));
    expect(media).toHaveLength(1);
    expect(Buffer.compare(Buffer.from(zip[media[0]!]!), Buffer.from(base64Bytes(png)))).toBe(0);
  });

  // Section 12.125's Word check (fixtures/pictures-dataurl/, Word 16.0.20326.20318 through the bridge, 2026-11-12).
  test("Word 365's saves: a control holding a picture refilled from bare base64, flag or none, and not from a data URL; one with no picture never filled; the node never written; the editor would refill pdu3, where Word keeps the control's picture", async () => {
    const orange = (b64: string): Uint8Array => base64Bytes(b64);
    const facts: Record<string, { media: 'orange' | 'blue' | 'none'; node: string; stale: number; extent?: string }> = {};
    for (const name of ['pdu1-base64-empty', 'pdu2-dataurl-empty', 'pdu3-dataurl-filled', 'pdu4-base64-filled', 'pdu5-base64-text-noflag', 'pdu6-base64-placeholder-flag', 'pdu7-base64-empty-content']) {
      const input = unzipSync(await fixture(`pictures-dataurl/${name}.docx`));
      const saved = unzipSync(await fixture(`pictures-dataurl/${name}-word365.docx`));
      const node = (bytes: Uint8Array | undefined): string => /<pic>([^<]*)<\/pic>/.exec(Buffer.from(bytes!).toString('utf8'))![1]!;
      const value = node(input['customXml/item1.xml']);
      const media = saved['word/media/image1.png'];
      const colour = media === undefined ? 'none' : Buffer.compare(Buffer.from(media), Buffer.from(orange(value))) === 0 ? 'orange' : Buffer.compare(Buffer.from(media), Buffer.from(input['word/media/image1.png']!)) === 0 ? 'blue' : 'other';
      const session = await EditorDocument.open(await fixture(`pictures-dataurl/${name}-word365.docx`));
      await session.template.load();
      const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(Buffer.from(saved['word/document.xml']!).toString('utf8'));
      facts[name] = { media: colour as 'orange' | 'blue' | 'none', node: node(saved['customXml/item1.xml']) === value ? 'kept' : 'written', stale: await stalePictureControls(session, session.doc), ...(extent ? { extent: `${extent[1]}x${extent[2]}` } : {}) };
    }
    expect(facts).toEqual({
      'pdu1-base64-empty': { media: 'none', node: 'kept', stale: 1 },
      'pdu2-dataurl-empty': { media: 'none', node: 'kept', stale: 1 },
      // The divergence (docs/developer/word-divergences.md): Word keeps the control's own picture over a data URL, the editor would refill it.
      'pdu3-dataurl-filled': { media: 'blue', node: 'kept', stale: 1, extent: '1143000x571500' },
      'pdu4-base64-filled': { media: 'orange', node: 'kept', stale: 0, extent: '1143000x571500' },
      // Round 2: pdu5 opened only through a repair (text without the flag), its text kept; pdu6, Word's form, filled with the
      // placeholder's width kept and the height fitted to the picture's shape; pdu7, nothing to swap, left empty.
      'pdu5-base64-text-noflag': { media: 'none', node: 'kept', stale: 1 },
      'pdu6-base64-placeholder-flag': { media: 'orange', node: 'kept', stale: 0, extent: '1270000x635000' },
      'pdu7-base64-empty-content': { media: 'none', node: 'kept', stale: 1 },
    });
  });

  test('withPictureWidth writes points as twentieths, auto, or removes the parameter; the binding reads it back', () => {
    const tag = 'od:xpath=x2&od:Handler=picture&width=4500';
    expect(withPictureWidth(tag, 100)).toBe('od:xpath=x2&od:Handler=picture&width=2000');
    expect(withPictureWidth(tag, 'auto')).toBe('od:xpath=x2&od:Handler=picture&width=auto');
    expect(withPictureWidth(tag, undefined)).toBe('od:xpath=x2&od:Handler=picture');
    expect(withPictureWidth('od:xpath=x1&od:Handler=picture', 12.5)).toBe('od:xpath=x1&od:Handler=picture&width=250');
  });

  test('check 38b: the Data tab\'s bind of an image node makes Word\'s picture control mapped to it; an SVG node is refused', async () => {
    const session = await EditorDocument.open(await fixture('pictures/38b-handler-bound.docx'));
    await session.template.load();
    const part = session.template.dataPart.part!;
    // The caret at the end of the last paragraph's text, outside every control.
    let end = 0;
    session.doc.descendants((n, pos) => { if (n.type.name === 'paragraph' && n.textContent === 'After the controls.') end = pos + 1 + n.content.size; return true; });
    expect(end).toBeGreaterThan(0);
    const logo = { xpath: '/data[1]/logo[1]', storeItemID: part.id };
    const svg = { xpath: '/data[1]/svg[1]', storeItemID: part.id };
    expect(dataBindRefusal(stateAt(session, end), svg)).toBe('imageFormat');
    expect(runCommand(session, 'contentControl.insertBound', svg, { from: end }).ok).toBe(false);
    expect(bindsPicture(stateAt(session, end), logo)).toBe(true);
    expect(dataBindRefusal(stateAt(session, end), logo)).toBeUndefined();
    expect(runCommand(session, 'contentControl.insertBound', logo, { from: end }).ok).toBe(true);
    const made = pictureControlsIn(session.doc, session.template).filter((c) => c.picture.shape === 'picture');
    expect(made).toHaveLength(1);
    expect(made[0]!.node.type.name).toBe('sdtInline');
    expect(made[0]!.picture.binding.xpath).toBe('/data[1]/logo[1]');
  });
});

// Section 12.126: the editor's new picture control in Word's form, which Word fills from a mapped node (round 2's pdu6), where it
// never fills the placeholder text the editor made before (pdu1, pdu5).
describe("a new picture control in Word's form (section 12.126)", () => {
  /** The first run-level control in the document. */
  const firstControl = (doc: PMNode): PMNode => { let c: PMNode | undefined; doc.descendants((n) => { if (!c && n.type.name === 'sdtInline') c = n; return !c; }); return c!; };
  const sha1 = (bytes: Uint8Array): string => createHash('sha1').update(bytes).digest('hex');

  test("at a caret: the flag and docx4j's placeholder picture, 100 points square; saved as Word's form, with no text and no w:placeholder, the image in a part of its own", async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    const before = unzipSync((await session.save({ selfCheck: false })).bytes);
    expect(runCommand(session, 'contentControl.insert', { kind: 'Picture' }, { from: 3 }).ok).toBe(true);
    const control = firstControl(session.doc);
    expect(readControl(control.attrs['sdtPr'])).toMatchObject({ kind: 'Picture', showingPlaceholder: true });
    expect(control.content.content.map((n) => [n.type.name, n.attrs['width'], n.attrs['height']])).toEqual([['image', 100, 100]]);
    const saved = (await session.save()).bytes;
    const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(await mainXml(saved))![0];
    expect(sdt).toMatch(/<w:picture\/>/);
    expect(sdt).toMatch(/<w:showingPlcHdr\/>/);
    expect(sdt).not.toMatch(/<w:placeholder>/);
    expect(sdt).not.toMatch(/<w:t[ >]/);
    expect(sdt).toMatch(/<wp:extent cx="1270000" cy="1270000"\/>/);
    const zip = unzipSync(saved);
    const added = Object.keys(zip).filter((name) => name.startsWith('word/media/') && !(name in before));
    expect(added.map((name) => sha1(zip[name]!))).toEqual(['e9e4ee21363e5c8d0f3200fd920b03d5e66e4ab7']);
  });

  test('around a selection, as before: the selection wrapped, with no flag (as the engine\'s insertContentControl wraps a span)', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    expect(runCommand(session, 'contentControl.insert', { kind: 'Picture' }, { from: 3, to: 6 }).ok).toBe(true);
    const control = firstControl(session.doc);
    expect(readControl(control.attrs['sdtPr']).showingPlaceholder).toBeFalsy();
    expect(control.textContent).toHaveLength(3);
  });

  test("fittedInBox: the node's picture inside the held picture's box, keeping its ratio, as Word 365 shows it (pdu6; core-ts's 2:1 box over a 1:2 picture)", () => {
    const emu = (size: { width: number; height: number } | undefined): string => (size ? `${Math.round(size.width * 12700)}x${Math.round(size.height * 12700)}` : 'none');
    expect(emu(fittedInBox(pngBase64(120, 60), { width: 100, height: 100 }))).toBe('1270000x635000');
    expect(emu(fittedInBox(pngBase64(60, 120), { width: 100, height: 50 }))).toBe('317500x635000');
    expect(emu(fittedInBox(`data:image/png;base64,${pngBase64(120, 60)}`, { width: 100, height: 100 }))).toBe('1270000x635000');
    expect(emu(fittedInBox('hello', { width: 100, height: 100 }))).toBe('none');
    expect(emu(fittedInBox(pngBase64(120, 60), { width: 0, height: 100 }))).toBe('none');
  });

  test("mapped to a node holding a picture, it is stale until refreshed: Word's reading (pdu6) and the engine's bind agree", async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    session.pkg.customXmlParts.add(`<data><pic>${pngBase64(120, 60)}</pic></data>`);
    await session.template.load();
    expect(runCommand(session, 'contentControl.insert', { kind: 'Picture' }, { from: 3 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/pic' }, { from: 4 }).ok).toBe(true);
    expect(await stalePictureControls(session, session.doc)).toBe(1);
    const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(await mainXml((await session.save()).bytes))![0];
    expect(sdt).toMatch(/<w:dataBinding [^>]*w:xpath="\/data\/pic"/);
    expect(sdt).toMatch(/<w:showingPlcHdr\/>/);
  });
});

/** The session's document as a state with the caret at `pos`, carrying the session as the view's does. */
function stateAt(session: EditorDocument, pos: number): EditorState {
  const state = EditorState.create({ doc: session.doc, plugins: [sessionPlugin(session)] });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
}
