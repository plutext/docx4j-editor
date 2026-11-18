// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.76: picture-bound controls shown. A Word picture control bound to a node holding
// base64 takes the image; a rich text control tagged od:Handler=picture with a width takes it scaled;
// one without a width is left and noted (the processor's, REQ-061); a node that is not an image is noted.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { deflateSync } from 'node:zlib';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, pictureControlsIn, stalePictureControls } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { refreshPictureControls } from '../src/view/picturebind.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

function pngBase64(w: number, h: number): string {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf: Buffer): number => { let c = 0xffffffff; for (const b of buf) c = crcTable[(c ^ b) & 0xff]! ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type: string, data: Buffer): Buffer => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type, 'ascii'), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((1 + w * 3) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = y * (1 + w * 3) + 1 + x * 3; raw[o] = 255; }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]).toString('base64');
}

async function open(values: Record<string, string>, name = 'HelloWordOnline.docx'): Promise<EditorDocument> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  session.pkg.customXmlParts.add(`<data>${Object.entries(values).map(([k, v]) => `<${k}>${v}</${k}>`).join('')}</data>`);
  await session.template.load();
  return session;
}
const imagesIn = (node: PMNode): PMNode[] => { const out: PMNode[] = []; node.descendants((n) => { if (n.type.name === 'image') out.push(n); return true; }); return out; };

describe('picture-bound controls (ED-005 section 12.76)', () => {
  test("Word's picture control bound to a base64 node takes the image fitted inside the placeholder's square, as Word 365 shows it (section 12.126); a second refresh over the same value changes nothing", async () => {
    const session = await open({ logo: pngBase64(96, 48) });
    expect(runCommand(session, 'contentControl.insert', { kind: 'Picture' }, { from: 3 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/logo' }, { from: 4 }).ok).toBe(true);
    expect(pictureControlsIn(session.doc, session.template)).toHaveLength(1);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const first = await refreshPictureControls(editor.view, session);
    expect(first.changed).toBe(1);
    const images = imagesIn(editor.state.doc.child(0));
    expect(images).toHaveLength(1);
    expect(images[0]!.attrs['width']).toBe(100);
    expect(images[0]!.attrs['height']).toBe(50);
    expect((await refreshPictureControls(editor.view, session)).changed).toBe(0);
    editor.destroy();
  });

  test("a node holding a data URL (Jason's tmp/testcase_sdt_pic.docx; section 12.125): Word's picture control takes the image, and is then not stale", async () => {
    const session = await open({ pic: `data:image/png;base64,${pngBase64(96, 48)}` });
    expect(runCommand(session, 'contentControl.insert', { kind: 'Picture' }, { from: 3 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/pic' }, { from: 4 }).ok).toBe(true);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    expect(await stalePictureControls(session, editor.state.doc)).toBe(1);
    const r = await refreshPictureControls(editor.view, session);
    expect(r).toEqual({ changed: 1, notes: [] });
    const images = imagesIn(editor.state.doc.child(0));
    expect(images.map((i) => [i.attrs['width'], i.attrs['height']])).toEqual([[100, 50]]);
    expect(await stalePictureControls(session, editor.state.doc)).toBe(0);
    editor.destroy();
  });

  test("a picture control holding text (made around a selection) takes the image at its natural size, as the engine fills one where Word leaves it (core-ts CR-005, E2)", async () => {
    const session = await open({ logo: pngBase64(96, 48) });
    expect(runCommand(session, 'contentControl.insert', { kind: 'Picture' }, { from: 3, to: 6 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/logo' }, { from: 4 }).ok).toBe(true);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    expect((await refreshPictureControls(editor.view, session)).changed).toBe(1);
    expect(imagesIn(editor.state.doc.child(0)).map((i) => [i.attrs['width'], i.attrs['height']])).toEqual([[72, 36]]);
    editor.destroy();
  });

  test('od:Handler=picture with width=1440 scales to 72 points; without a width it is left and noted; a node that is no image is noted', async () => {
    const session = await open({ logo: pngBase64(200, 100), text: 'hello' }, 'tracked-changes.docx');
    const para = session.doc.child(0);
    expect(runCommand(session, 'contentControl.insert', {}, { from: 1, to: 1 + para.content.size }).ok).toBe(true);
    const entry = session.template.upsertEntry({ xpath: '/data/logo' });
    expect(runCommand(session, 'contentControl.properties', { tag: `od:xpath=${entry.id}&od:Handler=picture&width=1440` }, { from: 2 }).ok).toBe(true);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const r = await refreshPictureControls(editor.view, session);
    expect(r.changed).toBe(1);
    const block = editor.state.doc.child(0);
    expect(block.type.name).toBe('sdtBlock');
    const images = imagesIn(block);
    expect(images).toHaveLength(1);
    expect(images[0]!.attrs['width']).toBe(72);
    expect(images[0]!.attrs['height']).toBe(36);
    // Fresh controls: one without a width is left, with a note; one over the text node is noted.
    const wrap = (index: number, tag: string): void => {
      let start = 0;
      for (let i = 0; i < index; i++) start += editor.state.doc.child(i).nodeSize;
      const block = editor.state.doc.child(index);
      expect(block.type.name).toBe('paragraph');
      expect(runCommand(session, 'contentControl.insert', {}, { from: start + 1, to: start + 1 + block.content.size }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.properties', { tag }, { from: start + 2 }).ok).toBe(true);
    };
    const textEntry = session.template.upsertEntry({ xpath: '/data/text' });
    wrap(1, `od:xpath=${entry.id}&od:Handler=picture`);
    wrap(2, `od:xpath=${textEntry.id}&od:Handler=picture&width=auto`);
    const again = await refreshPictureControls(editor.view, session);
    expect(again.changed).toBe(0);
    expect(again.notes.some((n) => /floating picture/.test(n))).toBe(true);
    expect(again.notes.some((n) => /not hold an image/.test(n))).toBe(true);
    editor.destroy();
  });
});
