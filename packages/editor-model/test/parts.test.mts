// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 4 (E4.b): the parts made editable. Which parts are edited and how (section
// 4.2); the check before an Apply, which refuses a text that does not parse, a value of the wrong
// type, and an element or attribute the object model would drop, each with where it is (4.3); an
// Apply's effects on what the session reads (4.4); the Apply in a room, a styles or numbering part
// replicated by element so a peer's concurrent change to another element survives (4.5).
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, type Part, type XmlPart } from '@docx4j/core-ts';
import { EditorDocument, partEditing, checkPartXml, applyPartXml, liveLabels, readControl } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';
import { room, RELEASE_ORDERS } from './room-harness.mjs';

const part = (session: EditorDocument, name: string): Part => {
  const p = session.pkg.parts.get(name);
  if (!p) throw new Error(`no part ${name}`);
  return p;
};

const xmlOf = async (p: Part): Promise<string> => (p as unknown as { getXml(): Promise<string> }).getXml();

/** A style's colour set: its own `w:color` replaced, or one added at the start of its run properties. */
function withColor(xml: string, id: string, color: string): string {
  const style = new RegExp(`<w:style [^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(xml)![0];
  const own = /<w:color [^>]*\/>/;
  const changed = own.test(style) ? style.replace(own, `<w:color w:val="${color}"/>`) : style.replace('<w:rPr>', `<w:rPr><w:color w:val="${color}"/>`);
  return xml.replace(style, changed);
}

describe('which parts, and how (section 4.2)', () => {
  test('package parts typed or as bytes; the projected ones, the comments, relationships and binaries read only', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    const modes = new Map<string, string>();
    for (const p of session.pkg.parts) {
      const e = partEditing(p);
      modes.set(p.partName.toString(), e.mode === 'none' ? `none:${e.reason}` : e.mode);
    }
    expect(modes.get('/word/document.xml')).toBe('none:projected');
    expect(modes.get('/word/styles.xml')).toBe('typed');
    expect(modes.get('/word/settings.xml')).toBe('typed');
    expect(modes.get('/word/theme/theme1.xml')).toBe('typed');
    expect(modes.get('/word/fontTable.xml')).toBe('typed');
    expect(modes.get('/docProps/core.xml')).toBe('typed');
    expect(modes.get('/docProps/app.xml')).toBe('typed');
    expect(modes.get('/word/comments.xml')).toBe('none:comments');
    expect([...modes].filter(([n]) => n.endsWith('.rels')).every(([, m]) => m === 'none:regenerated')).toBe(true);
    expect([...modes].filter(([n]) => /^\/customXml\/item\d+\.xml$/.test(n)).every(([, m]) => m === 'bytes')).toBe(true);
    expect([...modes].filter(([n]) => /^\/customXml\/itemProps\d+\.xml$/.test(n)).every(([, m]) => m === 'typed')).toBe(true);
    const images = await EditorDocument.open(await fixture('images.docx'));
    const media = [...images.pkg.parts].filter((p) => p.partName.toString().startsWith('/word/media/'));
    expect(media.length).toBeGreaterThan(0);
    expect(media.every((p) => partEditing(p).mode === 'none')).toBe(true);
    const notes = await EditorDocument.open(await fixture('footnote.docx'));
    expect(partEditing(part(notes, '/word/footnotes.xml'))).toEqual({ mode: 'none', reason: 'projected' });
  });
});

describe('the check before an Apply (section 4.3)', () => {
  test('an unchanged part passes; an unknown element, an unknown attribute and an element in an unknown namespace are refused with their path and line', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const styles = part(session, '/word/styles.xml');
    const xml = await xmlOf(styles);
    expect(await checkPartXml(styles, xml)).toEqual([]);
    // Pretty-printed as the pane shows it, so lines mean something.
    const pretty = xml.replace(/></g, '>\n<');
    expect(await checkPartXml(styles, pretty)).toEqual([]);
    const bogus = pretty.replace('<w:name w:val="Normal"/>', '<w:name w:val="Normal" w:foo="1"/>\n<w:bogus/>\n<x:thing xmlns:x="urn:example:x"/>');
    const problems = await checkPartXml(styles, bogus);
    const line = bogus.split('\n').findIndex((l) => l === '<w:bogus/>') + 1;
    expect(problems).toContainEqual(expect.objectContaining({ kind: 'element', name: 'w:bogus', line }));
    expect(problems).toContainEqual(expect.objectContaining({ kind: 'element', name: 'x:thing', line: line + 1 }));
    expect(problems).toContainEqual(expect.objectContaining({ kind: 'attribute', name: 'w:foo', line: line - 1 }));
    const bogusProblem = problems.find((p) => p.name === 'w:bogus')!;
    expect(bogusProblem.path).toMatch(/^\/w:styles\[1\]\/w:style\[\d+\]$/);
  });

  test('a text that does not parse, and a value of the wrong type, are refused with the message and its line', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const styles = part(session, '/word/styles.xml');
    const pretty = (await xmlOf(styles)).replace(/></g, '>\n<');
    const malformed = await checkPartXml(styles, pretty.replace('<w:name w:val="Normal"/>', '<w:name w:val="Normal"'));
    expect(malformed).toHaveLength(1);
    expect(malformed[0]!.kind).toBe('parse');
    expect(malformed[0]!.line).toBeGreaterThan(1);
    // The engine reads a value its type refuses through, dropped, with a warning a person can act on (core-ts 0.7.0): quoted.
    const wrong = await checkPartXml(styles, pretty.replace(/<w:sz w:val="\d+"\/>/, '<w:sz w:val="large"/>'));
    expect(wrong).toHaveLength(1);
    expect(wrong[0]!.kind).toBe('value');
    expect(wrong[0]!.message).toMatch(/large/);
    expect(wrong[0]!.message).toMatch(/Word 365 refuses/);
    expect(wrong[0]!.line).toBeGreaterThan(1);
  });

  test('a custom XML data part needs only to parse', async () => {
    const session = await EditorDocument.open(await fixture('opendope/binding-simple.docx'));
    const data = [...session.pkg.parts].find((p) => /^\/customXml\/item\d+\.xml$/.test(p.partName.toString()) && partEditing(p).mode === 'bytes')!;
    expect(await checkPartXml(data, '<anything><at><all/></at></anything>')).toEqual([]);
    expect((await checkPartXml(data, '<unclosed>'))[0]!.kind).toBe('parse');
  });
});

describe('an Apply and what follows (section 4.4)', () => {
  test('a style changed: the part written, the resolver reading it, the main part untouched at save, and a reload has it', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    const styles = part(session, '/word/styles.xml');
    const before = session.pkg.propertyResolver.getEffectiveRPr(undefined, { pStyle: { val: 'Heading1' } } as never);
    expect(before.color?.val).not.toBe('FF0000');
    const xml = withColor(await xmlOf(styles), 'Heading1', 'FF0000');
    expect(await checkPartXml(styles, xml)).toEqual([]);
    const applied = await applyPartXml(session, styles, xml);
    expect(applied.dangling).toEqual([]);
    expect(session.touchedParts.has('/word/styles.xml')).toBe(true);
    expect(session.pkg.propertyResolver.getEffectiveRPr(undefined, { pStyle: { val: 'Heading1' } } as never).color?.val).toBe('FF0000');
    const saved = await session.save();
    expect(comparePackages(bytes, saved.bytes, ['word/styles.xml', 'docProps/app.xml'])).toEqual([]);
    const reloaded = await WordprocessingMLPackage.load(saved.bytes);
    const heading = (await reloaded.getMainDocumentPart().styleDefinitionsPart!.readContents()).style!.find((s) => s.styleId === 'Heading1')!;
    expect(heading.rPr?.color?.val).toBe('FF0000');
  });

  test('a style removed is said as dangling, not refused', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const styles = part(session, '/word/styles.xml');
    const xml = (await xmlOf(styles)).replace(/<w:style [^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/, '');
    expect(await checkPartXml(styles, xml)).toEqual([]);
    const applied = await applyPartXml(session, styles, xml);
    expect(applied.dangling).toContainEqual(expect.objectContaining({ kind: 'paragraphStyle', id: 'Heading1' }));
  });

  test('a numbering format changed: the live labels follow', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const numbering = part(session, '/word/numbering.xml');
    const labelsOf = (): string[] => [...liveLabels(session.doc, session).values()].map((l) => l.listString);
    const before = labelsOf();
    expect(before.length).toBeGreaterThan(0);
    const xml = (await xmlOf(numbering)).replace(/w:numFmt w:val="decimal"/g, 'w:numFmt w:val="upperRoman"');
    expect(xml).toContain('upperRoman');
    expect(await checkPartXml(numbering, xml)).toEqual([]);
    await applyPartXml(session, numbering, xml);
    const after = labelsOf();
    expect(after).not.toEqual(before);
    // The one decimal list in use: "1)" becomes "I)".
    expect(after.some((l) => /^[IVX]+[.)]?$/.test(l))).toBe(true);
  });

  test('the settings read again: the default tab stop', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const settings = part(session, '/word/settings.xml');
    const xml = (await xmlOf(settings)).replace(/<w:defaultTabStop w:val="\d+"\/>/, '<w:defaultTabStop w:val="1440"/>');
    expect(xml).toContain('w:val="1440"');
    await applyPartXml(session, settings, xml);
    expect(session.defaultTabStop).toBe(1440);
  });

  test('a data part edited: the bound control\'s verdict follows', async () => {
    const session = await EditorDocument.open(await fixture('opendope/binding-simple.docx'));
    await session.template.load();
    let control: unknown;
    session.doc.descendants((n) => { if (!control && n.attrs['sdtPr'] && readControl(n.attrs['sdtPr']).binding) control = n.attrs['sdtPr']; return !control; });
    expect(control).toBeDefined();
    const binding = readControl(control).binding!;
    const view = session.template.partNamed(binding.storeItemID ?? '')!;
    const before = session.template.verdictFor(control).value;
    const data = view.part as unknown as Part;
    const xml = view.getXml();
    const changed = before ? xml.replace(`>${before}<`, '>Changed by the parts panel<') : xml;
    expect(changed).not.toBe(xml);
    expect(await checkPartXml(data, changed)).toEqual([]);
    let heard = 0;
    session.template.subscribe(() => { heard++; });
    await applyPartXml(session, data, changed);
    expect(heard).toBeGreaterThan(0);
    expect(session.template.verdictFor(control).value).toBe('Changed by the parts panel');
  });
});

describe('an Apply in a room (section 4.5)', () => {
  for (const order of RELEASE_ORDERS) {
    test(`two peers apply the styles part at once, each changing another style (Heading1, Title): one entry each, both kept (${order})`, async () => {
      const r = await room(await fixture('tracked-changes.docx'));
      await r.creator.partsReady();
      r.link.hold();
      const a = part(r.creator, '/word/styles.xml');
      const b = part(r.joiner, '/word/styles.xml');
      await applyPartXml(r.creator, a, withColor(await xmlOf(a), 'Heading1', 'FF0000'));
      await applyPartXml(r.joiner, b, withColor(await xmlOf(b), 'Title', '00FF00'));
      await r.creator.partsReady();
      await r.joiner.partsReady();
      // Each Apply of the whole part published the one style it changed.
      expect([...r.creator.store.styles.keys()]).toEqual(['style:Heading1']);
      expect([...r.joiner.store.styles.keys()]).toEqual(['style:Title']);
      r.link.release(order);
      for (let i = 0; i < 3; i++) { await r.creator.partsReady(); await r.joiner.partsReady(); }
      for (const s of [r.creator, r.joiner]) {
        const tree = await (part(s, '/word/styles.xml') as XmlPart<{ style?: { styleId?: string; rPr?: { color?: { val?: string } } }[] }>).getContents();
        expect(tree.style!.find((x) => x.styleId === 'Heading1')!.rPr?.color?.val).toBe('FF0000');
        expect(tree.style!.find((x) => x.styleId === 'Title')!.rPr?.color?.val).toBe('00FF00');
      }
      const [x, y] = [await r.creator.save({ selfCheck: false }), await r.joiner.save({ selfCheck: false })];
      expect(comparePackages(x.bytes, y.bytes)).toEqual([]);
    });
  }

  test('a numbering part applied whole reaches the peer as the definitions it changed; a settings part as the part', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    await r.creator.partsReady();
    const numbering = part(r.creator, '/word/numbering.xml');
    const xml = await xmlOf(numbering);
    const first = /<w:abstractNum [^>]*w:abstractNumId="(\d+)"[\s\S]*?<\/w:abstractNum>/.exec(xml)!;
    await applyPartXml(r.creator, numbering, xml.replace(first[0], first[0].replace(/w:numFmt w:val="\w+"/, 'w:numFmt w:val="lowerLetter"')));
    await r.creator.partsReady();
    expect([...r.creator.store.numbering.keys()]).toEqual([`abstract:${first[1]}`]);
    const settings = part(r.creator, '/word/settings.xml');
    await applyPartXml(r.creator, settings, (await xmlOf(settings)).replace(/<w:defaultTabStop w:val="\d+"\/>/, '<w:defaultTabStop w:val="1080"/>'));
    for (let i = 0; i < 3; i++) { await r.creator.partsReady(); await r.joiner.partsReady(); }
    expect(r.joiner.defaultTabStop).toBe(1080);
    expect(await xmlOf(part(r.joiner, '/word/numbering.xml'))).toContain('lowerLetter');
  });
});
