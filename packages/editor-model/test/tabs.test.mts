// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.8 item 5 (step C5): the stops merged from the three sources in the order
// Word applies them (through the resolver), and the width computation over given positions,
// since jsdom cannot lay out; the browser test measures a paragraph against its stop.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, stateFor, runCommand, documentStyles, nextTabStop, tabWidth, tabStopsOf, paragraphTabs, DEFAULT_TAB_STOP, applyParagraphAttrs, paragraphAttrsOf, tabAttrsOf, type TabStop, type ParagraphAttrs } from '../src/index.mjs';
import type { Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { fixture, comparePackages, mainXml } from './helpers.mjs';

describe('nextTabStop over given positions', () => {
  const stops: TabStop[] = [{ pos: 720, val: 'left' }, { pos: 2160, val: 'center' }, { pos: 9990, val: 'right', leader: 'dot' }];
  test('the first custom stop past the start, then the defaults past the last custom stop', () => {
    expect(nextTabStop(0, stops)).toEqual({ pos: 720, stop: stops[0] });
    expect(nextTabStop(720, stops)).toEqual({ pos: 2160, stop: stops[1] });   // at a stop: the next one
    expect(nextTabStop(3000, stops)).toEqual({ pos: 9990, stop: stops[2] });
    expect(nextTabStop(9990, stops)).toEqual({ pos: 10080 });                  // past the last: the next multiple of 720 beyond it
    expect(nextTabStop(10100, stops)).toEqual({ pos: 10800 });
  });
  test('no custom stops: the default interval from the column edge, at any position', () => {
    expect(nextTabStop(0, [])).toEqual({ pos: 720 });
    expect(nextTabStop(100, [])).toEqual({ pos: 720 });
    expect(nextTabStop(720, [])).toEqual({ pos: 1440 });
    expect(nextTabStop(1441, [], 567)).toEqual({ pos: 1701 });
    expect(nextTabStop(5, [], 0)).toEqual({ pos: DEFAULT_TAB_STOP });        // a nonsense interval: Word's default
  });
  test('the default stops begin after the last custom stop, not between them', () => {
    // A custom stop at 3600 with the start at 100: 3600, not 720.
    expect(nextTabStop(100, [{ pos: 3600, val: 'left' }])).toEqual({ pos: 3600, stop: { pos: 3600, val: 'left' } });
  });
  test('a hanging indent is an implicit stop at the paragraph\'s left indent', () => {
    const indent = { start: 1080, end: null, firstLine: null, hanging: 360 };
    expect(nextTabStop(400, [], 720, indent)).toEqual({ pos: 1080, stop: { pos: 1080, val: 'left' } });
    expect(nextTabStop(400, stops, 720, indent)).toEqual({ pos: 720, stop: stops[0] });
    expect(nextTabStop(800, stops, 720, indent)).toEqual({ pos: 1080, stop: { pos: 1080, val: 'left' } });
    // A first-line indent is not one.
    expect(nextTabStop(400, [], 720, { start: 1080, end: null, firstLine: 360, hanging: null })).toEqual({ pos: 720 });
  });
  test('tabStopsOf drops clear and bar entries and sorts', () => {
    expect(tabStopsOf({ tab: [{ pos: 1440, val: 'right', leader: 'dot' }, { pos: 720, val: 'clear' }, { pos: 360, val: 'bar' }, { pos: 100, val: 'left' }] })).toEqual([
      { pos: 100, val: 'left', leader: undefined }, { pos: 1440, val: 'right', leader: 'dot' },
    ]);
    expect(tabStopsOf(undefined)).toEqual([]);
  });
});

/** The first paragraph whose effective stops are not empty, with its position. */
function paragraphWithStops(session: EditorDocument): { node: PMNode; pos: number; stops: TabStop[] } | undefined {
  const state = stateFor(session);
  let found: { node: PMNode; pos: number; stops: TabStop[] } | undefined;
  state.doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type.name !== 'paragraph') return true;
    const { stops } = paragraphTabs(node, state);
    if (stops.length > 0) found = { node, pos, stops };
    return false;
  });
  return found;
}

describe('the effective stops through the resolver (5.6)', () => {
  test('headers-footers.docx: a paragraph\'s own w:tabs over its style\'s, ascending, clear removed; the default interval from the settings part', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    expect(session.defaultTabStop).toBe(720);
    const found = paragraphWithStops(session);
    expect(found).toBeDefined();
    const positions = found!.stops.map((s) => s.pos);
    expect([...positions].sort((a, b) => a - b)).toEqual(positions);
    expect(found!.stops.every((s) => s.val !== 'clear' && s.val !== 'bar')).toBe(true);
    // A style with stops applied to a plain paragraph: its stops are the paragraph's now.
    const state = stateFor(session);
    const resolver = session.pkg.propertyResolverOrUndefined!;
    const styled = (await documentStyles(session)).filter((st) => st.type === 'paragraph').map((st) => st.id).find((id) => (resolver.getEffectivePPr({ pStyle: { val: id } })?.tabs?.tab?.length ?? 0) > 0);
    expect(styled).toBeDefined();
    let plainPos = -1;
    state.doc.descendants((node, pos) => { if (plainPos < 0 && node.type.name === 'paragraph' && (node.attrs as ParagraphAttrs).pStyle === null && node.childCount > 0) plainPos = pos; return plainPos < 0; });
    expect(paragraphTabs(state.doc.nodeAt(plainPos)!, state).stops).toEqual([]);
    runCommand(session, 'paragraph.style', { value: styled }, { from: plainPos + 1 });
    const after = stateFor(session);
    expect(paragraphTabs(after.doc.nodeAt(plainPos)!, after).stops.length).toBeGreaterThan(0);
  });

  test('a document that states no default tab stop has Word\'s, and a session without a resolver answers no stops', async () => {
    const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
    const session = await EditorDocument.open(pkg);
    expect(session.defaultTabStop).toBeGreaterThan(0);
    const bare = await EditorDocument.open(await fixture('headers-footers.docx'), { resolver: false });
    expect(bare.defaultTabStop).toBe(DEFAULT_TAB_STOP);
    const state = stateFor(bare);
    expect(paragraphTabs(state.doc.child(0), state)).toEqual({ stops: [], indent: null });
  });
});

// ---------------------------------------------------------------------------------------------
// ED-003 section 8.5 (step F4): the paragraph's own stops as the `tabs` attribute, the command,
// the settings part's default interval written.
// ---------------------------------------------------------------------------------------------

/** The first paragraph that states its own `w:tabs`, with its source element. */
function paragraphWithOwnTabs(session: EditorDocument): { node: PMNode; pos: number; source: Element<wml.P> } | undefined {
  const state = stateFor(session);
  let found: { node: PMNode; pos: number; source: Element<wml.P> } | undefined;
  state.doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type.name !== 'paragraph') return true;
    const attrs = node.attrs as ParagraphAttrs;
    if (attrs.tabs !== null && attrs.sourceId) found = { node, pos, source: session.ctx.ids.get(attrs.sourceId) as Element<wml.P> };
    return false;
  });
  return found;
}

describe('the paragraph\'s own stops (8.5, step F4)', () => {
  test('headers-footers.docx: a paragraph\'s w:tabs is its tabs attribute, in the element\'s order, every entry kept; an unchanged attribute leaves the element alone', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const found = paragraphWithOwnTabs(session);
    expect(found).toBeDefined();
    const element = found!.source.value.pPr!.tabs!.tab;
    const attrs = found!.node.attrs as ParagraphAttrs;
    expect(attrs.tabs).toEqual(element.map((t) => ({ pos: t.pos, val: String(t.val), leader: t.leader === undefined ? null : String(t.leader) })));
    const base = paragraphAttrsOf(found!.source.value.pPr);
    expect(applyParagraphAttrs(structuredClone(found!.source.value.pPr), attrs, base)?.tabs).toEqual(found!.source.value.pPr!.tabs);
  });

  test('applyParagraphAttrs replaces w:tabs whole when the attribute differs, clear entries included, and removes it for null', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const found = paragraphWithOwnTabs(session)!;
    const base = paragraphAttrsOf(found.source.value.pPr);
    const attrs = found.node.attrs as ParagraphAttrs;
    const changed = applyParagraphAttrs(structuredClone(found.source.value.pPr), { ...attrs, tabs: [{ pos: 720, val: 'clear', leader: null }, { pos: 1440, val: 'right', leader: 'dot' }] }, base);
    expect(changed?.tabs?.tab).toMatchObject([{ pos: 720, val: 'clear' }, { pos: 1440, val: 'right', leader: 'dot' }]);
    expect('leader' in changed!.tabs!.tab[0]!).toBe(false);
    const removed = applyParagraphAttrs(structuredClone(found.source.value.pPr), { ...attrs, tabs: null }, base);
    expect(removed?.tabs).toBeUndefined();
  });

  test('tabAttrsOf: points to twips, ascending, one per position (the last wins), a none leader dropped, a value w:tab does not take refused', () => {
    expect(tabAttrsOf([{ position: 72, alignment: 'right', leader: 'dot' }, { position: 36 }, { position: 36, alignment: 'center', leader: 'none' }]))
      .toEqual([{ pos: 720, val: 'center', leader: null }, { pos: 1440, val: 'right', leader: 'dot' }]);
    expect(tabAttrsOf([])).toEqual([]);
    expect(tabAttrsOf([{ position: 10, alignment: 'middle' }])).toBeUndefined();
    expect(tabAttrsOf([{ position: 10, leader: 'stars' }])).toBeUndefined();
    expect(tabAttrsOf([{ position: Number.NaN }])).toBeUndefined();
  });

  test('paragraph.tabStops on a plain paragraph: the stop is effective at once and exported as w:tabs; [] removes it; a bad value is refused', async () => {
    const bytes = await fixture('headers-footers.docx');
    const session = await EditorDocument.open(bytes);
    const state = stateFor(session);
    let plainPos = -1;
    state.doc.descendants((node, pos) => { if (plainPos < 0 && node.type.name === 'paragraph' && (node.attrs as ParagraphAttrs).tabs === null && (node.attrs as ParagraphAttrs).pStyle === null && node.childCount > 0) plainPos = pos; return plainPos < 0; });
    expect(paragraphTabs(state.doc.nodeAt(plainPos)!, state).stops).toEqual([]);
    expect(runCommand(session, 'paragraph.tabStops', { stops: [{ position: 100, alignment: 'right', leader: 'dot' }, { position: 36 }] }, { from: plainPos + 1 }).ok).toBe(true);
    const after = stateFor(session);
    const node = after.doc.nodeAt(plainPos)!;
    expect((node.attrs as ParagraphAttrs).tabs).toEqual([{ pos: 720, val: 'left', leader: null }, { pos: 2000, val: 'right', leader: 'dot' }]);
    expect(paragraphTabs(node, after).stops).toEqual([{ pos: 720, val: 'left', leader: undefined }, { pos: 2000, val: 'right', leader: 'dot' }]);
    const saved = await session.save();
    expect(saved.changed).toBe(true);
    const xml = await mainXml(saved.bytes);
    expect(xml).toMatch(/<w:tab [^>]*w:pos="2000"/);
    expect(xml).toMatch(/<w:tab [^>]*w:leader="dot"/);
    expect(comparePackages(bytes, saved.bytes, ['word/document.xml', 'docProps/app.xml'])).toEqual([]);
    expect(runCommand(session, 'paragraph.tabStops', { stops: [] }, { from: plainPos + 1 }).ok).toBe(true);
    expect((stateFor(session).doc.nodeAt(plainPos)!.attrs as ParagraphAttrs).tabs).toBeNull();
    expect(await mainXml((await session.save()).bytes)).not.toMatch(/w:pos="2000"/);
    expect(runCommand(session, 'paragraph.tabStops', { stops: [{ position: 10, alignment: 'sideways' }] }, { from: plainPos + 1 }).ok).toBe(false);
  });

  test('a clear entry cancels a style\'s stop at that position and leaves the others', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const state = stateFor(session);
    const resolver = session.pkg.propertyResolverOrUndefined!;
    const styled = (await documentStyles(session)).filter((st) => st.type === 'paragraph').map((st) => st.id).find((id) => (resolver.getEffectivePPr({ pStyle: { val: id } })?.tabs?.tab?.length ?? 0) > 1)!;
    let plainPos = -1;
    state.doc.descendants((node, pos) => { if (plainPos < 0 && node.type.name === 'paragraph' && (node.attrs as ParagraphAttrs).tabs === null && (node.attrs as ParagraphAttrs).pStyle === null && node.childCount > 0) plainPos = pos; return plainPos < 0; });
    runCommand(session, 'paragraph.style', { value: styled }, { from: plainPos + 1 });
    const withStyle = stateFor(session);
    const stops = paragraphTabs(withStyle.doc.nodeAt(plainPos)!, withStyle).stops;
    expect(stops.length).toBeGreaterThan(1);
    const first = stops[0]!;
    expect(runCommand(session, 'paragraph.tabStops', { stops: [{ position: first.pos / 20, alignment: 'clear' }] }, { from: plainPos + 1 }).ok).toBe(true);
    const cleared = stateFor(session);
    const remaining = paragraphTabs(cleared.doc.nodeAt(plainPos)!, cleared).stops;
    expect(remaining.map((s) => s.pos)).toEqual(stops.slice(1).map((s) => s.pos));
    expect((cleared.doc.nodeAt(plainPos)!.attrs as ParagraphAttrs).tabs).toEqual([{ pos: first.pos, val: 'clear', leader: null }]);
  });

  test('setDefaultTabStop writes the settings part: the session reads the new interval, the part is marked, the save carries it, the main part is untouched; a bad value is refused', async () => {
    const bytes = await fixture('headers-footers.docx');
    const session = await EditorDocument.open(bytes);
    expect(session.defaultTabStop).toBe(720);
    await session.setDefaultTabStop(1440);
    expect(session.defaultTabStop).toBe(1440);
    expect(session.touchedParts.has('/word/settings.xml')).toBe(true);
    expect(nextTabStop(0, [], session.defaultTabStop)).toEqual({ pos: 1440 });
    const saved = await session.save();
    expect(comparePackages(bytes, saved.bytes, ['word/settings.xml', 'docProps/app.xml'])).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    expect(reopened.defaultTabStop).toBe(1440);
    await expect(session.setDefaultTabStop(0)).rejects.toThrow();
    await expect(session.setDefaultTabStop(Number.NaN)).rejects.toThrow();
  });
});

describe('tabWidth: the tab shortened by the text that follows it for a centre, right or decimal stop (11.20, step F5)', () => {
  const at = (val: string): { pos: number; stop: TabStop } => ({ pos: 6000, stop: { pos: 6000, val } });
  test('a left stop and a default stop reach the stop whatever follows', () => {
    expect(tabWidth(1000, at('left'), { width: 2000, toDecimal: null })).toBe(5000);
    expect(tabWidth(1000, { pos: 1440 }, { width: 2000, toDecimal: null })).toBe(440);
  });
  test('a right stop ends the following text at the stop; a centre stop centres it; a decimal stop puts the separator there', () => {
    expect(tabWidth(1000, at('right'), { width: 2000, toDecimal: null })).toBe(3000);
    expect(tabWidth(1000, at('end'), { width: 2000, toDecimal: null })).toBe(3000);
    expect(tabWidth(1000, at('center'), { width: 2000, toDecimal: null })).toBe(4000);
    expect(tabWidth(1000, at('decimal'), { width: 2000, toDecimal: 500 })).toBe(4500);
    expect(tabWidth(1000, at('decimal'), { width: 2000, toDecimal: null })).toBe(3000);   // no separator: as a right stop
  });
  test('text wider than the room collapses the tab', () => {
    expect(tabWidth(5000, at('right'), { width: 2000, toDecimal: null })).toBe(0);
    expect(tabWidth(5000, at('center'), { width: 2000, toDecimal: null })).toBe(0);
  });
  test('the decimal symbol comes from the settings part, with a full stop when it states none', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    expect(session.decimalSymbol).toBe('.');
    const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
    expect((await EditorDocument.open(pkg)).decimalSymbol.length).toBe(1);
  });
});
