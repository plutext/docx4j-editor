// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word fragments (ED-005 proposal 46, accepted 2026-11-11): a control bound to a node holding a Word
// document as escaped Flat OPC (`pkg:package`). Two forms: OpenDoPE's, a control tagged
// `od:progid=Word.Document` bound through its `od:xpath` entry, which a processor resolves (the
// specification's "Word Fragments", REQ-059 and REQ-060); and Word's own, a rich text control mapped
// by `w15:dataBinding`, which Word 365 fills from the node when it opens the file (the proposal's
// part 6(a), `test/fixtures/fragments/`). The engine's bind step resolves both (core-ts CR-005, the
// proposal's parts 1 and 2); this module reads what a control is and what its fragment holds.
import type { Node as PMNode } from 'prosemirror-model';
import { WordprocessingMLPackage, applyFlatOpcBindingTo, type FragmentCreated } from '@docx4j/core-ts';
import { readTag } from './opendope.mjs';
import { isFlatOpc, boundNodeOf } from './bound.mjs';
import { sdtTag, sdtKind, sdtBinding, sdtId, type DataBinding } from './sdtpr.mjs';
import type { TemplateModel } from './template.mjs';
import type { EditorDocument } from './session.mjs';
import { LOCAL_ORIGIN } from './yjs.mjs';

/** The `od:progid` value that marks OpenDoPE's Word fragment control. */
export const WORD_DOCUMENT_PROGID = 'Word.Document';

/** Whether a control's tag makes it OpenDoPE's Word fragment control. */
export function isFragmentControl(sdtPr: unknown): boolean {
  return readTag(sdtTag(sdtPr)).progid === WORD_DOCUMENT_PROGID;
}

export interface FragmentBinding {
  binding: DataBinding;
  /** OpenDoPE's tagged control, or Word's rich text control mapped by `w15:dataBinding`. */
  shape: 'progid' | 'w15';
}

/**
 * A control's fragment binding, or undefined: OpenDoPE's through the tag's `od:xpath` entry (the template must be
 * loaded), or Word's `w15:dataBinding` on a rich text control. Whether the node holds Flat OPC is the caller's to read:
 * a `w15`-mapped rich text control holding plain text is Word's ordinary mapping.
 */
export function fragmentBindingOf(sdtPr: unknown, template: TemplateModel): FragmentBinding | undefined {
  if (isFragmentControl(sdtPr)) {
    const id = readTag(sdtTag(sdtPr)).xpath;
    const entry = id !== undefined ? template.xpaths?.entries.find((e) => e.id === id) : undefined;
    if (!entry || entry.xpath === '') return undefined;
    return { binding: { xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings, w15: false }, shape: 'progid' };
  }
  const binding = sdtBinding(sdtPr);
  if (sdtKind(sdtPr) === 'RichText' && binding?.w15 === true && binding.xpath !== '') return { binding, shape: 'w15' };
  return undefined;
}

/** What a fragment holds, and what binding it would leave behind (the proposal's part 1, from Word 365's measurements). */
export interface FragmentFacts {
  /** The package's size as text, in bytes (UTF-8). */
  bytes: number;
  /** The body's own paragraphs and tables, not those inside tables or controls. */
  paragraphs: number;
  tables: number;
  /** Pictures anywhere in the body (`w:drawing`, `w:pict`). */
  pictures: number;
  /** Footnote and endnote references: kept, as Word keeps them. */
  notes: number;
  /** Controls inside the fragment: kept, and not bound. */
  controls: number;
  /** Comment references: dropped until measured. */
  comments: number;
  /** Section breaks inside the fragment (a paragraph's `w:sectPr`): dropped, as Word drops the fragment's section. */
  sectionBreaks: number;
  /** The body's text, its runs' `w:t` in order, for a loose comparison with what a control shows (the stale count). */
  text: string;
}

const factsCache = new Map<string, Promise<FragmentFacts | undefined>>();

const decodeXml = (s: string): string => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const count = (xml: string, pattern: RegExp): number => (xml.match(pattern) ?? []).length;

/**
 * The facts of a node's value as a Word fragment, read through the engine, or undefined when the value is not Flat OPC
 * or does not load. Cached for the last sixteen values, since the Data and Control tabs ask on every render and a
 * fragment runs to tens of kilobytes.
 */
export function fragmentFactsOf(value: string): Promise<FragmentFacts | undefined> {
  const cached = factsCache.get(value);
  if (cached) return cached;
  const facts = isFlatOpc(value) ? readFacts(value).catch(() => undefined) : Promise.resolve(undefined);
  if (factsCache.size >= 16) factsCache.delete(factsCache.keys().next().value!);
  factsCache.set(value, facts);
  return facts;
}

async function readFacts(value: string): Promise<FragmentFacts | undefined> {
  const pkg = await WordprocessingMLPackage.load(value);
  const main = pkg.getMainDocumentPart();
  const document = await main.getContents();
  const content = (document.body?.content ?? []) as unknown[];
  const typeOf = (item: unknown): string | undefined => {
    const o = item as { TYPE_NAME?: string; value?: { TYPE_NAME?: string } } | null;
    return o?.TYPE_NAME ?? o?.value?.TYPE_NAME;
  };
  const xml = await main.getXml();
  // The body's final w:sectPr is the fragment's page, which Word drops; one inside a paragraph's properties is a break.
  const sections = count(xml, /<w:sectPr[\s>/]/g) - (document.body?.sectPr ? 1 : 0);
  return {
    bytes: new TextEncoder().encode(value).length,
    paragraphs: content.filter((c) => typeOf(c) === 'org_docx4j_wml.P').length,
    tables: content.filter((c) => typeOf(c) === 'org_docx4j_wml.Tbl').length,
    pictures: count(xml, /<w:(?:drawing|pict)[\s>]/g),
    notes: count(xml, /<w:(?:footnote|endnote)Reference[\s>/]/g),
    controls: count(xml, /<w:sdt[\s>]/g),
    comments: count(xml, /<w:commentReference[\s>/]/g),
    sectionBreaks: Math.max(0, sections),
    text: [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => decodeXml(m[1]!)).join(''),
  };
}

export interface FragmentControl {
  pos: number;
  node: PMNode;
  fragment: FragmentBinding;
}

/** Every fragment-bound control of a document, either form, in document order (a `w15` mapping whatever its node holds). */
export function fragmentControlsIn(doc: PMNode, template: TemplateModel): FragmentControl[] {
  const out: FragmentControl[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'sdtBlock' && node.type.name !== 'sdtInline') return true;
    const fragment = fragmentBindingOf(node.attrs['sdtPr'], template);
    if (fragment) out.push({ pos, node, fragment });
    return true;
  });
  return out;
}

export interface FragmentRefresh {
  /** Controls whose content was written. */
  changed: number;
  /** What the engine's pass left or dropped, its own words. */
  notes: string[];
}

/** Per session, the value each control was last filled from and what that run added, so an unchanged value is not filled again. */
const lastRuns = new WeakMap<EditorDocument, Map<number, { value: string; created?: FragmentCreated }>>();

/**
 * The editor's refresh of its Word fragment controls (ED-005 proposal 46, part 3): each control whose node holds a Word
 * document as Flat OPC, and whose value changed since it was last filled, is filled by the engine's bind pass
 * (`applyFlatOpcBindingTo`, core-ts CR-005 section 11) in one tree run, untracked and out of the undo history, as the
 * XHTML and picture refreshes write. Both forms: OpenDoPE's `od:progid` control and Word's `w15`-mapped rich text, whose
 * mapping is kept (the editor authors; dropping it is a processor's choice). `ids` limits it to those controls (a tab's own
 * control), `shape` to one form (the stale banner's Refresh, Word's alone). Never at open: a file round-trips as it is, and Word's form is offered through the stale banner instead.
 */
export async function refreshFragments(session: EditorDocument, options: { ids?: ReadonlySet<number>; shape?: FragmentBinding['shape'] } = {}): Promise<FragmentRefresh> {
  const result: FragmentRefresh = { changed: 0, notes: [] };
  if (!session.template.loaded) return result;
  let cache = lastRuns.get(session);
  if (!cache) { cache = new Map(); lastRuns.set(session, cache); }
  const due: { id: number; value: string }[] = [];
  for (const c of fragmentControlsIn(session.doc, session.template)) {
    const id = sdtId(c.node.attrs['sdtPr']);
    if (id === undefined || (options.ids && !options.ids.has(id)) || (options.shape && c.fragment.shape !== options.shape)) continue;
    const value = boundNodeOf(session, c.fragment.binding)?.nodeValue ?? '';
    if (!isFlatOpc(value) || cache.get(id)?.value === value) continue;
    due.push({ id, value });
  }
  if (due.length === 0) return result;
  const run = await session.runOnTree(async ({ pkg, body }) => {
    const outcomes: { id: number; value: string; done: boolean; created?: FragmentCreated; note?: string }[] = [];
    for (const d of due) {
      const control = body.contentControls.find((k) => k.id === d.id);
      if (!control) continue;
      const outcome = await applyFlatOpcBindingTo(control, pkg.customXmlParts, pkg);
      outcomes.push({ id: d.id, value: d.value, done: outcome.done, ...(outcome.created ? { created: outcome.created } : {}), ...(outcome.note ? { note: outcome.note } : {}) });
    }
    // What a control's previous fill added and nothing references now goes: its lists and its pictures (core-ts's
    // `created`). A copied style stays, as a person may use it by now and a stray style costs nothing.
    const earlier = outcomes.filter((o) => o.done).map((o) => cache!.get(o.id)?.created).filter((c): c is FragmentCreated => c !== undefined);
    if (earlier.length > 0) await removeUnused(pkg, earlier);
    return outcomes;
  }, LOCAL_ORIGIN, { author: session.trackingAuthor ?? '', mode: 'Off' }, { history: false });
  if (run.error) throw run.error;
  let done = false;
  for (const o of run.value ?? []) {
    if (o.note) result.notes.push(o.note);
    if (!o.done) continue;
    cache.set(o.id, { value: o.value, ...(o.created ? { created: o.created } : {}) });
    result.changed++;
    done = true;
  }
  // The parts a fill writes besides the body, published to a room's peers as any part edit is (the pictures are binary,
  // which the part sync does not carry: a peer's save lacks them, as a picture made in the editor's does).
  if (done) {
    const main = session.main;
    for (const part of [main.styleDefinitionsPart, main.numberingDefinitionsPart, main.footnotesPart, main.endnotesPart]) {
      if (part) session.markPartChanged(part.partName.toString());
    }
  }
  return result;
}

/** Every relationship id and list a story or the styles can reference, as XML: the main part, the notes, the headers and footers, the styles. */
async function referencesXml(pkg: WordprocessingMLPackage): Promise<string> {
  const main = pkg.getMainDocumentPart();
  const parts: { getXml(): Promise<string> }[] = [main];
  for (const p of [main.styleDefinitionsPart, main.footnotesPart, main.endnotesPart, ...main.headerParts, ...main.footerParts]) if (p) parts.push(p as unknown as { getXml(): Promise<string> });
  return (await Promise.all(parts.map((p) => p.getXml()))).join('\n');
}

/** A previous fill's lists and pictures that nothing references now, removed from the numbering part and the package. */
async function removeUnused(pkg: WordprocessingMLPackage, created: FragmentCreated[]): Promise<void> {
  const xml = await referencesXml(pkg);
  const main = pkg.getMainDocumentPart();
  const numbering = main.numberingDefinitionsPart ? await main.numberingDefinitionsPart.getContents() : undefined;
  if (numbering) {
    const nums = new Set(created.flatMap((c) => c.numIds).filter((id) => !new RegExp(`<w:numId w:val="${id}"\\s*/>`).test(xml)));
    if (nums.size > 0) numbering.num = (numbering.num ?? []).filter((n) => !nums.has(String(n.numId)));
    const used = new Set((numbering.num ?? []).map((n) => String(n.abstractNumId?.val)));
    const abstracts = new Set(created.flatMap((c) => c.abstractNumIds).filter((id) => !used.has(id)));
    if (abstracts.size > 0) numbering.abstractNum = (numbering.abstractNum ?? []).filter((a) => !abstracts.has(String(a.abstractNumId)));
  }
  const rels = main.relationshipsPart;
  if (!rels) return;
  for (const name of created.flatMap((c) => c.parts).filter((p) => p.startsWith('/word/media/'))) {
    const pointing = rels.list.filter((r) => { try { return rels.resolveTarget(r).toString() === name; } catch { return false; } });
    if (pointing.every((r) => !xml.includes(`"${r.id}"`))) rels.removePart(name);
  }
}

