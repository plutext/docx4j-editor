// SPDX-License-Identifier: GPL-3.0-or-later
//
// The parts made editable (ED-005 section 4, proposal 21): which parts the Developer
// perspective's parts panel edits and how, the check before an Apply, and the Apply with its
// effects. The package parts are edited, never the projected ones: a raw edit of the main
// document, a note, a header or a comment would leave the projection stale, so those stay read
// only and say where they are edited. A typed part is checked by unmarshalling the text through
// the engine into a scratch part of the same class and marshalling it back: an element or an
// attribute the object model drops is listed with where it is, and the Apply refuses. The verdict
// is the engine's parse alone; `@xmldom/xmldom`, which records lines, is asked only where each
// problem is, and a line is given only when the element it lands on has the problem's name
// (section 4.3). A part edit has no undo (proposal 16): the pane's Revert is another Apply.
import {
  XmlPart, DefaultXmlPart, BinaryPart, RelationshipsPart, CustomXmlDataStoragePart, PartUnmarshalException,
  MainDocumentPart, FootnotesPart, EndnotesPart, HeaderPart, FooterPart, CommentsPart, CommentsExtendedPart, CommentsIdsPart,
  CommentsExtensiblePart, PeoplePart, StyleDefinitionsPart, NumberingDefinitionsPart, DocumentSettingsPart, ThemePart, FontTablePart,
  type Part,
} from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { EditorDocument } from './session.mjs';

/** How the parts panel edits a part. */
export type PartEditing =
  /** Unmarshalled and checked that nothing is lost: styles, numbering, settings, theme, fonts, `docProps`, a properties part. */
  | { mode: 'typed' }
  /** Checked for well-formedness only: a custom XML data part or an untyped part, which the engine keeps as XML. */
  | { mode: 'bytes' }
  /** Read only, and why: the reason names where it is edited instead. */
  | { mode: 'none'; reason: 'projected' | 'comments' | 'regenerated' | 'binary' | 'other' };

const PROJECTED = [MainDocumentPart, FootnotesPart, EndnotesPart, HeaderPart, FooterPart] as const;
const COMMENTS = [CommentsPart, CommentsExtendedPart, CommentsIdsPart, CommentsExtensiblePart, PeoplePart] as const;

/** The rule of section 4.2. */
export function partEditing(part: Part): PartEditing {
  if (part instanceof RelationshipsPart) return { mode: 'none', reason: 'regenerated' };
  if (PROJECTED.some((c) => part instanceof c)) return { mode: 'none', reason: 'projected' };
  if (COMMENTS.some((c) => part instanceof c)) return { mode: 'none', reason: 'comments' };
  if (part instanceof BinaryPart) return { mode: 'none', reason: 'binary' };
  if (part instanceof XmlPart) return { mode: 'typed' };
  if (part instanceof DefaultXmlPart) return { mode: 'bytes' };
  return { mode: 'none', reason: 'other' };
}

/** What would go wrong with an Apply: the text does not parse, a value has the wrong type, or an element or attribute would be dropped. */
export interface PartXmlProblem {
  kind: 'parse' | 'value' | 'element' | 'attribute';
  /** The name dropped, as the text writes it (`w:bogus`, `w:foo` for an attribute). */
  name?: string;
  /** The element that holds it, or the one the engine was on: `/w:styles[1]/w:style[4]`. */
  path?: string;
  /** The parser's or the engine's message. */
  message?: string;
  line?: number;
  column?: number;
}

/** A line and column in a parser's message ("line 3, column 5", "error on line 3 at column 5", "[3:5]"), when it gives them. */
function locationIn(message: string): { line?: number; column?: number } {
  const m = /line[^0-9]{0,8}(\d+)[^0-9]{1,20}col(?:umn)?[^0-9]{0,8}(\d+)/i.exec(message) ?? /\[(\d+):(\d+)\]/.exec(message);
  if (m) return { line: Number(m[1]), column: Number(m[2]) };
  const l = /line[^0-9]{0,8}(\d+)/i.exec(message);
  return l ? { line: Number(l[1]) } : {};
}

/**
 * Where a text that did not parse goes wrong, by xmldom's locator (its error handler's context,
 * or the error it throws): the engine's own message has no position in Node, and a browser's
 * parser says it in its words (`locationIn` reads those). Only the position is taken.
 */
async function parseErrorAt(xml: string): Promise<{ line?: number; column?: number }> {
  type Located = { locator?: { lineNumber?: number; columnNumber?: number } } | undefined;
  let found: { line?: number; column?: number } = {};
  const read = (at: Located): void => {
    const l = at?.locator;
    if (found.line !== undefined || typeof l?.lineNumber !== 'number') return;
    found = { line: l.lineNumber, ...(typeof l.columnNumber === 'number' ? { column: l.columnNumber } : {}) };
  };
  try {
    const { DOMParser } = await import('@xmldom/xmldom');
    new DOMParser({ onError: (level: string, _message: string, context: unknown) => { if (level !== 'warning') read(context as Located); } }).parseFromString(xml, 'text/xml');
  } catch (e) {
    read(e as Located);
  }
  return found;
}

/** An element's or attribute's expanded name. */
const expanded = (n: { namespaceURI: string | null; localName: string | null; nodeName: string }): string => `{${n.namespaceURI ?? ''}}${n.localName ?? n.nodeName}`;

interface Entry {
  kind: 'element' | 'attribute';
  /** As the text writes it. */
  name: string;
  /** The holder's path, as the text writes it. */
  path: string;
  node: Element;
}

/** Every element and attribute under a root, keyed by expanded names and positions among same-named siblings (so prefixes and sibling order do not matter). */
function inventory(root: Element): Map<string, Entry> {
  const out = new Map<string, Entry>();
  const visit = (el: Element, key: string, path: string): void => {
    for (const a of Array.from(el.attributes)) {
      if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) continue;
      out.set(`${key}/@${expanded(a)}`, { kind: 'attribute', name: a.name, path, node: el });
    }
    const counts = new Map<string, number>();
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const child = c as Element;
      const x = expanded(child);
      const n = (counts.get(x) ?? 0) + 1;
      counts.set(x, n);
      const childKey = `${key}/${x}[${n}]`;
      const childPath = `${path}/${child.nodeName}[${n}]`;
      out.set(childKey, { kind: 'element', name: child.nodeName, path, node: child });
      visit(child, childKey, childPath);
    }
  };
  visit(root, `/${expanded(root)}[1]`, `/${root.nodeName}[1]`);
  return out;
}

/** An element's line and column as xmldom recorded them. */
function lineOf(node: Node): { line?: number; column?: number } {
  const n = node as unknown as { lineNumber?: number; columnNumber?: number };
  return typeof n.lineNumber === 'number' ? { line: n.lineNumber, ...(typeof n.columnNumber === 'number' ? { column: n.columnNumber } : {}) } : {};
}

/** The element at an inventory key in another parse of the same text. */
function elementAt(root: Element, key: string): Element | undefined {
  const steps = key.split(/\/(?![^{]*\})/).filter((s) => s !== '' && !s.startsWith('@'));
  let el: Element | undefined = root;
  if (steps.length === 0 || `${expanded(root)}[1]` !== steps[0]) return undefined;
  for (const step of steps.slice(1)) {
    const m = /^(\{[^}]*\}[^[]+)\[(\d+)\]$/.exec(step);
    if (!m || !el) return undefined;
    let n = 0;
    let found: Element | undefined;
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && expanded(c as Element) === m[1] && ++n === Number(m[2])) { found = c as Element; break; }
    }
    el = found;
  }
  return el;
}

/**
 * The problems an Apply of `xml` to `part` would have; empty when it can go ahead. A byte part
 * needs only to parse to a document with a root element; a typed part is unmarshalled into a
 * scratch part of its class (with the package's preprocessor, as the part itself would be) and
 * marshalled back, and whatever the text has that the marshal lacks is a problem.
 */
export async function checkPartXml(part: Part, xml: string): Promise<PartXmlProblem[]> {
  const raw = new DefaultXmlPart(part.partName.toString());
  raw.setXml(xml);
  let doc: Document;
  try {
    // The engine's parse throws for a text that does not parse, a browser's `parsererror`
    // document included (core-ts 0.2.1, CR-002 section 33), with the parser's message.
    doc = await raw.getDocument();
  } catch (e) {
    const message = (e as Error).message;
    const where = locationIn(message);
    return [{ kind: 'parse', message, ...(where.line !== undefined ? where : await parseErrorAt(xml)) }];
  }
  const root = doc.documentElement;
  if (!root) return [{ kind: 'parse', message: 'no root element' }];
  if (partEditing(part).mode !== 'typed' || !(part instanceof XmlPart)) return [];
  const Ctor = part.constructor as new (name: string) => XmlPart<unknown>;
  const scratch = new Ctor(part.partName.toString());
  scratch.preprocessor = part.preprocessor;
  scratch.setXml(xml);
  try {
    await scratch.getContents();
  } catch (e) {
    if (e instanceof PartUnmarshalException) return [{ kind: 'value', message: e.message, ...(e.location ? { path: e.location } : {}), ...(e.line !== undefined ? { line: e.line } : {}), ...(e.column !== undefined ? { column: e.column } : {}) }];
    const message = (e as Error).message;
    return [{ kind: 'value', message, ...locationIn(message) }];
  }
  // What the engine read through and reported, in its words (core-ts 0.7.0): a value its type refuses, dropped, since Word
  // refuses a document holding it (CR-015 D8); a child its container never admits, which a save refuses (CR-016).
  const reported: PartXmlProblem[] = [];
  const refusedAttributes = new Set<string>();
  for (const w of scratch.warnings) {
    if (w.code !== 'value.refused' && w.code !== 'target.misplaced') continue;
    if (w.code === 'value.refused') refusedAttributes.add(w.attribute);
    reported.push({ kind: 'value', message: w.message, ...(w.line !== undefined ? { line: w.line } : {}), ...(w.column !== undefined ? { column: w.column } : {}) });
  }
  const back = await scratch.marshalToNode();
  const kept = inventory(back);
  const problems: { key: string; entry: Entry }[] = [];
  // A refused value's attribute is gone from the tree too: said once, in the engine's words.
  for (const [key, entry] of inventory(root)) if (!kept.has(key) && !(entry.kind === 'attribute' && refusedAttributes.has(entry.name))) problems.push({ key, entry });
  // A lost element takes its attributes and children with it: only the outermost loss is said.
  const lost = problems.filter((p) => !problems.some((q) => q !== p && q.entry.kind === 'element' && p.key.startsWith(`${q.key}/`)));
  return lost.length === 0 ? reported : [...reported, ...await locate(xml, root, lost)];
}

/** Lines for the problems: the engine's DOM when it records them (xmldom, in Node), else a second parse with xmldom, trusted only where it lands on the same name. */
async function locate(xml: string, root: Element, lost: { key: string; entry: Entry }[]): Promise<PartXmlProblem[]> {
  let other: Element | undefined;
  if (lineOf(root).line === undefined) {
    try {
      const { DOMParser } = await import('@xmldom/xmldom');
      other = (new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document).documentElement ?? undefined;
    } catch { other = undefined; }
  }
  return lost.map(({ key, entry }) => {
    const problem: PartXmlProblem = { kind: entry.kind, name: entry.name, path: entry.path };
    const elementKey = entry.kind === 'attribute' ? key.slice(0, key.lastIndexOf('/@')) : key;
    const own = elementAt(root, elementKey);
    let where = own ? lineOf(own) : {};
    if (where.line === undefined && other) {
      const theirs = elementAt(other, elementKey);
      const expect = entry.kind === 'element' ? entry.name : entry.node.nodeName;
      if (theirs && theirs.nodeName === expect) where = lineOf(theirs);
    }
    return { ...problem, ...where };
  });
}

/** A reference an Apply left without a definition (section 4.4: counted and said, not refused, since Word opens such a file). */
export interface DanglingReference {
  kind: 'paragraphStyle' | 'runStyle' | 'numbering';
  id: string;
  /** How many places name it. */
  count: number;
}

export interface AppliedPart {
  partName: string;
  /** What the part now leaves undefined that the document names. */
  dangling: DanglingReference[];
}

/** The styles or numbering definitions the document names and the parts lack. */
async function danglingAfter(session: EditorDocument, part: Part): Promise<DanglingReference[]> {
  const doc = session.doc;
  const counts = new Map<string, DanglingReference>();
  const add = (kind: DanglingReference['kind'], id: string): void => {
    const k = `${kind}:${id}`;
    const r = counts.get(k);
    if (r) r.count++; else counts.set(k, { kind, id, count: 1 });
  };
  if (part instanceof StyleDefinitionsPart) {
    const styles = (await part.getContents()) as wml.Styles;
    const defined = new Set((styles.style ?? []).map((s) => s.styleId ?? ''));
    doc.descendants((node) => {
      const p = node.attrs['pStyle'];
      if (node.type.name === 'paragraph' && typeof p === 'string' && !defined.has(p)) add('paragraphStyle', p);
      for (const m of node.marks) if (m.type.name === 'runStyle' && typeof m.attrs['id'] === 'string' && !defined.has(m.attrs['id'] as string)) add('runStyle', m.attrs['id'] as string);
      return true;
    });
  } else if (part instanceof NumberingDefinitionsPart) {
    const numbering = (await part.getContents()) as wml.Numbering;
    const defined = new Set((numbering.num ?? []).map((n) => String(n.numId)));
    doc.descendants((node) => {
      const id = (node.attrs['numPr'] as { numId: string | null } | null | undefined)?.numId;
      if (node.type.name === 'paragraph' && id && id !== '0' && !defined.has(id)) add('numbering', id);
      return true;
    });
  }
  return [...counts.values()];
}

/**
 * Apply checked text to a part (section 4.3): a typed part set and unmarshalled at once, so what
 * the session reads from it is the new text; a custom XML data part through its view, so the
 * Template perspective's tabs and verdicts see it; the part marked as edited in place, which
 * writes it into the store and so to the room (ED-004 section 3.3: styles and numbering by
 * element, the rest whole). Then what the session reads from the part is read again (section
 * 4.4): the resolver for styles, the numbering definitions, the settings. The view's stylesheet
 * and labels are the caller's (`DocxEditorView.partsChanged`). Call `checkPartXml` first.
 */
export async function applyPartXml(session: EditorDocument, part: Part, xml: string): Promise<AppliedPart> {
  const partName = part.partName.toString();
  const editing = partEditing(part);
  if (editing.mode === 'none') throw new Error(`${partName} is not edited here (${editing.reason})`);
  if (part instanceof XmlPart) {
    part.setXml(xml);
    await part.getContents();
  } else if (part instanceof CustomXmlDataStoragePart) {
    const view = session.pkg.customXmlParts.items.find((p) => p.part === part);
    if (view) view.setXml(xml); else part.setXml(xml);
  } else if (part instanceof DefaultXmlPart) {
    part.setXml(xml);
  }
  session.markPartChanged(partName);
  if (part instanceof StyleDefinitionsPart) session.pkg.refreshPropertyResolver();
  else if (part instanceof NumberingDefinitionsPart || part instanceof ThemePart || part instanceof FontTablePart) await session.pkg.refresh();
  else if (part instanceof DocumentSettingsPart) await session.readSettings();
  if (part instanceof CustomXmlDataStoragePart || partName.startsWith('/customXml/')) session.template.invalidate();
  return { partName, dangling: await danglingAfter(session, part) };
}
