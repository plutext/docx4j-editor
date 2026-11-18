// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word's XML mapping, both ways (ED-005 sections 12.54 and 12.55). A content control with a
// `w:dataBinding` shows its node's value, and the node holds the control's value:
//
// - **bound** (`writeBoundValue`): binding a control puts its node's value in it at once;
// - **control to data** (`writeBackBoundControls`): a local edit inside a bound control writes the
//   control's value to its node, an undo or a redo included;
// - **data to controls** (`refreshBoundControls`): a node's new value is put into every control bound
//   to it whose content says otherwise, as one local change kept out of the undo history (a part edit
//   has no undo, section 9 item 7 of ED-003), untracked.
//
// Only local changes act: a peer's typing in a bound control comes with the data part its page wrote,
// and refreshing on a peer's part edit would make each page write the same text, twice over in the
// merge. Nothing walks the document per keystroke: a transaction's bound controls are found from its
// changed positions, and the document's bound controls are memoised per top-level block.
import type { Node as PMNode, Mark } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { DEFAULT_ENTRIES, DEFAULT_PLACEHOLDER_TEXT, defaultEntryFor, type CustomXmlNode, type DefaultEntry } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { EditorDocument } from './session.mjs';
import { applyControlProperties, readControl, sdtBinding, sdtId, sdtItem, sdtKind, type DataBinding, type SdtKind, type SdtPrJson } from './sdtpr.mjs';
import { isDeleted } from './revisions.mjs';
import { UNTRACKED_META, markDeletedRange, markInsertedRange, type Tracking } from './track.mjs';

/** On a transaction that wrote a bound value into a control (a bind, a refresh): its content came from the node, so it writes nothing back. */
export const BOUND_WRITE_META = 'docx4j-bound-write';

/** Word's default placeholder text for a text control (the engine's `DEFAULT_PLACEHOLDER_TEXT`), and the style every placeholder run carries (`PLACEHOLDER_STYLE`). */
export const PLACEHOLDER_TEXT = DEFAULT_PLACEHOLDER_TEXT;
export const PLACEHOLDER_STYLE = 'PlaceholderText';

/** A default placeholder entry of Word's glossary: its name and its text (the engine's `DefaultEntry`). */
export type DefaultPlaceholder = DefaultEntry;

/**
 * Word's default placeholder entry for each kind of control that has one, as Word 365 makes them (core-ts CR-002
 * section 63.22, Jason's `placeholders.docx`): a picture, a check box and a repeating section have none. The engine's
 * `DEFAULT_ENTRIES`, exported since core-ts 0.7.0 (UPSTREAM-24, closed for the entries).
 */
export const DEFAULT_PLACEHOLDERS = DEFAULT_ENTRIES;

/** The default entry Word gives a control of this kind (the engine's `defaultEntryFor`); undefined for a kind it gives none. */
export function defaultPlaceholderFor(sdtPr: unknown): DefaultPlaceholder | undefined {
  return defaultEntryFor(sdtPr as wml.SdtPr | undefined);
}

/**
 * A control made to show its placeholder, in Word's form (ED-008 section 8.8 item 14, as the engine's
 * `showPlaceholder` writes one): its properties with `w:showingPlcHdr` and, for a kind Word gives a default
 * entry, `w:placeholder` naming that entry, whose text the placeholder run shows; the save writes the entry
 * into the glossary where the file lacks it (`EditorDocument.ensurePlaceholders`). A kind with no entry
 * shows `PLACEHOLDER_TEXT` with no reference, as the engine does.
 */
export function placeholderFor(sdtPr: unknown): { sdtPr: SdtPrJson; text: string } {
  const entry = defaultPlaceholderFor(sdtPr);
  return {
    sdtPr: applyControlProperties(sdtPr, { showingPlaceholder: true, ...(entry ? { placeholderEntry: entry.name } : {}) }),
    text: entry?.text ?? PLACEHOLDER_TEXT,
  };
}

/** The default entry a control's `w:placeholder` names, where it names one (the save's glossary write). */
export function namedDefaultPlaceholder(sdtPr: unknown): DefaultPlaceholder | undefined {
  const name = readControl(sdtPr).placeholderEntry;
  return name ? Object.values(DEFAULT_PLACEHOLDERS).find((e) => e.name === name) : undefined;
}

/** The kinds whose content is the bound value's text, as Word shows it. */
const VALUE_KINDS: ReadonlySet<SdtKind> = new Set<SdtKind>(['PlainText', 'RichText', 'ComboBox', 'DropDownList']);

/** The kinds whose value is written back to their node: the text kinds whose content is the value itself (rich text Word writes as WordprocessingML, ED-005 section 12.55). */
const WRITTEN_BACK: ReadonlySet<SdtKind> = new Set<SdtKind>(['PlainText', 'ComboBox', 'DropDownList']);

/**
 * A control's kind as Word's mapping takes it (core-ts check 35, ED-005 section 12.58): a rich text
 * control bound with `w:dataBinding` is plain text to Word, which shows and writes its node's text
 * and adds `<w:text/>` on saving it (1f); Word binds rich text with `w15:dataBinding` (1a).
 */
export function mappedKind(sdtPr: unknown): SdtKind {
  const kind = sdtKind(sdtPr);
  return kind === 'RichText' && sdtBinding(sdtPr)?.w15 === false ? 'PlainText' : kind;
}

/**
 * Whether a node's value is a Flat OPC package (`pkg:package`), as Word writes a mapped rich text
 * control's content into its node (core-ts check 35, 1e): the editor neither shows it as text nor
 * writes over it.
 */
export function isFlatOpc(value: string): boolean {
  return /^\s*(?:<\?[^>]*\?>\s*)*<pkg:package[\s>]/.test(value.slice(0, 1000));
}

/** The items a control's old content keeps, ahead of the bound value: markers with no text of their own, which a bound value does not replace. */
const KEPT_MARKERS: ReadonlySet<string> = new Set(['bookmarkStart', 'bookmarkEnd', 'permStart', 'permEnd', 'commentReference']);

/** Whether a mark is formatting a bound value keeps: not a revision, a comment or the placeholder's style. */
function isFormattingMark(m: Mark): boolean {
  return !['insertion', 'deletion', 'formatChange', 'comment'].includes(m.type.name) && !(m.type.name === 'runStyle' && m.attrs['id'] === PLACEHOLDER_STYLE);
}

/** The marks a control's new text keeps from its old. */
function formattingMarks(node: PMNode | undefined): readonly Mark[] {
  return node ? node.marks.filter(isFormattingMark) : [];
}

type ListItem = { value?: string; displayText?: string };

function listItemsOf(sdtPr: unknown, kind: SdtKind): ListItem[] {
  const list = sdtItem(sdtPr, kind === 'ComboBox' ? 'comboBox' : 'dropDownList')?.value as { listItem?: ListItem[] } | undefined;
  return list?.listItem ?? [];
}

/** A drop-down or combo box's display text for a value: its list item's, else the value itself (Word shows the value then). */
function displayTextOf(sdtPr: unknown, kind: SdtKind, value: string): string {
  const item = listItemsOf(sdtPr, kind).find((i) => (i.value ?? i.displayText) === value);
  return item?.displayText ?? item?.value ?? value;
}

/** What a control would show for a value: its text, a list item's display text. */
function shownFor(sdtPr: unknown, kind: SdtKind, value: string): string {
  return kind === 'ComboBox' || kind === 'DropDownList' ? displayTextOf(sdtPr, kind, value) : value;
}

/** The paragraph a control's text is in: a run control itself, or a block control's first paragraph. */
function textHolderOf(control: PMNode): PMNode | undefined {
  if (control.type.name === 'sdtInline') return control;
  const first = control.firstChild;
  return first?.type.name === 'paragraph' ? first : undefined;
}

/** A control's text as its content reads in the accepted view: a line break as a newline, deleted text left out. */
function controlText(control: PMNode): string {
  const holder = textHolderOf(control);
  if (!holder) return '';
  let out = '';
  holder.descendants((n) => {
    if (n.isText) { if (!isDeleted(n)) out += n.text; return false; }
    if (n.type.name === 'hardBreak' && n.attrs['kind'] === 'line' && !isDeleted(n)) out += '\n';
    return true;
  });
  return out;
}

/**
 * The bound value written into a block or run control (the node's value as it is, as Word's XML mapping
 * shows it; the Verdict trims, as an OpenDoPE processor would):
 * a plain or rich text control's text, a drop-down or combo box's display text. The text takes the
 * formatting of the control's first text (Word keeps it), with a line break for each line of a
 * multi-line plain text control. An empty value shows the placeholder, as Word does. A date, a check
 * box, a picture and a section are left as they are. A bookmark's or a permission range's markers and
 * a comment's reference inside stay, ahead of the value, never dropped. A caret in the control stays
 * in it, after the value. The control's properties, `showingPlcHdr` set or cleared to match;
 * returned for the caller to write.
 *
 * With `tracking` (a bind over text that was not bound, ED-005 section 12.57): the old content stays,
 * marked deleted, and the new is inserted after it, marked, as Word's mapping writes it with tracking
 * on (core-ts check 34, 1j); a placeholder goes for real, not being content. The transaction is then
 * marked `UNTRACKED_META`, so the tracker does not track it twice. A block control whose last block
 * is not a paragraph is written untracked.
 */
export function writeBoundValue(tr: Transaction, pos: number, sdtPr: unknown, value: string | undefined, fallbackMarks: readonly Mark[] = [], tracking?: Tracking): unknown {
  const kind = mappedKind(sdtPr);
  if (!VALUE_KINDS.has(kind) || value === undefined || isFlatOpc(value)) return sdtPr;
  tr.setMeta(BOUND_WRITE_META, true);
  const schema = tr.doc.type.schema;
  const node = tr.doc.nodeAt(pos);
  if (!node) return sdtPr;
  // Where the text goes: a run control's content, or a block control's first paragraph (the rest go).
  const block = node.type.name === 'sdtBlock';
  const paragraph = textHolderOf(node);
  if (!paragraph) return sdtPr;
  const placeholder = readControl(sdtPr).showingPlaceholder;
  let first: PMNode | undefined;
  paragraph.descendants((n) => { if (!first && n.isText) first = n; return !first; });
  let content: PMNode[];
  let shownSdtPr: SdtPrJson | undefined;
  if (value === '') {
    const shown = placeholderFor(sdtPr);
    shownSdtPr = shown.sdtPr;
    content = [schema.text(shown.text, [schema.marks['runStyle']!.create({ id: PLACEHOLDER_STYLE })])];
  } else {
    // The formatting of the control's first text; an empty control takes what its maker gives (a caret's marks).
    const marks = placeholder ? [] : first ? formattingMarks(first) : fallbackMarks.filter(isFormattingMark);
    const text = shownFor(sdtPr, kind, value);
    const multiLine = kind === 'PlainText' && (sdtItem(sdtPr, 'text')?.value as { multiLine?: boolean } | undefined)?.multiLine === true;
    const lines = multiLine ? text.split(/\r\n|\r|\n/) : [text.replace(/\s*[\r\n]+\s*/g, ' ')];
    content = [];
    lines.forEach((line, i) => {
      if (i > 0) content.push(schema.nodes['hardBreak']!.create({ kind: 'line' }));
      if (line !== '') content.push(schema.text(line, marks));
    });
  }
  // A caret in the control stays in it, after the value, so the panels still show this control.
  const inside = tr.selection.from > pos && tr.selection.to < pos + node.nodeSize;
  // Tracked: the value after the old content, which stays, marked deleted (its markers stay where they are).
  const end = block ? (node.lastChild?.type.name === 'paragraph' ? pos + node.nodeSize - 2 : undefined) : pos + node.nodeSize - 1;
  if (tracking && !placeholder && end !== undefined) {
    const size = content.reduce((n, c) => n + c.nodeSize, 0);
    tr.insert(end, content);
    // The placeholder is no content: marked only where it replaces text, so that a reject takes it away.
    if (value !== '' || controlText(node) !== '') markInsertedRange(tr, end, end + size, tracking);
    const steps = tr.steps.length;
    markDeletedRange(tr, block ? pos + 2 : pos + 1, end, tracking);
    if (inside) tr.setSelection(TextSelection.create(tr.doc, tr.mapping.slice(steps).map(end + size)));
    tr.setMeta(UNTRACKED_META, true);
    return shownSdtPr ?? applyControlProperties(sdtPr, { showingPlaceholder: false });
  }
  // A bookmark's or a permission range's markers, and a comment's reference, stay, ahead of the value.
  const kept: PMNode[] = [];
  paragraph.descendants((n) => { if (n.type.name === 'opaqueInline' && KEPT_MARKERS.has(n.attrs['kind'] as string)) kept.push(n); return true; });
  const keptSize = kept.reduce((n, c) => n + c.nodeSize, 0);
  content = [...kept, ...content];
  const size = content.reduce((n, c) => n + c.nodeSize, 0);
  const start = block ? pos + 2 : pos + 1;
  if (block) tr.replaceWith(pos + 1, pos + node.nodeSize - 1, paragraph.type.create(paragraph.attrs, content, paragraph.marks));
  else tr.replaceWith(pos + 1, pos + node.nodeSize - 1, content);
  if (inside) tr.setSelection(TextSelection.create(tr.doc, start + size));
  // Tracked over a placeholder: the value is inserted, the placeholder having been no content.
  if (tracking && placeholder && value !== '') {
    markInsertedRange(tr, start + keptSize, start + size, tracking);
    tr.setMeta(UNTRACKED_META, true);
  }
  return shownSdtPr ?? applyControlProperties(sdtPr, { showingPlaceholder: false });
}

/** A block or run control with a `w:dataBinding`, at its position in a document. */
export interface BoundControl {
  pos: number;
  node: PMNode;
  binding: DataBinding;
}

const isControl = (n: PMNode): boolean => n.type.name === 'sdtInline' || n.type.name === 'sdtBlock';

const blockControls = new WeakMap<PMNode, { offset: number; node: PMNode; binding: DataBinding }[]>();

/** A top-level block's bound controls, their offsets within it. Memoised per block. */
function boundIn(block: PMNode): { offset: number; node: PMNode; binding: DataBinding }[] {
  let out = blockControls.get(block);
  if (out) return out;
  const found: { offset: number; node: PMNode; binding: DataBinding }[] = [];
  const visit = (node: PMNode, offset: number): void => {
    if (isControl(node)) {
      const binding = sdtBinding(node.attrs['sdtPr']);
      if (binding && binding.xpath !== '') found.push({ offset, node, binding });
    }
    node.forEach((child, childOffset) => { if (!child.isText) visit(child, offset + 1 + childOffset); });
  };
  visit(block, 0);
  out = found;
  blockControls.set(block, out);
  return out;
}

/** Every bound control in a document, in document order. */
export function boundControlsIn(doc: PMNode): BoundControl[] {
  const out: BoundControl[] = [];
  doc.forEach((block, pos) => { for (const c of boundIn(block)) out.push({ pos: pos + c.offset, node: c.node, binding: c.binding }); });
  return out;
}

/**
 * The bound controls a transaction may have changed, in its document: the innermost control around
 * each changed position, and every control inside a changed range (an undo, a remote update or a
 * paste replaces whole blocks). Whether one's value changed is the caller's question.
 */
export function boundControlsTouched(tr: Transaction): BoundControl[] {
  const seen = new Map<number, BoundControl>();
  const within = (from: number, to: number): void => {
    if (to <= from) return;
    tr.doc.nodesBetween(Math.max(0, from), Math.min(to, tr.doc.content.size), (node, pos) => {
      if (node.isText) return false;
      if (isControl(node)) {
        const binding = sdtBinding(node.attrs['sdtPr']);
        if (binding && binding.xpath !== '') seen.set(pos, { pos, node, binding });
      }
      return true;
    });
  };
  const at = (pos: number): void => {
    const $pos = tr.doc.resolve(Math.max(0, Math.min(pos, tr.doc.content.size)));
    for (let d = $pos.depth; d >= 1; d--) {
      const node = $pos.node(d);
      if (!isControl(node)) continue;
      const binding = sdtBinding(node.attrs['sdtPr']);
      if (binding && binding.xpath !== '') seen.set($pos.before(d), { pos: $pos.before(d), node, binding });
      return;
    }
  };
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1);
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      const from = rest.map(newStart, -1);
      const to = rest.map(newEnd, 1);
      at(from);
      at(to);
      within(from, to);
    });
  });
  return [...seen.values()];
}

/**
 * What a control's content says, as the node it is bound to would hold it: a plain text control's
 * text (a line break a newline), a drop-down's or combo box's chosen item's value (its text when it
 * names no item). Undefined for a kind not written back (rich text, a date, a check box, a picture)
 * and for a control showing its placeholder.
 */
export function controlValue(control: PMNode): string | undefined {
  const sdtPr = control.attrs['sdtPr'];
  const kind = mappedKind(sdtPr);
  if (!WRITTEN_BACK.has(kind) || readControl(sdtPr).showingPlaceholder) return undefined;
  const text = controlText(control);
  if (kind === 'PlainText') return text;
  const item = listItemsOf(sdtPr, kind).find((i) => (i.displayText ?? i.value) === text);
  return item?.value ?? item?.displayText ?? text;
}

/** The node a binding selects in its part, or undefined (no such part, an XPath that selects nothing or fails). */
export function boundNodeOf(session: EditorDocument, binding: DataBinding): CustomXmlNode | undefined {
  const part = session.template.partNamed(binding.storeItemID);
  if (!part) return undefined;
  try {
    return part.selectSingleNode(binding.xpath, binding.prefixMappings || undefined) ?? undefined;
  } catch {
    return undefined;
  }
}

/**
 * Word's mapping, the control's side: each control's value written to its node where the two differ.
 * The parts are marked changed (no undo: a part edit has none; an undo of the edit writes the old
 * value back, being an edit inside the control itself) and the template view forgets its reads. The
 * DOM nodes written, for `refreshBoundControls`. The template view must be loaded.
 */
export function writeBackBoundControls(session: EditorDocument, controls: readonly BoundControl[]): Set<Node> {
  const written = new Set<Node>();
  for (const c of controls) {
    const value = controlValue(c.node);
    if (value === undefined) continue;
    const node = boundNodeOf(session, c.binding);
    // A node holding a rich text control's WordprocessingML is not written over: Word refuses to map it to a plain text control (core-ts check 35, 1e).
    if (!node || node.nodeValue === value || isFlatOpc(node.nodeValue)) continue;
    node.nodeValue = value;
    node.ownerPart.touch();
    session.markPartChanged(node.ownerPart.part.partName.toString());
    written.add(node.node);
  }
  if (written.size > 0) session.template.invalidate();
  return written;
}

/**
 * Word's mapping, the data's side: a transaction putting each bound control's node value into it where
 * its content says otherwise; every bound control, or those whose node is one of `only`, never one of
 * `except` (the controls whose edit wrote the node). Untracked and kept out of the undo history (apply
 * it with `applyBoundRefresh`). Undefined when no control differs. The template view must be loaded.
 */
export function refreshBoundControls(session: EditorDocument, state: EditorState, only?: ReadonlySet<Node>, except?: ReadonlySet<number>): Transaction | undefined {
  const tr = state.tr;
  const controls = boundControlsIn(state.doc).filter((c) => VALUE_KINDS.has(mappedKind(c.node.attrs['sdtPr'])));
  // From the last to the first, so that a control written does not move those before it.
  for (const c of controls.reverse()) {
    const sdtPr = c.node.attrs['sdtPr'];
    const id = sdtId(sdtPr);
    if (except && id !== undefined && except.has(id)) continue;
    const node = boundNodeOf(session, c.binding);
    if (!node || (only && !only.has(node.node))) continue;
    const value = node.nodeValue;
    const showing = readControl(sdtPr).showingPlaceholder;
    if (isFlatOpc(value)) continue;
    if (value === '' ? showing : !showing && controlText(c.node) === shownFor(sdtPr, mappedKind(sdtPr), value)) continue;
    const next = writeBoundValue(tr, c.pos, sdtPr, value);
    tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, sdtPr: next });
  }
  return tr.docChanged ? tr.setMeta(UNTRACKED_META, true).setMeta('addToHistory', false) : undefined;
}

/**
 * Refresh the bound controls (all, or those bound to `only`) and write the result to the store as one
 * local change out of the undo history (`session.replaceOutOfHistory`): a view sees it as it sees any
 * change to the store, without its own transaction, so typing's undo steps stay as they were. Whether
 * anything changed. `state` is the document the store holds (a view's, or a runner's last state).
 */
export function applyBoundRefresh(session: EditorDocument, state: EditorState, only?: ReadonlySet<Node>, except?: ReadonlySet<number>): boolean {
  const tr = refreshBoundControls(session, state, only, except);
  if (!tr) return false;
  session.replaceOutOfHistory(tr.doc, state.doc);
  return true;
}

/** The `w:id`s of controls, to find them again in a later document (a write-back after the template view loads). */
export function controlIds(controls: readonly BoundControl[]): Set<number> {
  const out = new Set<number>();
  for (const c of controls) { const id = sdtId(c.node.attrs['sdtPr']); if (id !== undefined) out.add(id); }
  return out;
}

/** The bound controls of a document with one of these `w:id`s. */
export function boundControlsWithIds(doc: PMNode, ids: ReadonlySet<number>): BoundControl[] {
  return boundControlsIn(doc).filter((c) => { const id = sdtId(c.node.attrs['sdtPr']); return id !== undefined && ids.has(id); });
}

/**
 * After a local transaction applied to the store (an edit, or an undo or a redo in a view): the bound
 * controls it changed something inside write their values to their nodes, and the controls bound to
 * the nodes written are refreshed (`schedule` runs that, after a view's dispatch). `current` gives
 * the state the store holds then. With the template view not loaded, it loads, then writes back what
 * those controls hold by then (found by their `w:id`s).
 */
export function syncBoundAfter(session: EditorDocument, tr: Transaction, current: () => EditorState, schedule: (fn: () => void) => void = (fn) => fn()): void {
  if (!tr.docChanged || tr.getMeta(BOUND_WRITE_META)) return;
  // Only the controls whose value this transaction changed: an undo replaces whole blocks, and a control beside the one
  // undone, still showing the node's value, must not write it back over the undo.
  const touched = boundControlsTouched(tr);
  if (touched.length === 0) return;
  const before = new Map<number, string | undefined>();
  for (const c of boundControlsWithIds(tr.before, controlIds(touched))) before.set(sdtId(c.node.attrs['sdtPr'])!, controlValue(c.node));
  const changed = touched.filter((c) => {
    const id = sdtId(c.node.attrs['sdtPr']);
    return id === undefined || !before.has(id) || before.get(id) !== controlValue(c.node);
  });
  if (changed.length === 0) return;
  const run = (controls: readonly BoundControl[]): void => {
    const written = writeBackBoundControls(session, controls);
    // The controls bound to the nodes written follow; the ones whose edit wrote them are left as typed.
    if (written.size > 0) schedule(() => { applyBoundRefresh(session, current(), written, controlIds(controls)); });
  };
  if (session.template.loaded) { run(changed); return; }
  const ids = controlIds(changed);
  void session.template.load().then(() => run(boundControlsWithIds(current().doc, ids))).catch(() => undefined);
}
