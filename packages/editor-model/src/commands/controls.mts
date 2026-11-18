// SPDX-License-Identifier: GPL-3.0-or-later
//
// Content controls (ED-003 section 4.3): insert, wrap rows and cells, delete, properties, bind,
// and the OpenDoPE roles, in content-API shapes. Block and run controls are nodes; row and cell
// groups are attribute stacks whose properties live in the `controls` map, written through the
// transaction's control writes so that a command is one undo step (yjs.mts). The part writes an
// OpenDoPE command makes (an XPaths entry, a condition) go through the template view in the same
// command and are not undone (section 9 item 7): the panel says so beside its Apply.
import { TextSelection, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import { Fragment, type Node as PMNode, type Attrs } from 'prosemirror-model';
import { isInTable, selectedRect, CellSelection, TableMap } from 'prosemirror-tables';
import { find } from '@docx4j/generated-objects-ts/builders/wml';
import type { SdtGroup } from '../import.mjs';
import type { EditorDocument } from '../session.mjs';
import { controlWritesOf, addControlWrite, readControlEntries } from '../yjs.mjs';
import { controlAt, sdtPrOf, rowExtent, cellExtent, rowPosition, isNodeTarget, type ControlTarget, type NodeTarget } from '../controls.mjs';
import { newSdtPr, allocateSdtId, applyControlProperties, withBinding, withoutSdtItem, sdtId, sdtTag, sdtKind, sdtBinding, sdtItem, readControl, type SdtKind, type ControlProperties } from '../sdtpr.mjs';
import { writeBoundValue, boundNodeOf, boundControlsIn, mappedKind, placeholderFor, PLACEHOLDER_TEXT, PLACEHOLDER_STYLE } from '../bound.mjs';
import { PICTURE_PLACEHOLDER_BASE64, PICTURE_PLACEHOLDER_EMU, type CustomXmlNode } from '@docx4j/core-ts';
import { relativeToRepeat, parseTag, formatTag, withTagParam, readTag, OpenDopeError, type Operand } from '../opendope.mjs';
import { XHTML_CONTENT_TYPE, xhtmlKindOf } from '../xhtml.mjs';
import { looksLikeImageBase64, pictureFactsOf } from '../pictures.mjs';
import { WORD_DOCUMENT_PROGID } from '../fragments.mjs';
import { isFlatOpc } from '../bound.mjs';
import { register, sessionOf, storyOf } from './spec.mjs';
import { blocksEnd } from './paragraph.mjs';
import type { Tracking } from '../track.mjs';
import { normalizeRowGroups } from './tables.mjs';

export { PLACEHOLDER_TEXT, PLACEHOLDER_STYLE } from '../bound.mjs';

/** Every `w:id` in use: the tree's, the document's nodes', the groups' (ED-003 section 4.2: never collide with a control the editor holds). */
export function usedSdtIds(session: EditorDocument | undefined, doc: PMNode): Set<number> {
  const used = new Set<number>();
  const add = (sdtPr: unknown): void => { const id = sdtId(sdtPr); if (id !== undefined) used.add(id); };
  if (session) {
    for (const pr of find(session.body, 'org_docx4j_wml.SdtPr')) add(pr);
    for (const entry of readControlEntries(session.store).values()) add(entry.sdtPr);
  }
  doc.descendants((n) => { if (n.type.name === 'sdtBlock' || n.type.name === 'sdtInline') add(n.attrs['sdtPr']); return true; });
  return used;
}

/** Points to EMU: 12,700 to the point. */
const EMU_PER_POINT = 12700;

function newGroupId(): string {
  return `new:${Math.random().toString(36).slice(2, 10)}`;
}

/** Write a target's properties: the node's attribute, or the group's map entry. */
function writeSdtPr(tr: Transaction, target: ControlTarget, sdtPr: unknown): void {
  if (isNodeTarget(target)) {
    tr.setNodeMarkup(target.pos, undefined, { ...target.node.attrs, sdtPr });
    return;
  }
  addControlWrite(tr, { groupId: target.groupId, entry: { form: target.form, sdtPr } });
}

export interface InsertControlArgs {
  /** `RichText` (the default), `PlainText`, `CheckBox`, `DatePicker`, `DropDownList`, `ComboBox`, `Picture`, `RepeatingSection`, `Group`. */
  kind?: SdtKind;
  tag?: string;
  title?: string;
}

/**
 * Whether the selection lies inside a control that cannot hold a control (a plain text control, a check box, a date
 * picker, a drop-down, a picture): Word refuses a control there, and so does the editor (ED-005 section 12.85).
 */
export function insideControlHoldingNoControls(state: EditorState): boolean {
  const { $from, $to } = state.selection;
  for (const $p of [$from, $to]) {
    for (let d = $p.depth; d > 0; d--) {
      const n = $p.node(d);
      if ((n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') && !HOLDS_CONTROLS.has(readControl(n.attrs['sdtPr']).kind)) return true;   // HOLDS_CONTROLS: below, the Data tab's rule
    }
  }
  return false;
}

register<InsertControlArgs | undefined>({
  name: 'contentControl.insert',
  tracking: 'untracked',
  label: 'Insert content control',
  choices: { kind: ['RichText', 'PlainText', 'CheckBox', 'DatePicker', 'DropDownList', 'ComboBox', 'Picture', 'RepeatingSection', 'Group'] },
  args: { kind: "'RichText' (default) | 'PlainText' | 'CheckBox' | 'DatePicker' | 'DropDownList' | 'ComboBox' | 'Picture' | 'RepeatingSection' | 'Group'", tag: 'w:tag', title: 'w:alias' },
  run: ({ kind = 'RichText', tag, title } = {}) => (state, dispatch) => {
    if (state.selection instanceof CellSelection || insideControlHoldingNoControls(state)) return false;
    const session = sessionOf(state);
    const schema = state.schema;
    const { $from, $to, empty } = state.selection;
    const used = usedSdtIds(session, state.doc);
    const id = allocateSdtId(used);
    const options = { id, ...(tag !== undefined ? { tag } : {}), ...(title !== undefined ? { title } : {}) };
    let sdtPr = newSdtPr(kind, options);
    const blockOnly = kind === 'RepeatingSection' || kind === 'Group';
    const inParagraph = $from.parent.type.name === 'paragraph' || $from.parent.type.name === 'sdtInline';
    const wholeBlock = !empty && $from.parentOffset === 0 && $to.parentOffset === $to.parent.content.size && $from.parent.content.size > 0;
    const inline = !blockOnly && inParagraph && $from.sameParent($to) && !wholeBlock;
    const tr = state.tr;
    if (inline) {
      // A run-level control around the selection; at a caret, one showing Word's placeholder.
      let content: Fragment;
      if (empty && kind === 'Picture' && session && storyOf(state) === undefined) {
        // Word's form for a picture control (ED-005 section 12.126): the flag and the ports' placeholder picture, 100 points
        // square (core-ts CR-005 section 10.2), which Word fills from a mapped node by swapping the picture; it never fills
        // placeholder text (section 12.125). The image part is a part write, made only when the command applies, as the
        // bind's XPaths entry is, and not undone. In a note or a header the image would need a relationship from its part,
        // so the control there keeps the text placeholder.
        if (!dispatch) return true;
        const made = session.addInlinePicture(PICTURE_PLACEHOLDER_BASE64, { width: PICTURE_PLACEHOLDER_EMU / EMU_PER_POINT });
        content = Fragment.from(schema.nodes['image']!.create({ sourceId: made.sourceId, width: made.width, height: made.height }));
        sdtPr = applyControlProperties(sdtPr, { showingPlaceholder: true });
      } else if (empty) {
        const shown = placeholderFor(sdtPr);
        content = Fragment.from(schema.text(shown.text, [schema.marks['runStyle']!.create({ id: PLACEHOLDER_STYLE })]));
        sdtPr = shown.sdtPr;
      } else {
        content = state.doc.slice($from.pos, $to.pos).content;
      }
      const node = schema.nodes['sdtInline']!.create({ sourceId: null, sdtPr }, content);
      tr.replaceWith($from.pos, $to.pos, node);
      tr.setSelection(TextSelection.create(tr.doc, $from.pos + 1, $from.pos + 1 + content.size));
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    // Block-level: around the whole blocks the selection touches, inside the enclosing control or cell; a selection
    // ending at the next paragraph's start (the selection bar's) does not touch that paragraph (blocksEnd).
    const range = $from.blockRange(state.doc.resolve(blocksEnd(state)));
    if (!range) return false;
    const parent = range.parent.type.name;
    if (parent !== 'doc' && parent !== 'sdtBlock' && parent !== 'table_cell') return false;
    const sdtBlock = schema.nodes['sdtBlock']!;
    const wrappers: { type: typeof sdtBlock; attrs: Attrs }[] = [{ type: sdtBlock, attrs: { sourceId: null, sdtPr } }];
    if (kind === 'RepeatingSection') {
      // Word's shape: the section holds one repeating section item, which holds the content.
      used.add(id);
      wrappers.push({ type: sdtBlock, attrs: { sourceId: null, sdtPr: newSdtPr('RepeatingSectionItem', { id: allocateSdtId(used) }) } });
    }
    tr.wrap(range, wrappers);
    keepParagraphAfter(tr, range.end + 2 * wrappers.length);
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: (state) => !(state.selection instanceof CellSelection) && !insideControlHoldingNoControls(state),
});

/**
 * Word never puts the final paragraph mark of a document, a cell or a block control's content inside
 * another content control: a block control made around the last block of any of them gets an empty
 * paragraph after it, so the caret can leave the control and typing after it is possible (Jason,
 * 2026-10-08: a control bound into the empty last paragraph held every keystroke). `end` is the
 * position just after the new control.
 */
function keepParagraphAfter(tr: Transaction, end: number): void {
  const $end = tr.doc.resolve(end);
  const holder = $end.parent.type.name;
  if ((holder !== 'doc' && holder !== 'table_cell' && holder !== 'sdtBlock') || $end.parentOffset !== $end.parent.content.size) return;
  const paragraph = tr.doc.type.schema.nodes['paragraph']!.createAndFill();
  if (paragraph) tr.insert(end, paragraph);
}

/** The group ids on the selected rows or cells, each entirely inside or entirely around the selection; undefined when one crosses it. */
function outerCount(stacks: SdtGroup[][], extentOf: (id: string) => [number, number] | undefined, first: number, last: number): number[] | undefined {
  const counts: number[] = [];
  for (const stack of stacks) {
    let outer = 0;
    let sawInner = false;
    for (const g of stack) {
      const e = extentOf(g.groupId);
      if (!e) return undefined;
      const isOuter = e[0] <= first && e[1] >= last;
      const isInner = e[0] >= first && e[1] <= last;
      if (!isOuter && !isInner) return undefined;
      if (isOuter && !sawInner) outer++;
      else sawInner = true;
    }
    counts.push(outer);
  }
  return counts;
}

register<InsertControlArgs | undefined>({
  name: 'contentControl.wrapRows',
  tracking: 'untracked',
  label: 'Wrap rows in content control',
  args: { kind: "'RichText' (default) or 'RepeatingSection'", tag: 'w:tag', title: 'w:alias' },
  choices: { kind: ['RichText', 'RepeatingSection'] },
  run: ({ kind = 'RichText', tag, title } = {}) => (state, dispatch) => {
    if (!isInTable(state)) return false;
    const session = sessionOf(state);
    const rect = selectedRect(state);
    const table = rect.table;
    const tablePos = rect.tableStart - 1;
    const top = rect.top;
    const bottom = rect.bottom - 1;
    const stacks: SdtGroup[][] = [];
    for (let r = top; r <= bottom; r++) stacks.push(((table.child(r).attrs['sdt'] as SdtGroup[] | null) ?? []).map((g) => ({ ...g })));
    const depths = outerCount(stacks, (id) => rowExtent(table, id), top, bottom);
    if (!depths) return false;
    const used = usedSdtIds(session, state.doc);
    const groups: { groupId: string; sdtPr: unknown }[] = [];
    const id = allocateSdtId(used);
    groups.push({ groupId: newGroupId(), sdtPr: newSdtPr(kind, { id, ...(tag !== undefined ? { tag } : {}), ...(title !== undefined ? { title } : {}) }) });
    if (kind === 'RepeatingSection') { used.add(id); groups.push({ groupId: newGroupId(), sdtPr: newSdtPr('RepeatingSectionItem', { id: allocateSdtId(used) }) }); }
    const tr = state.tr;
    for (let r = top; r <= bottom; r++) {
      const stack = stacks[r - top]!;
      const at = depths[r - top]!;
      stack.splice(at, 0, ...groups.map((g) => ({ groupId: g.groupId, first: r === top, last: r === bottom })));
      const row = table.child(r);
      tr.setNodeMarkup(rowPosition(table, tablePos, r), undefined, { ...row.attrs, sdt: stack });
    }
    for (const g of groups) addControlWrite(tr, { groupId: g.groupId, entry: { form: 'row', sdtPr: g.sdtPr } });
    normalizeRowGroups(tr);
    // The cells of the wrapped rows selected: the highlight then shows what the control holds, where the cells
    // dragged over, left selected, read as the control's extent (Jason, 2026-10-08).
    const map = TableMap.get(table);
    tr.setSelection(CellSelection.create(tr.doc, rect.tableStart + map.map[top * map.width]!, rect.tableStart + map.map[bottom * map.width + map.width - 1]!));
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: isInTable,
});

register<InsertControlArgs | undefined>({
  name: 'contentControl.wrapCells',
  tracking: 'untracked',
  label: 'Wrap cells in content control',
  args: { kind: "'RichText' (default)", tag: 'w:tag', title: 'w:alias' },
  run: ({ kind = 'RichText', tag, title } = {}) => (state, dispatch) => {
    if (!isInTable(state) || kind === 'RepeatingSection') return false;
    const session = sessionOf(state);
    const rect = selectedRect(state);
    if (rect.bottom - rect.top !== 1) return false;   // one row: a cell-level control wraps cells of a row
    const table = rect.table;
    const tablePos = rect.tableStart - 1;
    const map = TableMap.get(table);
    const row = table.child(rect.top);
    const rowPos = rowPosition(table, tablePos, rect.top);
    // The cells of this row the columns cover, as cell indexes.
    const indexes = new Set<number>();
    for (let c = rect.left; c < rect.right; c++) {
      const pos = map.map[rect.top * map.width + c]!;
      if (map.findCell(pos).top !== rect.top) return false;   // a merge from above: not this row's cell
      let offset = 0;
      row.forEach((cell, o, i) => { if (rowPos + 1 + o === rect.tableStart + pos) { indexes.add(i); offset = o; } });
      void offset;
    }
    const cells = [...indexes].sort((a, b) => a - b);
    if (cells.length === 0) return false;
    const first = cells[0]!;
    const last = cells[cells.length - 1]!;
    const stacks: SdtGroup[][] = cells.map((i) => ((row.child(i).attrs['sdt'] as SdtGroup[] | null) ?? []).map((g) => ({ ...g })));
    const depths = outerCount(stacks, (id) => cellExtent(row, id), first, last);
    if (!depths) return false;
    const groupId = newGroupId();
    const sdtPr = newSdtPr(kind, { id: allocateSdtId(usedSdtIds(session, state.doc)), ...(tag !== undefined ? { tag } : {}), ...(title !== undefined ? { title } : {}) });
    const tr = state.tr;
    cells.forEach((i, k) => {
      const stack = stacks[k]!;
      stack.splice(depths[k]!, 0, { groupId, first: i === first, last: i === last });
      let cellPos = rowPos + 1;
      for (let j = 0; j < i; j++) cellPos += row.child(j).nodeSize;
      tr.setNodeMarkup(cellPos, undefined, { ...row.child(i).attrs, sdt: stack });
    });
    addControlWrite(tr, { groupId, entry: { form: 'cell', sdtPr } });
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: (state) => isInTable(state) && selectedRect(state).bottom - selectedRect(state).top === 1,
});

register<{ keepContent?: boolean } | undefined>({
  name: 'contentControl.delete',
  tracking: 'untracked',
  label: 'Delete content control',
  args: { keepContent: 'true unwraps the control and keeps what it held (the default false removes both)' },
  choices: { keepContent: ['false', 'true'] },
  run: ({ keepContent = false } = {}) => (state, dispatch) => {
    const target = controlAt(state);
    if (!target) return false;
    const tr = state.tr;
    const schema = state.schema;
    if (isNodeTarget(target)) {
      const { pos, node } = target;
      if (keepContent) {
        tr.replaceWith(pos, pos + node.nodeSize, node.content);
      } else {
        tr.delete(pos, pos + node.nodeSize);
        const $pos = tr.doc.resolve(pos);
        if ($pos.parent.childCount === 0 && !$pos.parent.inlineContent) tr.insert(pos, schema.nodes['paragraph']!.create());
      }
      tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(pos, tr.doc.content.size))));
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    const { groupId, table, tablePos } = target;
    const without = (stack: SdtGroup[] | null): SdtGroup[] | null => {
      const next = (stack ?? []).filter((g) => g.groupId !== groupId);
      return next.length > 0 ? next : null;
    };
    if (target.form === 'row') {
      const [first, last] = target.rows;
      if (keepContent) {
        for (let r = first; r <= last; r++) tr.setNodeMarkup(rowPosition(table, tablePos, r), undefined, { ...table.child(r).attrs, sdt: without(table.child(r).attrs['sdt'] as SdtGroup[] | null) });
      } else if (last - first + 1 >= table.childCount) {
        tr.delete(tablePos, tablePos + table.nodeSize);
        const $pos = tr.doc.resolve(tablePos);
        if ($pos.parent.childCount === 0 && !$pos.parent.inlineContent) tr.insert(tablePos, schema.nodes['paragraph']!.create());
      } else {
        const from = rowPosition(table, tablePos, first);
        const to = rowPosition(table, tablePos, last) + table.child(last).nodeSize;
        tr.delete(from, to);
      }
    } else {
      const row = table.child(target.row);
      const rowPos = rowPosition(table, tablePos, target.row);
      const [first, last] = target.cells;
      // Right to left, so the positions before an edit stay valid.
      for (let i = last; i >= first; i--) {
        let cellPos = rowPos + 1;
        for (let j = 0; j < i; j++) cellPos += row.child(j).nodeSize;
        const cell = row.child(i);
        if (!keepContent) tr.replaceWith(cellPos + 1, cellPos + cell.nodeSize - 1, schema.nodes['paragraph']!.create());
        tr.setNodeMarkup(cellPos, undefined, { ...cell.attrs, sdt: without(cell.attrs['sdt'] as SdtGroup[] | null) });
      }
    }
    addControlWrite(tr, { groupId, entry: null });
    normalizeRowGroups(tr);
    tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(state.selection.from, tr.doc.content.size))));
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: (state) => controlAt(state) !== undefined,
});

register<ControlProperties>({
  name: 'contentControl.properties',
  tracking: 'untracked',
  label: 'Content control properties',
  choices: { appearance: ['BoundingBox', 'Tags', 'Hidden'], cannotDelete: ['true', 'false'], cannotEdit: ['true', 'false'], removeWhenEdited: ['true', 'false'] },
  args: { tag: 'w:tag', title: 'w:alias', appearance: "'BoundingBox' | 'Tags' | 'Hidden'", color: "'#RRGGBB' or ''", cannotDelete: 'boolean', cannotEdit: 'boolean', removeWhenEdited: 'boolean', showingPlaceholder: 'boolean', multiLine: 'boolean (plain text)' },
  run: (patch) => (state, dispatch) => {
    const target = controlAt(state);
    const session = sessionOf(state);
    if (!target) return false;
    const current = sdtPrOf(target, session?.controlEntry ?? (() => undefined));
    const next = applyControlProperties(current, patch);
    const tr = state.tr;
    writeSdtPr(tr, target, next);
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => controlAt(state) !== undefined,
});

export interface BindArgs {
  xpath: string;
  prefixMappings?: string;
  /** The part; the data part by the specification's rules when omitted. */
  storeItemID?: string;
  /** Also write the XPaths entry and the `od:xpath` tag (REC-002); the default is yes when the template has an XPaths part or the control an `od:xpath`. */
  opendope?: boolean;
}

/**
 * A control bound in place to a node holding a Word document as Flat OPC (ED-005 section 12.124): made Word's form, a
 * block-level rich text control mapped by `w15:dataBinding`, keeping its id, tag and title. A plain text control becomes rich
 * text, its old binding replaced; a run-level control alone in its paragraph becomes a block-level control around that
 * paragraph, its content the paragraph's (`fragmentPlace`). Any other is refused, as `dataBindRefusal` says. Nothing is filled
 * here: the view's refresh fills the control through the engine's bind pass (part 3), as for a new control.
 */
function bindFragment(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, session: EditorDocument, target: ControlTarget, current: unknown, binding: BindArgs): boolean {
  const kind = sdtKind(current);
  const place = fragmentPlace(state, target);
  if ((kind !== 'PlainText' && kind !== 'RichText') || !place) return false;
  const rich = withBinding(kind === 'PlainText' ? withoutSdtItem(current, 'text') : current, null);
  let next = boundSdtPr(session, rich, binding);
  if (next === undefined || sdtBinding(next)?.w15 !== true) return false;
  // An XHTML control no longer: the two switches exclude each other.
  const params = parseTag(sdtTag(next));
  if (params.some((p) => p.key === 'od:ContentType')) next = applyControlProperties(next, { tag: formatTag(withTagParam(params, 'od:ContentType', undefined)) });
  const tr = state.tr;
  if ('block' in place) {
    writeSdtPr(tr, place.block, next);
  } else {
    const paragraph = state.doc.nodeAt(place.paragraph)!;
    const content: PMNode[] = [];
    paragraph.forEach((child) => { if (child === place.control.node) child.forEach((c) => { content.push(c); }); else content.push(child); });
    tr.replaceWith(place.paragraph, place.paragraph + paragraph.nodeSize, state.schema.nodes['sdtBlock']!.create({ sourceId: null, sdtPr: next }, paragraph.copy(Fragment.from(content))));
    tr.setSelection(TextSelection.create(tr.doc, place.paragraph + 2));
  }
  dispatch?.(tr.scrollIntoView());
  return true;
}

/**
 * Where a Word document bound to an existing control goes (REQ-060: block-level or in a cell): a block-level control itself;
 * a run-level control with nothing beside it in its paragraph but markers, that paragraph, where a block control can stand;
 * undefined for any other.
 */
function fragmentPlace(state: EditorState, target: ControlTarget): { block: NodeTarget } | { paragraph: number; control: NodeTarget } | undefined {
  if (!isNodeTarget(target)) return undefined;
  if (target.form === 'block') return { block: target };
  const $at = state.doc.resolve(target.pos);
  const paragraph = $at.parent;
  if (paragraph.type.name !== 'paragraph' || $at.depth < 1 || !['doc', 'sdtBlock', 'table_cell'].includes($at.node($at.depth - 1).type.name)) return undefined;
  let alone = true;
  paragraph.forEach((child) => {
    if (child === target.node) return;
    if (child.isText ? child.text!.trim() !== '' : !(child.type.name === 'opaqueInline' && PLAIN_TEXT_MARKERS.has(child.attrs['kind'] as string))) alone = false;
  });
  return alone ? { paragraph: $at.before(), control: target } : undefined;
}

/** The template view of a session, loaded; throws naming the fix otherwise. */
function loadedTemplate(session: EditorDocument | undefined): EditorDocument['template'] {
  if (!session) throw new OpenDopeError('no session in the editor state');
  if (!session.template.loaded) throw new OpenDopeError('the custom XML parts are not loaded: await session.template.load() once');
  return session.template;
}

register<BindArgs | null>({
  name: 'contentControl.bind',
  tracking: 'untracked',
  label: 'Bind content control',
  args: { xpath: 'the XPath (null unbinds: w:dataBinding and od:xpath removed)', prefixMappings: "xmlns:ns0='...'", storeItemID: 'the part; the data part when omitted', opendope: 'also the XPaths entry and od:xpath' },
  run: (binding) => (state, dispatch) => {
    const target = controlAt(state);
    const session = sessionOf(state);
    if (!target || !session) return false;
    const current = sdtPrOf(target, session.controlEntry);
    const params = parseTag(sdtTag(current));
    const tr = state.tr;
    if (binding === null) {
      let next: unknown = withBinding(current, null);
      if (params.some((p) => p.key === 'od:xpath')) next = applyControlProperties(next, { tag: formatTag(withTagParam(params, 'od:xpath', undefined)) });
      writeSdtPr(tr, target, next);
      dispatch?.(tr);
      return true;
    }
    const base = isNodeTarget(target) ? repeatBaseAround(state.doc, target.pos + 1, session) : undefined;
    const relative = base !== undefined ? relativeToRepeat(binding.xpath, base) : undefined;
    const asBound = relative !== undefined ? { ...binding, xpath: relative } : binding;
    // A Word document as Flat OPC (ED-005 proposal 46, decision (e); section 12.124): Word's form, as a bind with no control
    // at the caret makes it, the view's refresh filling it.
    if (session.template.loaded && isFlatOpc(nodeOfBinding(session, binding)?.text ?? '')) return bindFragment(state, dispatch, session, target, current, asBound);
    // A node holding XHTML binds through the XPaths entry with the content type, the view converting it (section 12.79).
    if (session.template.loaded && xhtmlKindOf(nodeOfBinding(session, binding)?.text ?? '') !== 'none') {
      const xhtml = xhtmlBoundSdtPr(session, current, asBound);
      if (xhtml === undefined) return false;
      writeSdtPr(tr, target, xhtml);
      dispatch?.(tr);
      return true;
    }
    let next = boundSdtPr(session, current, asBound);
    if (next === undefined) return false;
    // A node mapped to a rich text control takes no other kind of control, as Word refuses it (core-ts check 35, 1e).
    if (mappedKind(next) !== 'RichText' && mappedToRichText(state, binding)) return false;
    const template = session.template;
    // The control shows the bound value at once, as Word's XML mapping does (ED-005 section 12.54); tracked
    // over text that was not bound, never over a bound control's, which is the data's (section 12.57).
    if (template.loaded && isNodeTarget(target)) {
      const bound = sdtBinding(next);
      next = writeBoundValue(tr, target.pos, next, bound ? boundNodeOf(session, bound)?.nodeValue : undefined, [], isBound(current) ? undefined : trackingOf(state));
    }
    writeSdtPr(tr, target, next);
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => controlAt(state) !== undefined,
});


register<{ on: boolean }>({
  name: 'contentControl.xhtml',
  tracking: 'untracked',
  label: 'Bind the control to XHTML',
  args: { on: 'true: the bound node holds XHTML markup, converted into the control (od:ContentType=application/xhtml+xml on the tag, the binding as an od:xpath entry, no w:dataBinding); false: a text binding again' },
  // ED-005 proposal 44 (section 12.75): the control keeps its binding as the XPaths entry docx4j's
  // bind.xslt reads, and loses its w:dataBinding, which Word would otherwise fill with the escaped
  // markup on open. The view converts the markup after the command.
  run: ({ on }) => (state, dispatch) => {
    const target = controlAt(state);
    const session = sessionOf(state);
    if (!target || !session) return false;
    if (session.template.loaded === false) return false;
    const current = sdtPrOf(target, session.controlEntry);
    const params = parseTag(sdtTag(current));
    const tag = readTag(sdtTag(current));
    const template = session.template;
    const tr = state.tr;
    if (on) {
      // The binding: the w:dataBinding's, else the entry's.
      const binding = sdtBinding(current) ?? (tag.xpath !== undefined ? template.xpaths?.entries.find((e) => e.id === tag.xpath) : undefined);
      if (!binding || binding.xpath === '') return false;
      const entryId = tag.xpath !== undefined && template.xpaths?.entries.some((e) => e.id === tag.xpath) ? tag.xpath
        : template.upsertEntry({ xpath: binding.xpath, storeItemID: binding.storeItemID, prefixMappings: binding.prefixMappings }).id;
      let next: unknown = withBinding(current, null);
      next = applyControlProperties(next, { tag: formatTag(withTagParam(withTagParam(withTagParam(params, 'od:xpath', entryId), 'od:ContentType', XHTML_CONTENT_TYPE), 'od:progid', undefined)) });
      writeSdtPr(tr, target, next);
    } else {
      if (tag.contentType !== XHTML_CONTENT_TYPE) return false;
      let next: unknown = applyControlProperties(current, { tag: formatTag(withTagParam(params, 'od:ContentType', undefined)) });
      // A text binding again: the w:dataBinding back from the entry, so the text machinery sees it.
      const entry = tag.xpath !== undefined ? template.xpaths?.entries.find((e) => e.id === tag.xpath) : undefined;
      if (entry) next = withBinding(next, { xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings });
      writeSdtPr(tr, target, next);
    }
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => controlAt(state) !== undefined,
});

register<{ on: boolean }>({
  name: 'contentControl.wordFragment',
  tracking: 'untracked',
  label: 'Bind the control to a Word fragment',
  args: { on: 'true: OpenDoPE\'s Word fragment, the bound node holding a Word document as Flat OPC (od:progid=Word.Document on the tag, the binding as an od:xpath entry, no w:dataBinding), which a processor resolves; false: Word\'s own mapping again (w15:dataBinding on a rich text control)' },
  // ED-005 proposal 46 (decision (e)): the Control tab's switch between Word's form, a rich text control mapped by
  // w15:dataBinding that Word fills on open, and OpenDoPE's, which a processor resolves and nothing refills. As the
  // XHTML switch does, the binding moves to the XPaths entry docx4j's bind.xslt reads; the two switches are exclusive.
  run: ({ on }) => (state, dispatch) => {
    const target = controlAt(state);
    const session = sessionOf(state);
    if (!target || !session) return false;
    if (session.template.loaded === false) return false;
    const current = sdtPrOf(target, session.controlEntry);
    const params = parseTag(sdtTag(current));
    const tag = readTag(sdtTag(current));
    const template = session.template;
    const tr = state.tr;
    if (on) {
      const binding = sdtBinding(current) ?? (tag.xpath !== undefined ? template.xpaths?.entries.find((e) => e.id === tag.xpath) : undefined);
      if (!binding || binding.xpath === '') return false;
      const entryId = tag.xpath !== undefined && template.xpaths?.entries.some((e) => e.id === tag.xpath) ? tag.xpath
        : template.upsertEntry({ xpath: binding.xpath, storeItemID: binding.storeItemID, prefixMappings: binding.prefixMappings }).id;
      let next: unknown = withBinding(current, null);
      next = applyControlProperties(next, { tag: formatTag(withTagParam(withTagParam(withTagParam(params, 'od:xpath', entryId), 'od:progid', WORD_DOCUMENT_PROGID), 'od:ContentType', undefined)) });
      writeSdtPr(tr, target, next);
    } else {
      if (tag.progid !== WORD_DOCUMENT_PROGID) return false;
      let next: unknown = applyControlProperties(current, { tag: formatTag(withTagParam(params, 'od:progid', undefined)) });
      // Word's mapping again: a rich text control's w15:dataBinding from the entry, which Word fills on open.
      const entry = tag.xpath !== undefined ? template.xpaths?.entries.find((e) => e.id === tag.xpath) : undefined;
      if (entry) next = withBinding(next, { xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings }, sdtKind(current) === 'RichText');
      writeSdtPr(tr, target, next);
    }
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => controlAt(state) !== undefined,
});

/**
 * A control's properties bound to a node: the `w:dataBinding`, and with `opendope` (the default when the
 * template has an XPaths part or the control an `od:xpath`) the XPaths entry and the tag's `od:xpath`.
 * Undefined when the part is not there or the XPath selects nothing, as Word's mapping writes nothing then.
 */
function boundSdtPr(session: EditorDocument, current: unknown, binding: BindArgs): unknown | undefined {
  const template = session.template;
  const params = parseTag(sdtTag(current));
  let storeItemID = binding.storeItemID;
  if (template.loaded) {
    const part = storeItemID !== undefined ? template.parts.find((p) => p.id.replace(/[{}]/g, '').toLowerCase() === storeItemID!.replace(/[{}]/g, '').toLowerCase()) : template.dataPart.part;
    if (!part) return undefined;
    storeItemID = part.id;
    try {
      if (!part.selectSingleNode(binding.xpath, binding.prefixMappings)) return undefined;
    } catch {
      return undefined;
    }
  } else if (storeItemID === undefined) {
    return undefined;
  }
  // `w15:dataBinding` on a repeating section, and on a rich text control as Word's pane binds one (core-ts
  // check 35, 1a), unless it is bound with `w:dataBinding` already, which Word takes as plain text (1f).
  const kind = sdtKind(current);
  const w15 = kind === 'RepeatingSection' || (kind === 'RichText' && sdtBinding(current)?.w15 !== false);
  let next: unknown = withBinding(current, { xpath: binding.xpath, storeItemID, ...(binding.prefixMappings ? { prefixMappings: binding.prefixMappings } : {}) }, w15);
  const existingRef = readTag(sdtTag(current)).xpath;
  const opendope = binding.opendope ?? (template.loaded && (template.openDopeParts.xpaths !== undefined || existingRef !== undefined));
  if (opendope) {
    const entry = loadedTemplate(session).upsertEntry({ ...(existingRef !== undefined ? { id: existingRef } : {}), xpath: binding.xpath, storeItemID, ...(binding.prefixMappings !== undefined ? { prefixMappings: binding.prefixMappings } : {}) });
    next = applyControlProperties(next, { tag: formatTag(withTagParam(params, 'od:xpath', entry.id)) });
  }
  return next;
}

/** Whether a control is bound: a `w:dataBinding`, or an OpenDoPE `od:xpath`. Its content is the data's. */
function isBound(sdtPr: unknown): boolean {
  return sdtBinding(sdtPr) !== undefined || readTag(sdtTag(sdtPr)).xpath !== undefined;
}

/** The session's tracking where a bind can record it: never in a note or a header, whose projection holds no revision marks. */
function trackingOf(state: EditorState): Tracking | undefined {
  return storyOf(state) ? undefined : sessionOf(state)?.tracking();
}

/** The kinds that hold other controls, as Word lets them: the Data tab's bind makes its control inside one. */
const HOLDS_CONTROLS: ReadonlySet<SdtKind> = new Set<SdtKind>(['RichText', 'Group', 'RepeatingSection', 'RepeatingSectionItem', 'BuildingBlockGallery', 'Bibliography']);

/**
 * Whether the Data tab's bind binds a control in place rather than nesting one in it: a control that
 * holds no others (plain text, a drop-down, a date, ...); a rich text control bound with
 * `w:dataBinding`, which Word takes as plain text (core-ts check 35, 1f); and one bound by `od:xpath`
 * alone, an open question kept in place for now (Jason, 2026-10-03). Word's pane nests inside a rich
 * text control mapped with `w15:dataBinding` (1e), and inside the rest of the kinds that hold controls.
 */
function bindsInPlace(sdtPr: unknown): boolean {
  const kind = sdtKind(sdtPr);
  if (!HOLDS_CONTROLS.has(kind)) return true;
  if (kind !== 'RichText') return false;
  const binding = sdtBinding(sdtPr);
  return binding ? !binding.w15 : readTag(sdtTag(sdtPr)).xpath !== undefined;
}

/**
 * The control the Data tab's bind binds in place (Word's Map to Selected Content Control): the innermost
 * control at the selection, where `bindsInPlace` says so (ED-005 sections 12.57 and 12.58). Undefined
 * when there is no control there, or one to nest in: the bind then makes a plain text control there,
 * nested as Word's XML Mapping Pane nests one (core-ts check 34, 1h; check 35, 1e). A cell
 * selection's row or cell control is bound in place, since a cell selection cannot hold a plain text
 * control.
 */
export function controlToBind(state: EditorState): ControlTarget | undefined {
  const target = controlAt(state);
  if (!target || state.selection instanceof CellSelection) return target;
  return bindsInPlace(sdtPrOf(target, sessionOf(state)?.controlEntry ?? (() => undefined))) ? target : undefined;
}

/** The node a binding selects, in the part it names or else the data part; undefined when there is none. */
function nodeOfBinding(session: EditorDocument, binding: BindArgs): CustomXmlNode | undefined {
  const template = session.template;
  if (!template.loaded) return undefined;
  const part = binding.storeItemID !== undefined ? template.parts.find((p) => p.id.replace(/[{}]/g, '').toLowerCase() === binding.storeItemID!.replace(/[{}]/g, '').toLowerCase()) : template.dataPart.part;
  try {
    return part?.selectSingleNode(binding.xpath, binding.prefixMappings) ?? undefined;
  } catch {
    return undefined;
  }
}

/**
 * Whether the node a binding selects is mapped to a rich text control (`w15:dataBinding`) in the
 * document: Word refuses to map it to anything else, saying "The custom XML node is already mapped
 * to a rich text content control, so it can't be mapped to a plain text content control" (core-ts
 * check 35, 1e, measured for a control nested in that rich text control).
 */
function mappedToRichText(state: EditorState, binding: BindArgs): boolean {
  const session = sessionOf(state);
  const node = session ? nodeOfBinding(session, binding) : undefined;
  if (!session || !node) return false;
  return boundControlsIn(state.doc).some((c) => c.binding.w15 && sdtKind(c.node.attrs['sdtPr']) === 'RichText' && boundNodeOf(session, c.binding)?.node === node.node);
}

/**
 * Why the Data tab's bind is refused: the plain text control it would make cannot be (`PlainTextRefusal`), or the node is
 * mapped to a rich text control and what it would bind is not one (`mapped`), or block-level markup inside a paragraph
 * with content (`blockXhtml`), or image data the editor cannot place as a new picture (`imageFormat`: TIFF, EMF, WMF, SVG).
 */
export type BindRefusal = PlainTextRefusal | 'mapped' | 'blockXhtml' | 'imageFormat' | 'blockFragment' | 'fragmentKind';

/** Whether the node a binding selects holds a Word document as Flat OPC, and so binds as Word's mapped rich text (ED-005 proposal 46, decision (e)). */
export function bindsFragment(state: EditorState, binding: BindArgs): boolean {
  const session = sessionOf(state);
  return isFlatOpc(session ? nodeOfBinding(session, binding)?.text ?? '' : '');
}

/** Whether the node a binding selects holds image data (by its head), and so binds as a picture (section 12.117). */
export function bindsPicture(state: EditorState, binding: BindArgs): boolean {
  const session = sessionOf(state);
  const value = session ? nodeOfBinding(session, binding)?.text ?? '' : '';
  return xhtmlKindOf(value) === 'none' && looksLikeImageBase64(value);
}

/**
 * Why the Data tab's bind of `binding` at the selection is refused, or undefined when it can go ahead:
 * binding `controlToBind`'s control in place, or else making a plain text control
 * (`contentControl.insertBound`). Both commands refuse the same cases.
 */
export function dataBindRefusal(state: EditorState, binding: BindArgs): BindRefusal | undefined {
  const target = controlToBind(state);
  // A node holding XHTML (ED-005 section 12.79): block-level markup needs a paragraph of its own, so with no control
  // to bind in place it goes around an empty paragraph and is refused inside one with content; inline markup takes a
  // run-level control, under the plain text rules for where one can go.
  const session = sessionOf(state);
  const value = session ? nodeOfBinding(session, binding)?.text ?? '' : '';
  const kind = xhtmlKindOf(value);
  if (target === undefined && kind === 'block') return paragraphIsBlank(state) ? undefined : 'blockXhtml';
  // A Word document (proposal 46): block content, so a rich text control around a blank paragraph, as block-level XHTML; a
  // control at the caret made Word's form when it is plain or rich text and can be block-level (section 12.124).
  if (isFlatOpc(value)) {
    if (target === undefined) return paragraphIsBlank(state) ? undefined : 'blockFragment';
    const kind = sdtKind(sdtPrOf(target, session?.controlEntry ?? (() => undefined)));
    if (kind !== 'PlainText' && kind !== 'RichText') return 'fragmentKind';
    return fragmentPlace(state, target) ? undefined : 'blockFragment';
  }
  // A node holding image data (section 12.117): a picture control, under the plain text rules for where a run-level
  // control can go, when the editor can place the image; refused otherwise, saying so.
  const image = kind === 'none' && looksLikeImageBase64(value);
  if (target === undefined && image && pictureFactsOf(value)?.placeable !== true) return 'imageFormat';
  if (target === undefined) {
    const refusal = plainTextControlRefusal(state);
    if (refusal) return refusal;
  }
  if (kind !== 'none' || image) return undefined;
  const richText = target !== undefined && isNodeTarget(target) && mappedKind(target.node.attrs['sdtPr']) === 'RichText';
  return !richText && mappedToRichText(state, binding) ? 'mapped' : undefined;
}

/** Whether the selection sits in a paragraph with nothing but whitespace and markers (a block control can go around it). */
function paragraphIsBlank(state: EditorState): boolean {
  const { $from, $to } = state.selection;
  if (state.selection instanceof CellSelection || !$from.sameParent($to) || $from.parent.type.name !== 'paragraph') return false;
  const holder = $from.depth >= 1 ? $from.node($from.depth - 1).type.name : '';
  if (!['doc', 'sdtBlock', 'table_cell'].includes(holder)) return false;
  let blank = true;
  $from.parent.forEach((c) => { if (c.isText ? c.text!.trim() !== '' : !(c.type.name === 'opaqueInline' && PLAIN_TEXT_MARKERS.has(c.attrs['kind'] as string))) blank = false; });
  return blank;
}

/**
 * A control's properties bound to a node holding XHTML (section 12.79): as the XHTML switch writes them, the entry in
 * the XPaths part (relative to a repeat around the target) and the tag naming it with the content type, no `w:dataBinding`.
 */
function xhtmlBoundSdtPr(session: EditorDocument, current: unknown, binding: BindArgs): unknown | undefined {
  const template = session.template;
  const part = binding.storeItemID !== undefined ? template.parts.find((p) => p.id.replace(/[{}]/g, '').toLowerCase() === binding.storeItemID!.replace(/[{}]/g, '').toLowerCase()) : template.dataPart.part;
  if (!part) return undefined;
  const entry = template.upsertEntry({ xpath: binding.xpath, storeItemID: part.id, ...(binding.prefixMappings ? { prefixMappings: binding.prefixMappings } : {}) });
  const params = parseTag(sdtTag(current));
  let next: unknown = withBinding(current, null);
  next = applyControlProperties(next, { tag: formatTag(withTagParam(withTagParam(params, 'od:xpath', entry.id), 'od:ContentType', XHTML_CONTENT_TYPE)) });
  return next;
}

/** The inline items a plain text control may hold besides text: a tab, a line break, and markers with no content of their own. */
const PLAIN_TEXT_ITEMS: ReadonlySet<string> = new Set(['tab', 'hardBreak']);
const PLAIN_TEXT_MARKERS: ReadonlySet<string> = new Set(['bookmarkStart', 'bookmarkEnd', 'permStart', 'permEnd', 'commentReference', 'proofErr', 'lastRenderedPageBreak']);

/** Why the selection cannot become a new plain text control: table cells, more than one paragraph, or content other than text. */
export type PlainTextRefusal = 'cells' | 'paragraphs' | 'content';

/**
 * Whether the selection can become a plain text content control, which holds text within one paragraph
 * (Word's `w:text`): undefined when it can (a caret, or text within one paragraph or run control), else
 * why not. A selection holding a picture, a note reference, a field or another control is refused,
 * since binding would replace it with the node's value; so are a selection across paragraphs and one
 * holding a picture, where Word takes the first paragraph's part and keeps the picture (core-ts check
 * 34, 1f and 1g; refused at Jason's word, ED-005 section 12.57).
 */
export function plainTextControlRefusal(state: EditorState): PlainTextRefusal | undefined {
  const sel = state.selection;
  if (sel instanceof CellSelection) return 'cells';
  const { $from, $to } = sel;
  if (!$from.sameParent($to) || !['paragraph', 'sdtInline'].includes($from.parent.type.name)) return 'paragraphs';
  let other = false;
  state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
    if (other) return false;
    if (pos < $from.pos) return true;   // the paragraph or run control the selection is in
    if (node.isText) return false;
    if (!node.isInline) return true;
    if (PLAIN_TEXT_ITEMS.has(node.type.name)) return false;
    if (node.type.name === 'opaqueInline' && PLAIN_TEXT_MARKERS.has(node.attrs['kind'] as string)) return false;
    other = true;
    return false;
  });
  return other ? 'content' : undefined;
}

/**
 * A plain text content control made at the caret, or around a selection within one paragraph, and bound
 * to a node in one step, as Word's XML Mapping Pane inserts a mapped control (core-ts check 34): the
 * control shows the node's value (the selected text gives way to it, as Word's does and as
 * `contentControl.bind` replaces a control's), in the formatting there; at a caret in an empty
 * paragraph, a block-level control around the paragraph, as Word's (1d). Inside a control that holds
 * controls (rich text, a group, a repeating section), nested in it, as Word's (1h); a placeholder there
 * gives way to it. One undo step. With tracking on, the selected text is kept inside the new control,
 * marked deleted, and the value marked inserted, as Word's (1j); the control itself is not marked, the
 * model having no revision on a control (Word marks it with `w:customXmlInsRangeStart`). Refused where
 * `plainTextControlRefusal` says (Word takes the first paragraph of several and keeps a picture: the
 * editor says why instead, at Jason's word), where `controlToBind` has a control to bind (Word's Map to
 * Selected Content Control), and where the XPath selects nothing. A node holding XHTML makes a rich text control
 * (section 12.79), and one holding image data Word's picture control, the view placing the image (section 12.117).
 */
register<BindArgs>({
  name: 'contentControl.insertBound',
  tracking: 'untracked',
  label: 'Insert bound plain text control',
  args: { xpath: 'the XPath', prefixMappings: "xmlns:ns0='...'", storeItemID: 'the part; the data part when omitted', opendope: 'also the XPaths entry and od:xpath' },
  run: (binding) => (state, dispatch) => {
    const session = sessionOf(state);
    if (!session || !binding || !session.template.loaded || controlToBind(state) || dataBindRefusal(state, binding) !== undefined) return false;
    const base = repeatBaseAround(state.doc, state.selection.from, session);
    const relative = base !== undefined ? relativeToRepeat(binding.xpath, base) : undefined;
    const asBound = relative !== undefined ? { ...binding, xpath: relative } : binding;
    const kind = xhtmlKindOf(nodeOfBinding(session, binding)?.text ?? '');
    if (bindsFragment(state, binding)) {
      // A Word document as Flat OPC (ED-005 proposal 46, decision (e)): Word's form, a block-level rich text control mapped
      // by w15:dataBinding around the blank paragraph the caret is in (dataBindRefusal refused any other), as Word's XML
      // Mapping pane maps one. Word fills it from the node when it opens the file; the editor shows the fragment through the
      // engine's bind step (part 3). The Control tab's switch turns it into OpenDoPE's form.
      const made = boundSdtPr(session, newSdtPr('RichText', { id: allocateSdtId(usedSdtIds(session, state.doc)) }), asBound);
      if (made === undefined || sdtBinding(made)?.w15 !== true) return false;
      const $from = state.selection.$from;
      const at = $from.before();
      const tr = state.tr.replaceWith(at, $from.after(), state.schema.nodes['sdtBlock']!.create({ sourceId: null, sdtPr: made }, $from.parent));
      tr.setSelection(TextSelection.create(tr.doc, at + 2));
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    if (kind !== 'none') {
      // XHTML (section 12.79): a rich text control bound through the entry; block-level markup around the blank
      // paragraph the caret is in (dataBindRefusal refused any other), inline markup as a run-level control at the
      // selection. The view converts the markup into it.
      const made = xhtmlBoundSdtPr(session, newSdtPr('RichText', { id: allocateSdtId(usedSdtIds(session, state.doc)) }), asBound);
      if (made === undefined) return false;
      const tr = state.tr;
      const $from = state.selection.$from;
      if (kind === 'block') {
        const paragraph = $from.parent;
        const at = $from.before();
        tr.replaceWith(at, $from.after(), state.schema.nodes['sdtBlock']!.create({ sourceId: null, sdtPr: made }, paragraph));
        tr.setSelection(TextSelection.create(tr.doc, at + 2));
      } else {
        const { from, to } = state.selection;
        tr.replaceWith(from, to, state.schema.nodes['sdtInline']!.create({ sourceId: null, sdtPr: made }, state.doc.slice(from, to).content));
        tr.setSelection(TextSelection.create(tr.doc, from + 1));
      }
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    // A node holding image data: Word's picture control, mapped as Word's pane maps one (section 12.117).
    const picture = bindsPicture(state, binding);
    const sdtPr = boundSdtPr(session, newSdtPr(picture ? 'Picture' : 'PlainText', { id: allocateSdtId(usedSdtIds(session, state.doc)) }), asBound);
    const bound = sdtPr === undefined ? undefined : sdtBinding(sdtPr);
    if (sdtPr === undefined || !bound) return false;
    const tr = state.tr;
    // Inside a control showing its placeholder (a rich text control just made): the placeholder gives way, untracked.
    const around = controlAt(state);
    const holder = around && isNodeTarget(around) && readControl(around.node.attrs['sdtPr']).showingPlaceholder ? around : undefined;
    // The formatting the value takes: the selection's first text's, or at a caret the text's there; never the placeholder's.
    const marks = holder ? [] : state.selection.$from.marks();
    if (holder) {
      const $at = state.selection.$from;
      tr.setNodeMarkup(holder.pos, undefined, { ...holder.node.attrs, sdtPr: applyControlProperties(holder.node.attrs['sdtPr'], { showingPlaceholder: false }) });
      tr.delete($at.start(), $at.end());
      tr.setSelection(TextSelection.create(tr.doc, $at.start()));
    }
    const { from, to } = tr.selection;
    // A run-level control, at a caret in an empty paragraph too: Word's pane makes a block-level one there (core-ts
    // check 34, 1d), which the document's last paragraph cannot leave; Jason decided 2026-10-08 (docs/developer/word-divergences.md).
    const at = from;
    tr.replaceWith(from, to, state.schema.nodes['sdtInline']!.create({ sourceId: null, sdtPr }, tr.doc.slice(from, to).content));
    if (picture) {
      // The view places the image (`refreshPictureControls`, out of the undo history as for any bound picture control); until
      // it does, the control holds what was selected. Untracked, as a picture bind is (proposal 45 part 4, open).
      tr.setSelection(TextSelection.create(tr.doc, at + 1));
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    const next = writeBoundValue(tr, at, sdtPr, boundNodeOf(session, bound)?.nodeValue, marks, trackingOf(state));
    const control = tr.doc.nodeAt(at)!;
    tr.setNodeMarkup(at, undefined, { ...control.attrs, sdtPr: next });
    // The caret at the end of the new control's text.
    tr.setSelection(TextSelection.create(tr.doc, at + control.nodeSize - 1));
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: (state) => controlToBind(state) === undefined && plainTextControlRefusal(state) === undefined,
});

/** A role command: sets or clears a tag parameter, creating the definition when asked (REQ-007 and REQ-008 refuse). */
function roleCommand(key: 'od:condition' | 'od:repeat' | 'od:RptPosCon', define: (session: EditorDocument, id: string | undefined, args: RoleArgs) => string): (args: RoleArgs) => Command {
  return (args) => (state, dispatch) => {
    const target = controlAt(state);
    const session = sessionOf(state);
    if (!target || !session) return false;
    const current = sdtPrOf(target, session.controlEntry);
    const reading = readTag(sdtTag(current));
    let params = reading.params;
    if (args.id === null) {
      params = withTagParam(params, key, undefined);
    } else {
      if (key !== 'od:RptPosCon') {
        const role = key === 'od:condition' ? 'condition' : 'repeat';
        if (reading.role !== 'none' && reading.role !== role) return false;   // REQ-007
        if (sdtBinding(current) && sdtKind(current) !== 'RepeatingSection') return false;   // REQ-008
      }
      const template = loadedTemplate(session);
      let id: string;
      if (args.define) {
        id = define(session, args.id, args);
      } else {
        id = args.id;
        const exists = key === 'od:condition' ? template.conditions?.conditions.some((c) => c.id === id) : template.xpaths?.entries.some((e) => e.id === id);
        if (!exists) return false;   // REQ-030 for a condition; an entry the tag would dangle on otherwise
      }
      params = withTagParam(params, key, id);
    }
    const tr = state.tr;
    writeSdtPr(tr, target, applyControlProperties(current, { tag: formatTag(params) }));
    if (args.id !== null) {
      const repeatXPath = key === 'od:repeat' ? loadedTemplate(session).xpaths?.entries.find((e) => e.id === params.find((p) => p.key === key)?.value)?.xpath : undefined;
      adoptBindingsIntoXPaths(tr, state, session, target, repeatXPath !== undefined ? { target, xpath: repeatXPath } : undefined);
    }
    dispatch?.(tr);
    return true;
  };
}

/**
 * The XPath of the innermost repeat around a position (ED-005 section 12.77): a block or run control
 * tagged `od:repeat`, or a row group's; `override` is a role just set in the transaction, whose tag is
 * not in the document yet. Undefined outside any repeat, or when the repeat's entry is missing.
 */
function repeatBaseAround(doc: PMNode, pos: number, session: EditorDocument, override?: { target: ControlTarget; xpath: string }): string | undefined {
  const template = session.template;
  const entryXPath = (sdtPr: unknown): string | undefined => {
    const tag = readTag(sdtTag(sdtPr));
    if (tag.role !== 'repeat' || tag.roleId === undefined) return undefined;
    return template.xpaths?.entries.find((e) => e.id === tag.roleId)?.xpath;
  };
  const $pos = doc.resolve(Math.max(0, Math.min(pos, doc.content.size)));
  const groups = readControlEntries(session.store);
  for (let d = $pos.depth; d >= 1; d--) {
    const node = $pos.node(d);
    const at = $pos.before(d);
    if (node.type.name === 'sdtBlock' || node.type.name === 'sdtInline') {
      if (override && isNodeTarget(override.target) && override.target.pos === at) return override.xpath;
      const xpath = entryXPath(node.attrs['sdtPr']);
      if (xpath !== undefined) return xpath;
    }
    if (node.type.name === 'table_row' || node.type.name === 'table_cell') {
      const stack = ((node.attrs['sdt'] as { groupId: string }[] | null) ?? []).slice().reverse();
      for (const g of stack) {
        if (override && !isNodeTarget(override.target) && override.target.groupId === g.groupId) return override.xpath;
        const xpath = entryXPath(groups.get(g.groupId)?.sdtPr);
        if (xpath !== undefined) return xpath;
      }
    }
  }
  return undefined;
}

/**
 * Every bound control without an `od:xpath` given an XPaths entry and the tag naming it, once the
 * template has an XPaths part (ED-005 section 12.77). A control bound before the first role was
 * set got a `w:dataBinding` alone, since there was no part to enter it in (`boundSdtPr`'s default),
 * and docx4j's processor then warns "No XPaths part object" for each such binding inside a repeat
 * (its sanity test; it contextualises the `w:dataBinding` all the same). An entry with the same
 * XPath and part is reused; the role's own control is left (REQ-008: a role and a binding do not
 * meet, but a repeating section's is its own).
 */
function adoptBindingsIntoXPaths(tr: Transaction, state: EditorState, session: EditorDocument, except: ControlTarget, repeatSet?: { target: ControlTarget; xpath: string }): void {
  const template = session.template;
  if (!template.loaded) return;
  const same = (a: string, b: string): boolean => a.replace(/[{}]/g, '').toLowerCase() === b.replace(/[{}]/g, '').toLowerCase();
  const entryFor = (binding: { xpath: string; storeItemID: string; prefixMappings: string }, existing: string | undefined): string => {
    if (existing !== undefined && template.xpaths?.entries.some((e) => e.id === existing)) {
      const e = template.xpaths.entries.find((x) => x.id === existing)!;
      if (e.xpath !== binding.xpath) template.upsertEntry({ id: existing, xpath: binding.xpath, storeItemID: binding.storeItemID, ...(binding.prefixMappings ? { prefixMappings: binding.prefixMappings } : {}) });
      return existing;
    }
    const found = template.xpaths?.entries.find((e) => e.xpath === binding.xpath && same(e.storeItemID, binding.storeItemID) && (e.prefixMappings ?? '') === (binding.prefixMappings ?? ''));
    return (found ?? template.upsertEntry({ xpath: binding.xpath, storeItemID: binding.storeItemID, ...(binding.prefixMappings ? { prefixMappings: binding.prefixMappings } : {}) })).id;
  };
  // A control's properties with its binding entered and, inside a repeat, written relative to it; undefined when nothing changes.
  const adopted = (sdtPr: unknown, base: string | undefined): unknown | undefined => {
    const binding = sdtBinding(sdtPr);
    if (!binding || binding.xpath === '' || binding.storeItemID === '') return undefined;
    const params = parseTag(sdtTag(sdtPr));
    const existing = params.find((p) => p.key === 'od:xpath')?.value;
    const relative = base !== undefined ? relativeToRepeat(binding.xpath, base) : undefined;
    const xpath = relative ?? binding.xpath;
    const entryId = entryFor({ ...binding, xpath }, existing);
    if (existing === entryId && relative === undefined) return undefined;
    let next: unknown = sdtPr;
    if (relative !== undefined && relative !== binding.xpath) next = withBinding(next, { xpath: relative, storeItemID: binding.storeItemID, prefixMappings: binding.prefixMappings, ...(binding.w15 ? { w15: true } : {}) });
    if (existing !== entryId) next = applyControlProperties(next, { tag: formatTag(withTagParam(params, 'od:xpath', entryId)) });
    return next;
  };
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'sdtBlock' && node.type.name !== 'sdtInline') return true;
    if (isNodeTarget(except) && except.pos === pos) return true;
    const next = adopted(node.attrs['sdtPr'], repeatBaseAround(state.doc, pos + 1, session, repeatSet));
    if (next !== undefined) tr.setNodeMarkup(pos, undefined, { ...node.attrs, sdtPr: next });
    return true;
  });
  for (const [groupId, entry] of readControlEntries(session.store)) {
    if (!isNodeTarget(except) && except.groupId === groupId) continue;
    const next = adopted(entry.sdtPr, undefined);
    if (next !== undefined) addControlWrite(tr, { groupId, entry: { form: entry.form, sdtPr: next } });
  }
}

register<void>({
  name: 'opendope.adoptBindings',
  tracking: 'untracked',
  label: 'Enter the bound controls in the XPaths part',
  // For a template whose controls were bound before it had an XPaths part (ED-005 section 12.77): the
  // palette's and the console's way to give each its entry and od:xpath, which a role set does on its own.
  run: () => (state, dispatch) => {
    const session = sessionOf(state);
    if (!session || !session.template.loaded) return false;
    const tr = state.tr;
    adoptBindingsIntoXPaths(tr, state, session, { form: 'block', pos: -1, node: state.doc } as unknown as ControlTarget);
    if (!tr.docChanged && controlWritesOf(tr).length === 0) return false;
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => sessionOf(state)?.template.loaded === true,
});

export interface RoleArgs {
  /** The condition or entry id; null clears the parameter. */
  id: string | null;
  /** Create or replace the definition: a condition's operand tree, or an entry's expression (with `storeItemID` and `prefixMappings` as for bind). */
  define?: Operand | { xpath: string; storeItemID?: string; prefixMappings?: string };
  /** `@name` on the definition. */
  name?: string;
}

register<RoleArgs>({
  name: 'opendope.condition',
  tracking: 'untracked',
  label: 'Condition',
  args: { id: 'the condition id, or null to clear', define: "an operand tree, e.g. { op: 'xpathref', id: 'x5' }, to create or replace the condition", name: '@name' },
  run: roleCommand('od:condition', (session, id, args) => session.template.upsertCondition({ ...(id !== undefined ? { id } : {}), operand: args.define as Operand, ...(args.name ? { attrs: { name: args.name } } : {}) }).id),
  isEnabled: (state) => controlAt(state) !== undefined,
});

register<RoleArgs>({
  name: 'opendope.repeat',
  tracking: 'untracked',
  label: 'Repeat',
  args: { id: 'the XPaths entry id, or null to clear', define: "{ xpath, storeItemID?, prefixMappings? } to create or replace the entry (REQ-038: select the elements to repeat, not their container)", name: '@name' },
  run: roleCommand('od:repeat', (session, id, args) => {
    const d = args.define as { xpath: string; storeItemID?: string; prefixMappings?: string };
    return session.template.upsertEntry({ ...(id !== undefined ? { id } : {}), xpath: d.xpath, ...(d.storeItemID !== undefined ? { storeItemID: d.storeItemID } : {}), ...(d.prefixMappings !== undefined ? { prefixMappings: d.prefixMappings } : {}), ...(args.name ? { attrs: { name: args.name } } : {}) }).id;
  }),
  isEnabled: (state) => controlAt(state) !== undefined,
});

register<RoleArgs>({
  name: 'opendope.positionCondition',
  tracking: 'untracked',
  label: 'Position condition',
  args: { id: 'the XPaths entry id, or null to clear', define: "{ xpath: 'position()=last()-1' } to create or replace the entry (section 8.3)", name: '@name' },
  run: roleCommand('od:RptPosCon', (session, id, args) => {
    const d = args.define as { xpath: string };
    return session.template.upsertEntry({ ...(id !== undefined ? { id } : {}), xpath: d.xpath, ...(args.name ? { attrs: { name: args.name } } : {}) }).id;
  }),
  isEnabled: (state) => controlAt(state) !== undefined,
});

register<{ pos?: number; groupId?: string }>({
  name: 'contentControl.select',
  tracking: 'untracked',
  label: 'Select content control',
  args: { pos: "a block or run control's position (from listControls)", groupId: "a row or cell group's id" },
  run: ({ pos, groupId }) => (state, dispatch) => {
    const doc = state.doc;
    if (pos !== undefined) {
      const node = doc.nodeAt(pos);
      if (!node || (node.type.name !== 'sdtBlock' && node.type.name !== 'sdtInline')) return false;
      dispatch?.(state.tr.setSelection(TextSelection.between(doc.resolve(pos + 1), doc.resolve(pos + node.nodeSize - 1))).scrollIntoView());
      return true;
    }
    if (groupId === undefined) return false;
    let found: { table: PMNode; tablePos: number } | undefined;
    doc.descendants((n, p) => {
      if (found) return false;
      if (n.type.name === 'table' && rowExtent(n, groupId)) found = { table: n, tablePos: p };
      return !found;
    });
    if (found) {
      const { table, tablePos } = found;
      const [first, last] = rowExtent(table, groupId)!;
      const anchor = rowPosition(table, tablePos, first) + 1;
      const lastRow = table.child(last);
      const head = rowPosition(table, tablePos, last) + 1 + lastRow.content.size - lastRow.child(lastRow.childCount - 1).nodeSize;
      dispatch?.(state.tr.setSelection(CellSelection.create(doc, anchor, head)).scrollIntoView());
      return true;
    }
    // A cell group: the cells carrying it in their row.
    let cellSel: CellSelection | undefined;
    doc.descendants((n, p) => {
      if (cellSel || n.type.name !== 'table_row') return !cellSel;
      const extent = cellExtent(n, groupId);
      if (!extent) return false;
      let a = p + 1;
      for (let j = 0; j < extent[0]; j++) a += n.child(j).nodeSize;
      let h = p + 1;
      for (let j = 0; j < extent[1]; j++) h += n.child(j).nodeSize;
      cellSel = CellSelection.create(doc, a, h);
      return false;
    });
    if (!cellSel) return false;
    dispatch?.(state.tr.setSelection(cellSel).scrollIntoView());
    return true;
  },
});

/** The innermost control around the selection, for the panel (re-exported from the model). */
export { controlAt };
export type { EditorState };
