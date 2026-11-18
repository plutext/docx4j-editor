// SPDX-License-Identifier: GPL-3.0-or-later
//
// A content control's own properties, honoured on a local edit as Word honours them (ED-005 section
// 12.60):
//
// - **its lock** (`w:lock`): a control that cannot be deleted stays, and one whose contents cannot be
//   edited keeps them; an edit that would do otherwise is refused whole (`controlLockRefusal`), before
//   the tracker records it;
// - **its placeholder** (`w:showingPlcHdr`): what is typed or pasted in it replaces the placeholder,
//   which is no content, so goes untracked (`applyControlRules`, after the tracker);
// - **remove when edited** (`w:temporary`): an edit inside it removes the control and keeps its
//   content, in the same transaction (one undo step).
//
// The view's dispatch and the headless runner call both, as they call `recordTransaction`. What is
// not a person's edit passes: a peer's update or an undo (y-prosemirror's), a bound value written from
// the data (`BOUND_WRITE_META`: Word fills a locked mapped control), a resolution (accept, reject), and
// a comment. Only what a transaction touches is looked at: the controls around and inside its steps'
// ranges, never the whole document.
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { ySyncPluginKey } from 'y-prosemirror';
import type { SdtGroup } from './import.mjs';
import { applyControlProperties, readControl, sdtId } from './sdtpr.mjs';
import { BOUND_WRITE_META, PLACEHOLDER_STYLE } from './bound.mjs';
import { resolutionWriteOf } from './resolutions.mjs';
import { COMMAND_META, sessionOf } from './commands/spec.mjs';

/** Why a local edit was refused: it would remove a control that cannot be deleted, or change one whose contents cannot be edited. */
export type LockRefusal = 'cannotDelete' | 'cannotEdit';

/** A control a transaction touches: a block or run control's node, or a row's or a cell's for a group (`group` its id). */
interface Touched {
  pos: number;
  node: PMNode;
  sdtPr: unknown;
  group?: string;
}

/** The comment commands: a comment marks text, which a control's lock does not stop (Word lets locked content be commented). */
const COMMENT_COMMANDS = new Set(['range.insertComment', 'comment.reply', 'comment.setResolved', 'comment.delete']);

/** Whether a transaction is not a person's edit of the content: the rules leave it as it is. */
function exempt(tr: Transaction): boolean {
  if (!tr.docChanged || tr.getMeta(BOUND_WRITE_META) || resolutionWriteOf(tr) !== undefined) return true;
  if ((tr.getMeta(ySyncPluginKey) as { isChangeOrigin?: boolean } | undefined)?.isChangeOrigin) return true;
  const command = tr.getMeta(COMMAND_META) as string | undefined;
  return command !== undefined && COMMENT_COMMANDS.has(command);
}

/** Each step's range, in the document before the transaction. */
function stepRanges(tr: Transaction, steps = tr.steps.length): [number, number][] {
  const out: [number, number][] = [];
  for (let i = 0; i < steps; i++) {
    const step = tr.steps[i] as unknown as { from?: number; to?: number; pos?: number };
    const from = step.from ?? step.pos;
    if (from === undefined) continue;
    const to = step.to ?? from + 1;
    const back = tr.mapping.slice(0, i).invert();
    const size = tr.before.content.size;
    out.push([Math.max(0, Math.min(size, back.map(from, -1))), Math.max(0, Math.min(size, back.map(to, 1)))]);
  }
  return out;
}

/** The controls a transaction touches in the document before it: those around each step's range and inside it. */
function touchedControls(tr: Transaction, state: EditorState, steps?: number): Touched[] {
  const session = sessionOf(state);
  const entry = (groupId: string): unknown => session?.controlEntry(groupId)?.sdtPr;
  const seen = new Set<string>();
  const out: Touched[] = [];
  const add = (t: Touched): void => {
    const key = `${t.pos}:${t.group ?? ''}`;
    if (!seen.has(key)) { seen.add(key); out.push(t); }
  };
  for (const [a, b] of stepRanges(tr, steps)) {
    state.doc.nodesBetween(a, b, (node, pos) => {
      const name = node.type.name;
      if (name === 'sdtInline' || name === 'sdtBlock') add({ pos, node, sdtPr: node.attrs['sdtPr'] });
      else if (name === 'table_row' || name === 'table_cell') {
        for (const g of (node.attrs['sdt'] as SdtGroup[] | null) ?? []) add({ pos, node, sdtPr: entry(g.groupId), group: g.groupId });
      }
      return true;
    });
  }
  return out;
}

/** The node a touched control is after the transaction, or undefined when the transaction removed it. */
function survivor(tr: Transaction, t: Touched): { pos: number; node: PMNode } | undefined {
  const deletedRow = (n: PMNode): boolean => (n.attrs['revision'] as { deletion?: unknown } | null | undefined)?.deletion !== undefined;
  const same = (n: PMNode | null | undefined): boolean => {
    if (!n || n.type !== t.node.type) return false;
    // A row marked deleted with tracking on (`table.deleteRows` marks rather than removes) is removed, as an accept would make it.
    if (t.group !== undefined) return ((n.attrs['sdt'] as SdtGroup[] | null) ?? []).some((g) => g.groupId === t.group) && !(n.type.name === 'table_row' && deletedRow(n) && !deletedRow(t.node));
    const id = sdtId(t.node.attrs['sdtPr']);
    return id === undefined || sdtId(n.attrs['sdtPr']) === id;
  };
  const pos = tr.mapping.map(t.pos, 1);
  const at = tr.doc.nodeAt(pos);
  if (same(at)) return { pos, node: at! };
  // A step that rebuilt the block around it (a paragraph replaced by its edited copy): the same control elsewhere.
  if (t.group === undefined && sdtId(t.node.attrs['sdtPr']) === undefined) return undefined;
  let found: { pos: number; node: PMNode } | undefined;
  tr.doc.descendants((n, p) => {
    if (found) return false;
    if (same(n)) { found = { pos: p, node: n }; return false; }
    return true;
  });
  return found;
}

/**
 * Why a local transaction is refused by a content control's lock, or undefined when it may go ahead:
 * a control that cannot be deleted (`sdtLocked`, `sdtContentLocked`) it removes, a row's or a cell's
 * included, or a control whose contents cannot be edited (`contentLocked`, `sdtContentLocked`) whose
 * content it changes, formatting and tracked changes included. A control whose contents alone are
 * locked may still be deleted whole, as in Word.
 */
export function controlLockRefusal(tr: Transaction, state: EditorState): LockRefusal | undefined {
  if (exempt(tr)) return undefined;
  for (const t of touchedControls(tr, state)) {
    if (t.sdtPr === undefined || t.sdtPr === null) continue;
    const { cannotDelete, cannotEdit } = readControl(t.sdtPr);
    if (!cannotDelete && !cannotEdit) continue;
    const after = survivor(tr, t);
    if (!after) {
      if (cannotDelete) return 'cannotDelete';
      continue;
    }
    if (cannotEdit && !after.node.content.eq(t.node.content)) return 'cannotEdit';
  }
  return undefined;
}

/** The ranges, in the transaction's document, that its first `steps` steps inserted. */
function insertedRanges(tr: Transaction, steps: number): [number, number][] {
  const out: [number, number][] = [];
  for (let i = 0; i < steps; i++) {
    const rest = tr.mapping.slice(i + 1);
    tr.mapping.maps[i]!.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      if (newEnd > newStart) out.push([rest.map(newStart, 1), rest.map(newEnd, -1)]);
    });
  }
  return out.filter(([a, b]) => b > a);
}

/** Whether a text node is the placeholder's: in Word's `PlaceholderText` character style. */
const isPlaceholderText = (n: PMNode): boolean => n.isText && n.marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === PLACEHOLDER_STYLE);

/**
 * A local transaction, recorded, with the controls' own rules applied after it (the transaction
 * itself, steps appended): in a control showing its placeholder, what the first `steps` steps (the
 * edit before the tracker's) inserted replaces the placeholder text, loses the placeholder's style,
 * and the control no longer shows its placeholder; a control to remove when edited (`w:temporary`)
 * whose content the edit changed is removed, its content kept. A row's or a cell's group is left.
 */
export function applyControlRules(tr: Transaction, state: EditorState, steps: number): Transaction {
  if (exempt(tr)) return tr;
  const touched = touchedControls(tr, state, steps).filter((t) => t.group === undefined && t.sdtPr);
  const placeholders = touched.filter((t) => readControl(t.sdtPr).showingPlaceholder);
  const temporary = touched.filter((t) => readControl(t.sdtPr).removeWhenEdited);
  if (placeholders.length === 0 && temporary.length === 0) return tr;
  const schema = tr.doc.type.schema;
  const placeholderMark = schema.marks['runStyle']!.create({ id: PLACEHOLDER_STYLE });
  const inserted = insertedRanges(tr, steps);
  // From the last control to the first, so that one rewritten does not move those before it.
  for (const t of placeholders.sort((a, b) => b.pos - a.pos)) {
    const after = survivor(tr, t);
    if (!after) continue;
    const start = after.pos + 1;
    const end = after.pos + after.node.nodeSize - 1;
    const mine = inserted.map(([a, b]) => [Math.max(a, start), Math.min(b, end)] as const).filter(([a, b]) => b > a);
    if (mine.length === 0) continue;
    for (const [a, b] of mine) tr.removeMark(a, b, placeholderMark);
    const placeholderText: [number, number][] = [];
    tr.doc.nodeAt(after.pos)!.descendants((n, p) => {
      if (isPlaceholderText(n)) placeholderText.push([after.pos + 1 + p, after.pos + 1 + p + n.nodeSize]);
      return true;
    });
    for (const [a, b] of placeholderText.reverse()) tr.delete(a, b);
    const node = tr.doc.nodeAt(after.pos)!;
    tr.setNodeMarkup(after.pos, undefined, { ...node.attrs, sdtPr: applyControlProperties(node.attrs['sdtPr'], { showingPlaceholder: false }) });
  }
  for (const t of temporary.sort((a, b) => b.pos - a.pos)) {
    const after = survivor(tr, t);
    if (!after || after.node.content.eq(t.node.content)) continue;
    tr.replaceWith(after.pos, after.pos + after.node.nodeSize, after.node.content);
  }
  return tr;
}
