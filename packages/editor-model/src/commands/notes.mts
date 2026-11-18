// SPDX-License-Identifier: GPL-3.0-or-later
//
// Footnotes and endnotes inserted and deleted (ED-003 section 6.3; ED-001 section 3.2), in
// Office JS's shapes: `range.insertFootnote(text)`, `range.insertEndnote(text)`, `noteItem.delete()`.
// `prepare` does the part writes (the notes part made when the document has none, the two styles
// spliced from docx4j's defaults, neither undone) and makes the note in the side map; `run` puts
// the reference at the caret as an atom whose elements (the `w:footnoteReference` and its run,
// the caret run's properties plus `w:rStyle FootnoteReference`) are built here and registered in
// the session's ids, so export's identity path serves them as it serves an imported reference.
// A deleted reference takes its note with it at save (section 6.2); undo brings both back.
import { NodeSelection, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { deepCopy, runOf, type Element } from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { register, sessionOf, storyOf } from './spec.mjs';
import { NOTE_STYLES, type NoteKind } from '../stories.mjs';
import type { EditorDocument } from '../session.mjs';

export { NOTE_STYLES };

export interface InsertNoteArgs {
  /** Text for the new note, after its reference mark; nothing for an empty note with the caret in it. */
  text?: string;
  /** Filled by `prepare`: the note made in the side map. */
  id?: number;
  /** Filled by `prepare`: the reference's element ids, its run already in `sources` (ED-004 section 3.7). */
  reference?: { sourceId: string; runSource: string };
}

/** The transaction meta naming the note whose view should take the caret once it mounts (`footnote:7`). */
export const NOTE_FOCUS = 'docx4j-note-focus';

export function noteFocusOf(tr: Transaction): string | undefined {
  return tr.getMeta(NOTE_FOCUS) as string | undefined;
}

const inBodyParagraph = (state: EditorState): boolean => {
  const { $from, $to } = state.selection;
  return storyOf(state) === undefined && $from.parent.type.name === 'paragraph' && $from.sameParent($to);
};

/** The properties of the run the caret sits in (the text or atom before it, else after), for the reference run to copy. */
function caretRunProperties(state: EditorState, session: EditorDocument): wml.RPr | undefined {
  const $pos = state.selection.$to;
  for (const n of [$pos.nodeBefore, $pos.nodeAfter]) {
    if (!n) continue;
    const id = n.isText ? n.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] : n.attrs['runSource'];
    if (typeof id !== 'string') continue;
    const run = session.ctx.ids.get(id) as Element<wml.R> | undefined;
    if (run?.value.rPr) return run.value.rPr;
  }
  return undefined;
}

/**
 * The reference's elements, built and registered: a `w:footnoteReference` (or endnote) with the
 * id, in a `w:r` whose `rPr` is a copy of the caret run's with `w:rStyle` set to the reference
 * style, as Word writes it. The atom carries both ids; export emits them by identity.
 */
export function newReference(session: EditorDocument, kind: NoteKind, id: number, base: wml.RPr | undefined): { sourceId: string; runSource: string; run: Element } {
  const rPr: wml.RPr = base ? deepCopy(base) : { TYPE_NAME: 'org_docx4j_wml.RPr' };
  rPr.rStyle = { val: NOTE_STYLES[kind].reference };
  const reference = (kind === 'footnote' ? el.footnoteReference({ id }) : el.endnoteReference({ id })) as Element;
  const run = runOf([reference], rPr);
  const runSource = session.ctx.ids.idOf(run);
  session.ctx.sources.set(runSource, { holders: [] });
  return { sourceId: session.ctx.ids.idOf(reference), runSource, run };
}

function insertNote(kind: NoteKind) {
  return register<InsertNoteArgs | undefined>({
    name: kind === 'footnote' ? 'range.insertFootnote' : 'range.insertEndnote',
    label: kind === 'footnote' ? 'Insert footnote' : 'Insert endnote',
    args: { text: 'text for the new note, or nothing' },
    prepare: async (state, args) => {
      const session = sessionOf(state);
      if (!session || !inBodyParagraph(state)) return args;
      await session.ensureNotesPart(kind);
      await session.ensureStyles([NOTE_STYLES[kind].text, NOTE_STYLES[kind].reference]);
      const note = await session.stories.newNote(kind, args?.text ?? '');
      // The reference's run, which is not in the base file, into `sources` before the transaction
      // that inserts it, so another session resolves the atom (ED-004 section 3.7).
      const { run, ...reference } = newReference(session, kind, note.id, caretRunProperties(state, session));
      await session.writeInlineSource(run);
      return { ...(args ?? {}), id: note.id, reference };
    },
    run: (args) => (state, dispatch) => {
      const session = sessionOf(state);
      if (!session || !inBodyParagraph(state) || args?.id === undefined) return false;   // the part half did not run
      const tr = state.tr;
      // A selection collapses to its end, as Word's does.
      tr.setSelection(TextSelection.create(tr.doc, state.selection.to));
      const ids = args.reference ?? (({ sourceId, runSource }) => ({ sourceId, runSource }))(newReference(session, kind, args.id, caretRunProperties(state, session)));
      const node = state.schema.nodes['footnoteRef']!.create({ ...ids, footnoteId: args.id, kind });
      tr.replaceSelectionWith(node, false);
      tr.setSelection(TextSelection.create(tr.doc, state.selection.to + node.nodeSize));
      tr.setMeta(NOTE_FOCUS, `${kind}:${args.id}`);
      dispatch?.(tr.scrollIntoView());
      return true;
    },
    isEnabled: (state) => inBodyParagraph(state) && sessionOf(state) !== undefined,
  });
}

insertNote('footnote');
insertNote('endnote');

export interface DeleteNoteArgs {
  kind?: NoteKind;
  id?: number;
}

/** The reference the command deletes: the one given, else the selected atom, else the one just before the caret. */
function referenceTarget(state: EditorState, args: DeleteNoteArgs | undefined): { pos: number; node: PMNode } | undefined {
  if (storyOf(state) !== undefined) return undefined;
  if (args?.id !== undefined) {
    let found: { pos: number; node: PMNode } | undefined;
    state.doc.descendants((node, pos) => {
      if (found) return false;
      if (node.type.name === 'footnoteRef' && Number(node.attrs['footnoteId']) === args.id && (args.kind === undefined || node.attrs['kind'] === args.kind)) found = { pos, node };
      return !found && !node.isLeaf;
    });
    return found;
  }
  const sel = state.selection;
  if (sel instanceof NodeSelection && sel.node.type.name === 'footnoteRef') return { pos: sel.from, node: sel.node };
  const before = sel.empty ? sel.$from.nodeBefore : null;
  if (before && before.type.name === 'footnoteRef') return { pos: sel.from - before.nodeSize, node: before };
  return undefined;
}

register<DeleteNoteArgs | undefined>({
  name: 'note.delete',
  label: 'Delete note',
  args: { kind: "'footnote' | 'endnote', with id; or nothing for the reference at the caret", id: "the note's id, or nothing" },
  choices: { kind: ['footnote', 'endnote'] },
  run: (args) => (state, dispatch) => {
    const target = referenceTarget(state, args);
    if (!target) return false;
    dispatch?.(state.tr.delete(target.pos, target.pos + target.node.nodeSize).scrollIntoView());
    return true;
  },
  isEnabled: (state, args) => referenceTarget(state, args) !== undefined,
});
