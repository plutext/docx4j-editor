// SPDX-License-Identifier: GPL-3.0-or-later
//
// The headless runner (ED-002 section 4.1): a command by name against a session, with no view.
// Tests, E1.d's console and E4's tools use it; the view runs the same specs through its own state.
import { EditorState, Selection, TextSelection, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorDocument } from '../session.mjs';
import type { Story } from '../stories.mjs';
import { LOCAL_ORIGIN, storeWritesOf } from '../yjs.mjs';
import { commandNamed, sessionPlugin, storyPlugin, COMMAND_META } from './spec.mjs';
import { syncBoundAfter } from '../bound.mjs';
import { endCellsWithParagraph } from './tables.mjs';
import { recordTransaction } from '../track.mjs';
import { controlLockRefusal, applyControlRules, type LockRefusal } from '../controlrules.mjs';
import { resolutionWriteOf } from '../resolutions.mjs';

export type SelectionSpec = Selection | { from: number; to?: number };

/** An editor state over the session's current document, or a story's (with the story marker the commands read), with the session plugin installed. */
export function stateFor(session: EditorDocument, selection?: SelectionSpec, doc?: PMNode, story?: Story): EditorState {
  const over = doc ?? (story ? session.stories.currentDoc(story) : session.doc);
  const sel = selection === undefined ? undefined
    : selection instanceof Selection ? selection
    : TextSelection.create(over, selection.from, selection.to ?? selection.from);
  return EditorState.create({ doc: over, selection: sel, plugins: [sessionPlugin(session), ...(story ? [storyPlugin(story)] : [])] });
}

export interface RunResult {
  /** What the command returned: false means it did not apply. */
  ok: boolean;
  /** The state after the command, when it dispatched. */
  state?: EditorState;
  /** With tracking on, set when the command ran untracked: its name, or `story` (ED-005 section 5.4). */
  untracked?: string;
  /** Set when a content control's lock refused the edit (ED-005 section 12.60); `ok` is then false. */
  refused?: LockRefusal;
}

/**
 * Run a registered command: build a state, run it, record its transaction as the view does
 * (tracked with tracking on; ED-005 section 5.4), and push the resulting document into the store
 * as one local transaction. A command that is not typing is its own undo step. Over a story (a
 * note, a header; ED-003 section 6.2) when one is given, through its fragment.
 */
export function runCommand(session: EditorDocument, name: string, args?: unknown, selection?: SelectionSpec, story?: Story): RunResult {
  const spec = commandNamed(name);
  if (!spec) throw new Error(`no command named ${name}`);
  if (spec.prepare) throw new Error(`${name} has an asynchronous half: use runCommandAsync`);
  if (spec.settle?.(stateFor(session, selection, undefined, story), args)) throw new Error(`${name} has work for its asynchronous half here: use runCommandAsync`);
  return apply(session, spec, args, selection, story);
}

/** `runCommand` for a command with a `prepare` half (a list started or restarted): the part write first, then the transaction. */
export async function runCommandAsync(session: EditorDocument, name: string, args?: unknown, selection?: SelectionSpec, story?: Story): Promise<RunResult> {
  const spec = commandNamed(name);
  if (!spec) throw new Error(`no command named ${name}`);
  const prepared = spec.prepare ? await spec.prepare(stateFor(session, selection, undefined, story), args) : args;
  const work = spec.settle?.(stateFor(session, selection, undefined, story), prepared);
  const result = apply(session, spec, prepared, selection, story);
  if (result.ok && work) await work();
  return result;
}

function apply(session: EditorDocument, spec: NonNullable<ReturnType<typeof commandNamed>>, args: unknown, selection: SelectionSpec | undefined, story: Story | undefined): RunResult {
  const state = stateFor(session, selection, undefined, story);
  let next: EditorState | undefined;
  let applied: Transaction | undefined;
  let untracked: string | undefined;
  let refused: LockRefusal | undefined;
  const ok = spec.run(args)(state, (tr) => {
    // A content control's lock refuses the edit; its placeholder and its remove-when-edited act after the tracker (ED-005 section 12.60).
    tr.setMeta(COMMAND_META, spec.name);
    refused = controlLockRefusal(tr, state);
    if (refused) return;
    // A cell's last block stays a paragraph (ED-005 section 12.109), in the command's own transaction: one undo step, tracked with it.
    endCellsWithParagraph(tr);
    const steps = tr.steps.length;
    const recorded = recordTransaction(tr, state);
    applied = applyControlRules(recorded.tr, state, steps);
    untracked = recorded.untracked;
    next = state.apply(applied);
  });
  if (ok && next && applied && (next.doc !== state.doc || storeWritesOf(applied).length > 0)) {
    if (!spec.typing) session.undoManager.stopCapturing();
    // The control-map writes a command asked for go into the same Yjs transaction: one undo step.
    if (story) session.replaceStory(story, next.doc);
    else {
      session.replace(next.doc, LOCAL_ORIGIN, storeWritesOf(applied), resolutionWriteOf(applied), state.doc);
      // Word's two-way mapping (ED-005 section 12.55): an edit inside a bound control writes its node.
      const after = next;
      syncBoundAfter(session, applied, () => after);
    }
  }
  if (refused) return { ok: false, refused };
  return next ? { ok, state: next, ...(untracked ? { untracked } : {}) } : { ok };
}
