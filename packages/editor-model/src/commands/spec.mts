// SPDX-License-Identifier: GPL-3.0-or-later
//
// The shape of a command (ED-002 section 4.1): registered once under its content-API name, it
// produces a ProseMirror `Command`, the one shape a keymap, a menu, the palette, the console and
// the headless runner all consume. The session a command may need (the source elements, the
// resolver, the Yjs store) travels in the editor state as a plugin, so a command reads it from
// `state` and nothing outside the view ever sees a position.
import { Plugin, PluginKey, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorDocument } from '../session.mjs';
import type { Story } from '../stories.mjs';

/**
 * The spans a selection covers, first to last: one for a text selection; one per cell for a cell selection, whose
 * `from` and `to` prosemirror-tables sets to the head cell's content alone (its `ranges` are the cells). Every command
 * and reading that walks the selection walks these, so a table or some of its cells selected is formatted and read
 * cell by cell (2026-10-15).
 */
export function selectionSpans(state: EditorState): { from: number; to: number }[] {
  return state.selection.ranges.map((r) => ({ from: r.$from.pos, to: r.$to.pos })).sort((a, b) => a.from - b.from);
}

export interface CommandSpec<A = void> {
  /** The content-API name: `font.bold`, `paragraph.alignment`, `insertBreak`, `section.orientation`. */
  name: string;
  /** For the palette and the toolbar's tooltip. */
  label: string;
  /** The arguments, described for people and agents: `{ value: "boolean | 'toggle'" }`. */
  args?: Record<string, string>;
  /**
   * The values an enumerated argument takes, as the palette lists them in a drop-down (Jason,
   * 2026-09-25): the strings the command accepts, `'true'` and `'false'` for a boolean, in the
   * order to show. An argument without an entry is typed as text.
   */
  choices?: Record<string, readonly string[]>;
  /** Typing-like: coalesces with neighbouring keystrokes in the undo history. */
  typing?: boolean;
  /**
   * With tracking on (ED-005 section 5.4, proposal 5): `untracked` for what the model cannot
   * express as a revision (a column, a merge or split, a content control, a section, a table's
   * properties, a part write), which then runs untracked and says so in the status line, as Word
   * mostly does. Tracked when absent.
   */
  tracking?: 'tracked' | 'untracked';
  /**
   * A definition command's asynchronous half (ED-003 section 5.3): the part write it needs
   * before its transaction (a `w:num` made for a list started), done against the session and
   * answering the arguments `run` then takes, the write's result among them. A part write is
   * not undone (section 9 item 7); `run` is the document's half, one transaction. The view and
   * `runCommandAsync` await it; the synchronous runner refuses a command that has one.
   */
  prepare?(state: EditorState, args: A): Promise<A>;
  /**
   * The asynchronous half after the transaction (ED-005 section 12.26): what only the engine can
   * resolve (the revisions shown in Word only, `preserved.mts`), done in a tree run over the
   * document the transaction left. Asked of the state before the transaction, it answers the work
   * or undefined when there is none. The view and `runCommandAsync` run it; the synchronous runner
   * refuses a command whose settle has work.
   */
  settle?(state: EditorState, args: A): (() => Promise<unknown>) | undefined;
  run(args: A): Command;
  /**
   * The portable form (ED-003 section 7.2, section 9 item 20): the content-API line or lines
   * that do what `run` does, over the console's `selection` and `body` bindings, which core-ts
   * runs anywhere (the console, Node, an add-in through the `Word` shim, E4's tools). Absent for
   * a command that only the edit layer can express; undefined for arguments the content API has
   * no form for (a cleared value). `api.test.mts` proves each line against its command.
   */
  api?(args: A): string | undefined;
  isEnabled?(state: EditorState, args?: A): boolean;
  isActive?(state: EditorState, args?: A): boolean;
}

/** What a boolean-or-toggle argument takes, for the palette's drop-down. */
export const TOGGLE_CHOICES: readonly string[] = ['true', 'false', 'toggle'];

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const registry = new Map<string, CommandSpec<any>>();

export function register<A>(spec: CommandSpec<A>): CommandSpec<A> {
  if (registry.has(spec.name)) throw new Error(`command ${spec.name} registered twice`);
  registry.set(spec.name, spec);
  return spec;
}

/** The meta a command's transaction carries: its name, so the tracker knows whether it is tracked (ED-005 section 5.4). */
export const COMMAND_META = 'docx4j-command';

/**
 * The meta a paragraph split (Enter) carries: the positions, in its transaction's document, of the
 * paragraphs whose mark it made. Word puts the new mark at the caret, so the paragraph before the
 * caret ends with it and the one after keeps the old mark, at a paragraph's end too (core-ts check
 * 29). A mark before an empty one, and one before a mark like it, are the same tokens to the
 * changeset, which could take either; the tracker reads which from here (ED-005 section 12.17).
 */
export const NEW_MARK_META = 'docx4j-new-mark';

/** A dispatch that tags each transaction with the command that made it. */
export function taggedDispatch(name: string, dispatch: (tr: Transaction) => void): (tr: Transaction) => void {
  return (tr) => dispatch(tr.setMeta(COMMAND_META, name));
}

/**
 * A spec as one ProseMirror command for a view: a command with a `prepare` half runs it against
 * the view's state and dispatches `run` when it resolves (true meaning it was started), which is
 * what the keymap and the view's `run` need; one without is its `run`. Either way its
 * transactions carry its name (`COMMAND_META`).
 */
export function asCommand<A>(spec: CommandSpec<A>, args: A): Command {
  const settled = (work: (() => Promise<unknown>) | undefined): void => {
    if (work) void work().catch((e: unknown) => console.error(`${spec.name}: its settle half failed`, e));
  };
  if (!spec.prepare) {
    const run = spec.run(args);
    if (!spec.settle) return (state, dispatch, view) => run(state, dispatch && taggedDispatch(spec.name, dispatch), view);
    return (state, dispatch, view) => {
      const work = dispatch ? spec.settle!(state, args) : undefined;
      const ok = run(state, dispatch && taggedDispatch(spec.name, dispatch), view);
      if (ok) settled(work);
      return ok;
    };
  }
  const prepare = spec.prepare.bind(spec);
  return (state, dispatch, view) => {
    if (spec.isEnabled && !spec.isEnabled(state, args)) return false;
    if (!dispatch || !view) return false;
    void prepare(state, args).then((prepared) => {
      if ((view as { isDestroyed?: boolean }).isDestroyed) return;
      const work = spec.settle?.(view.state, prepared);
      if (spec.run(prepared)(view.state, taggedDispatch(spec.name, view.dispatch.bind(view)), view)) settled(work);
    });
    return true;
  };
}

export function commandNamed(name: string): CommandSpec<unknown> | undefined {
  return registry.get(name);
}

export function listCommands(): CommandSpec<unknown>[] {
  return [...registry.values()];
}

export const sessionKey = new PluginKey<EditorDocument>('docx4j-session');

/** Carries the session in the editor state; the view and the headless runner both install it. */
export function sessionPlugin(session: EditorDocument): Plugin<EditorDocument> {
  return new Plugin<EditorDocument>({
    key: sessionKey,
    state: { init: () => session, apply: (_tr, value) => value },
  });
}

export function sessionOf(state: EditorState): EditorDocument | undefined {
  return sessionKey.getState(state);
}

export const storyKey = new PluginKey<Story>('docx4j-story');

/**
 * Carries the story an editor state edits (a note, a header; ED-003 section 6.4), so a command
 * that only makes sense in the body (a section property, a page break, a note inserted) can
 * refuse. The body's state has no such plugin.
 */
export function storyPlugin(story: Story): Plugin<Story> {
  return new Plugin<Story>({
    key: storyKey,
    state: { init: () => story, apply: (_tr, value) => value },
  });
}

/** The story a state edits, or undefined for the body. */
export function storyOf(state: EditorState): Story | undefined {
  return storyKey.getState(state);
}
