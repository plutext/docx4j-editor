// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tracked changes as commands (ED-005 section 5.4, step C3): the document's tracking mode, its
// settings part's `w:trackRevisions`, turned on and off. A part write, so its half is `prepare`,
// with no undo, as every settings write; the transaction `run` dispatches changes nothing but
// tells the view's listeners (the toolbar) that the mode changed. Accept and reject (step C4,
// section 5.5): a change, the changes a selection touches, an author's, or all, each command one
// transaction and so one undo step, never itself tracked (`UNTRACKED_META`). The revisions shown in
// Word only are the engine's to resolve (core-ts 0.3.0, `preserved.mts`): a command that takes them
// (all, an author's, one entry's) resolves them in its `settle` half, a tree run after its
// transaction, so every resolution still goes through these four commands (ED-005 section 12.26).
import type { EditorState } from 'prosemirror-state';
import { register, sessionOf } from './spec.mjs';
import { UNTRACKED_META } from '../track.mjs';
import { trackedChangesOf, emptyChangesOf, resolveChanges, changeOfKey, liveMapping, preservedChangesOf, preservedGroupsOf, type TrackedChangeInfo, type ResolveContext } from '../review.mjs';
import { preservedTargetOf, type PreservedTarget } from '../preserved.mjs';
import { resolutionRecord, addResolutionWrite } from '../resolutions.mjs';
import { numberingOf } from './lists.mjs';
import { dropCommentsResolvedAway, startedOutsideRevision } from '../comments.mjs';

/** The meta of the transaction a mode change dispatches: the document is the same, the mode is not. */
export const TRACKING_CHANGED = 'docx4j-tracking-changed';

export type TrackingModeArg = 'TrackAll' | 'Off' | 'toggle';

register<{ value: TrackingModeArg }>({
  name: 'document.setChangeTrackingMode',
  label: 'Track changes',
  args: { value: "'TrackAll', 'Off', or 'toggle'" },
  choices: { value: ['TrackAll', 'Off', 'toggle'] },
  tracking: 'untracked',
  prepare: async (state, args) => {
    const session = sessionOf(state);
    if (!session) return args;
    const mode = args.value === 'toggle' ? (session.trackingMode === 'Off' ? 'TrackAll' : 'Off') : args.value;
    if (mode !== session.trackingMode) await session.setChangeTrackingMode(mode);
    return { value: mode };
  },
  run: () => (state, dispatch) => {
    dispatch?.(state.tr.setMeta(TRACKING_CHANGED, true));
    return true;
  },
  isActive: (state) => sessionOf(state)?.trackingMode === 'TrackAll',
});

/** The meta of an accept or reject transaction: how many changes it resolved (the status line's count). */
export const RESOLVED_META = 'docx4j-resolved';

/** Which changes accept or reject acts on: the one a key names, else those a range touches (the selection by default); an author narrows either. */
export interface ResolveArgs {
  /** A change's key (`EditorDocument.changeKey`), which survives remote edits around it. */
  key?: string;
  /** Several changes' keys, resolved in one transaction (a replacement's two halves, ED-005 section 9.7): those changes and nothing else in their span. */
  keys?: string[];
  from?: number;
  to?: number;
  author?: string;
  /** A revision shown in Word only (a panel entry, `preservedTargetOf`), resolved by the engine: nothing else. */
  preserved?: PreservedTarget;
}

/** The revisions shown in Word only a command takes: all (with no author, those the engine never lists too), or an author's the engine lists. */
function preservedFor(state: EditorState, author: string | undefined): number {
  const session = sessionOf(state);
  if (!session) return 0;
  const records = preservedChangesOf(state.doc, (id) => session.ctx.ids.get(id)?.value);
  if (author === undefined) return records.length;
  const groups = preservedGroupsOf(state.doc, records);
  return groups.filter((g) => (g.author ?? '') === author && preservedTargetOf(state.doc, g, groups) !== undefined).length;
}

/** Whether an entry's revision shown in Word only is still there to resolve. */
function hasPreserved(state: EditorState, target: PreservedTarget): boolean {
  const session = sessionOf(state);
  if (!session) return false;
  const groups = preservedGroupsOf(state.doc, preservedChangesOf(state.doc, (id) => session.ctx.ids.get(id)?.value));
  return groups.some((g) => {
    const t = preservedTargetOf(state.doc, g, groups);
    return t !== undefined && t.block === target.block && t.kind === target.kind && t.ordinal === target.ordinal;
  });
}

/** The changes a command acts on (a caret touches the change it is in or at the edge of). */
function changesFor(state: EditorState, args: ResolveArgs | undefined, all: boolean): TrackedChangeInfo[] {
  // `''` is the unknown author, as the panel groups it: the listing has that author as `null`.
  const byAuthor = (c: TrackedChangeInfo): boolean => args?.author === undefined || (c.author ?? '') === args.author;
  if (args?.preserved !== undefined) return [];
  const changes = trackedChangesOf(state.doc).filter(byAuthor);
  if (all) return [...changes, ...emptyChangesOf(state.doc).filter(byAuthor)];
  if (args?.key !== undefined || args?.keys !== undefined) {
    const session = sessionOf(state);
    if (!session) return [];
    // Through the view's live map when the state has one (ED-005 section 6.5): a map rebuilt from the store costs the whole document.
    const live = liveMapping(state, session.store.body);
    const found: TrackedChangeInfo[] = [];
    for (const key of args.keys ?? [args.key!]) {
      const change = changeOfKey(session.store.body, state.schema, state.doc, key, live);
      if (change && byAuthor(change) && !found.some((c) => c.from === change.from && c.type === change.type && c.author === change.author && c.move === change.move)) found.push(change);
    }
    return found.sort((a, b) => a.from - b.from);
  }
  const from = args?.from ?? state.selection.from;
  const to = args?.to ?? state.selection.to;
  return changes.filter((c) => (from === to ? c.from <= from && from <= c.to : c.from < to && c.to > from));
}

function resolveCommand(action: 'accept' | 'reject', all: boolean) {
  return (args: ResolveArgs | undefined) => (state: EditorState, dispatch?: (tr: import('prosemirror-state').Transaction) => void): boolean => {
    const changes = changesFor(state, args, all);
    // Only revisions shown in Word only: nothing for the transaction, the settle half resolves them.
    if (changes.length === 0) return args?.preserved !== undefined ? hasPreserved(state, args.preserved) : all && preservedFor(state, args?.author) > 0;
    if (dispatch) {
      const session = sessionOf(state);
      const numbering = numberingOf(state);
      const context: ResolveContext = {
        element: (id) => session?.ctx.ids.get(id)?.value as { name?: unknown; id?: unknown } | undefined,
        listIndent: (numId, ilvl) => {
          const ind = numbering?.definitions.indOf(numId, String(ilvl));
          return ind ? { start: ind.start ?? ind.left ?? null, end: ind.end ?? ind.right ?? null, firstLine: ind.firstLine ?? null, hanging: ind.hanging ?? null } : null;
        },
      };
      const tr = state.tr;
      // In a room, what it resolves is recorded with the state it saw, for the room repair (ED-005 section 12.15): a move's other half too, which it resolves with it.
      if (session?.inRoom) {
        const moves = new Set(changes.map((c) => c.move).filter((m) => m !== undefined));
        const all = moves.size > 0 ? [...changes, ...trackedChangesOf(state.doc).filter((c) => c.move !== undefined && moves.has(c.move) && !changes.includes(c))] : changes;
        const record = resolutionRecord(session.store.body, state.schema, all, action, liveMapping(state, session.store.body));
        if (record) addResolutionWrite(tr, record);
      }
      const count = resolveChanges(tr, changes, action, context);
      // A comment whose text the resolution took goes with it, as Word removes it, unless the file put its
      // start marker outside the revision (check 32 and Word 2010 and 15's saves, ED-005 section 12.49).
      if (session) dropCommentsResolvedAway(state.doc, tr, session.ctx.ids, (id) => startedOutsideRevision(session.ctx.paragraphs, id));
      dispatch(tr.setMeta(UNTRACKED_META, true).setMeta(RESOLVED_META, count).scrollIntoView());
    }
    return true;
  };
}

const RESOLVE_ARGS = { key: "a change's key, or nothing", keys: "several changes' keys, resolved together, or nothing", from: 'a position, or the selection', to: 'a position, or the selection', author: 'only this author\'s changes, or everyone\'s', preserved: 'a revision shown in Word only (the review panel\'s), or nothing' };

/** One entry's revision shown in Word only, by the engine, after the (empty) transaction. */
const settleOne = (action: 'accept' | 'reject') => (state: EditorState, args: ResolveArgs | undefined): (() => Promise<number>) | undefined => {
  const session = sessionOf(state);
  const target = args?.preserved;
  return session && target && hasPreserved(state, target) ? () => session.resolvePreserved(action, { target }) : undefined;
};

/** All (or an author's) revisions shown in Word only, by the engine, after the transaction resolved the rest. */
const settleAll = (action: 'accept' | 'reject') => (state: EditorState, args: { author?: string } | undefined): (() => Promise<number>) | undefined => {
  const session = sessionOf(state);
  const author = args?.author;
  if (!session || preservedFor(state, author) === 0) return undefined;
  return () => session.resolvePreserved(action, author === undefined ? { all: true } : { author });
};

/** The Office JS line over the changes the selection touches (the engine's `Range.getTrackedChanges`); none for a key or an author. */
const rangeLine = (verb: 'accept' | 'reject') => (args: ResolveArgs | undefined): string | undefined =>
  args?.key !== undefined || args?.keys !== undefined || args?.author !== undefined || args?.from !== undefined || args?.preserved !== undefined ? undefined : `for (const change of selection.range.getTrackedChanges()) change.${verb}();`;

register<ResolveArgs | undefined>({
  name: 'trackedChange.accept',
  label: 'Accept change',
  args: RESOLVE_ARGS,
  tracking: 'untracked',
  api: rangeLine('accept'),
  run: resolveCommand('accept', false),
  settle: settleOne('accept'),
  isEnabled: (state, args) => (args?.preserved !== undefined ? hasPreserved(state, args.preserved) : changesFor(state, args, false).length > 0),
});

register<ResolveArgs | undefined>({
  name: 'trackedChange.reject',
  label: 'Reject change',
  args: RESOLVE_ARGS,
  tracking: 'untracked',
  api: rangeLine('reject'),
  run: resolveCommand('reject', false),
  settle: settleOne('reject'),
  isEnabled: (state, args) => (args?.preserved !== undefined ? hasPreserved(state, args.preserved) : changesFor(state, args, false).length > 0),
});

register<{ author?: string } | undefined>({
  name: 'trackedChanges.acceptAll',
  label: 'Accept all changes',
  args: { author: "only this author's changes, or everyone's" },
  tracking: 'untracked',
  api: (args) => (args?.author !== undefined ? undefined : 'body.acceptAll();'),
  run: resolveCommand('accept', true),
  settle: settleAll('accept'),
  isEnabled: (state, args) => changesFor(state, args, true).length > 0 || preservedFor(state, args?.author) > 0,
});

register<{ author?: string } | undefined>({
  name: 'trackedChanges.rejectAll',
  label: 'Reject all changes',
  args: { author: "only this author's changes, or everyone's" },
  tracking: 'untracked',
  api: (args) => (args?.author !== undefined ? undefined : 'body.rejectAll();'),
  run: resolveCommand('reject', true),
  settle: settleAll('reject'),
  isEnabled: (state, args) => changesFor(state, args, true).length > 0 || preservedFor(state, args?.author) > 0,
});
