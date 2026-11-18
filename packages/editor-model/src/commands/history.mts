// SPDX-License-Identifier: GPL-3.0-or-later
//
// Undo and redo (ED-002 section 4.2): the session's `Y.UndoManager`, scoped to the local origin.
// The command touches the store, not the state; the state follows (the view through
// `ySyncPlugin`, the headless runner by reading the store back). ProseMirror's `history` plugin
// is not installed.
import { register, sessionOf } from './spec.mjs';

export const undo = register<void>({
  name: 'history.undo',
  tracking: 'untracked',
  label: 'Undo',
  run: () => (state) => {
    const session = sessionOf(state);
    if (!session) return false;
    return session.undoManager.undo() !== null;
  },
  isEnabled: (state) => (sessionOf(state)?.undoManager.canUndo() ?? false),
});

export const redo = register<void>({
  name: 'history.redo',
  tracking: 'untracked',
  label: 'Redo',
  run: () => (state) => {
    const session = sessionOf(state);
    if (!session) return false;
    return session.undoManager.redo() !== null;
  },
  isEnabled: (state) => (sessionOf(state)?.undoManager.canRedo() ?? false),
});
