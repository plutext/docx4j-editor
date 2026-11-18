// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Word keymap (ED-002 section 4.8) as a prosemirror-keymap binding table: every entry is a
// registered command with its arguments.
import type { Command } from 'prosemirror-state';
import { baseKeymap, chainCommands } from 'prosemirror-commands';
import { WORD_KEYMAP, commandNamed, asCommand, inRunControl } from '@docx4j/editor-model';

/** A key bound several times runs its commands in order until one handles it (Tab: the next cell, else a tab). */
export function wordKeymap(): Record<string, Command> {
  const commands = new Map<string, Command[]>();
  for (const binding of WORD_KEYMAP) {
    const spec = commandNamed(binding.command);
    if (!spec) throw new Error(`keymap binds ${binding.key} to an unknown command ${binding.command}`);
    const list = commands.get(binding.key) ?? [];
    list.push(asCommand(spec, binding.args));
    commands.set(binding.key, list);
  }
  const out: Record<string, Command> = {};
  for (const [key, list] of commands) out[key] = list.length === 1 ? list[0]! : chainCommands(...list);
  return out;
}

/**
 * ProseMirror's base keymap, after the Word keymap on every surface, its Enter never splitting a run-level control (ED-005
 * section 12.123). `insertParagraph` splits outside a control or refuses in its middle, and the base keymap's `splitBlock`
 * would then split the control, the copy in the new paragraph carrying its `w:id`, tag and binding.
 */
export function editorBaseKeymap(): Record<string, Command> {
  const enter = baseKeymap['Enter']!;
  return { ...baseKeymap, Enter: (state, dispatch, view) => (inRunControl(state.selection.$from) || inRunControl(state.selection.$to) ? true : enter(state, dispatch, view)) };
}
