// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word's keys (ED-001 section 3.4) as data (ED-002 section 4.8): key name in prosemirror-keymap's
// syntax to command name and arguments. E1.c turns it into a keymap plugin; the palette lists the
// same commands. `Mod` is Ctrl on Windows and Linux and Cmd on a Mac.
export interface KeyBinding {
  key: string;
  command: string;
  args?: unknown;
}

export const WORD_KEYMAP: readonly KeyBinding[] = [
  { key: 'Mod-b', command: 'font.bold', args: { value: 'toggle' } },
  { key: 'Mod-i', command: 'font.italic', args: { value: 'toggle' } },
  { key: 'Mod-u', command: 'font.underline', args: { value: 'toggle' } },
  { key: 'Mod-=', command: 'font.subscript', args: { value: 'toggle' } },
  { key: 'Mod-Shift-=', command: 'font.superscript', args: { value: 'toggle' } },
  { key: 'Mod-+', command: 'font.superscript', args: { value: 'toggle' } },
  // Shifted punctuation reaches the keymap as the shifted character on most layouts, and as the
  // unshifted one with Shift on some; both spellings are bound.
  { key: 'Mod-Shift-.', command: 'font.grow' },
  { key: 'Mod->', command: 'font.grow' },
  { key: 'Mod-Shift-,', command: 'font.shrink' },
  { key: 'Mod-<', command: 'font.shrink' },
  { key: 'Mod-Space', command: 'font.clear' },
  { key: 'Mod-l', command: 'paragraph.alignment', args: { value: 'Left' } },
  { key: 'Mod-e', command: 'paragraph.alignment', args: { value: 'Centered' } },
  { key: 'Mod-r', command: 'paragraph.alignment', args: { value: 'Right' } },
  { key: 'Mod-j', command: 'paragraph.alignment', args: { value: 'Justified' } },
  { key: 'Mod-m', command: 'paragraph.increaseIndent' },
  { key: 'Mod-Shift-m', command: 'paragraph.decreaseIndent' },
  { key: 'Mod-q', command: 'paragraph.clearFormatting' },
  { key: 'Mod-Shift-l', command: 'list.toggle', args: { bullet: true } },
  { key: 'Mod-Alt-1', command: 'paragraph.style', args: { value: 'Heading1' } },
  { key: 'Mod-Alt-2', command: 'paragraph.style', args: { value: 'Heading2' } },
  { key: 'Mod-Alt-3', command: 'paragraph.style', args: { value: 'Heading3' } },
  { key: 'Mod-Shift-n', command: 'paragraph.style', args: { value: null } },
  { key: 'Enter', command: 'insertParagraph' },
  { key: 'Backspace', command: 'joinBackward' },
  { key: 'Delete', command: 'joinForward' },
  { key: 'Shift-Enter', command: 'insertBreak', args: { type: 'Line' } },
  { key: 'Mod-Enter', command: 'insertBreak', args: { type: 'Page' } },
  { key: 'Mod-Shift-Enter', command: 'insertBreak', args: { type: 'Column' } },
  // Several bindings of one key run in order until one handles it: Tab at the start of a list
  // item changes its level (Word), else moves between cells in a table, else inserts a tab.
  { key: 'Tab', command: 'list.indent', args: { atStart: true } },
  { key: 'Shift-Tab', command: 'list.outdent', args: { atStart: true } },
  { key: 'Tab', command: 'table.nextCell' },
  { key: 'Shift-Tab', command: 'table.previousCell' },
  { key: 'Tab', command: 'insertTab' },
  // Word's Alt+Ctrl+F and Alt+Ctrl+D (ED-003 section 6.3).
  { key: 'Alt-Mod-f', command: 'range.insertFootnote' },
  { key: 'Alt-Mod-d', command: 'range.insertEndnote' },
  { key: 'Mod-z', command: 'history.undo' },
  { key: 'Mod-y', command: 'history.redo' },
  { key: 'Mod-Shift-z', command: 'history.redo' },
];
