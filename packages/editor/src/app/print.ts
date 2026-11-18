// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Print command (ED-006 section 3.6): Ctrl+P, the app bar's Print and the palette's. It readies the
// document for the browser's print dialog and opens it: the fonts loaded and the images decoded (a page
// printed before them prints blanks), the notes laid out after the body when the document has any
// (section 3.6 step 3, `notes`), the page rules written (the view's `preparePrint`, which `beforeprint`
// also runs), then `print`. What is printed is the document alone (`print.css`); the browser lays out
// the pages. The status line says so once the dialog closes; whether the person printed or cancelled,
// the dialog does not tell.
import type { EditorDocument, Note } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import { StoryView } from '../view/StoryView.js';
import { decorationsKey } from '../view/plugins/decorations.js';
import { t } from './i18n.js';

export interface PrintOptions {
  /** The status line. */
  onStatus: (text: string) => void;
  /**
   * What the print needs laid out before the dialog opens and taken away after it (the notes after the
   * body), answering how to take it away.
   */
  prepare?: () => Promise<(() => void) | undefined>;
}

/** Ready the document and open the browser's print dialog. */
export async function printDocument(editor: DocxEditorView, options: PrintOptions): Promise<void> {
  const doc = editor.view.dom.ownerDocument;
  const win = doc.defaultView;
  if (!win) return;
  options.onStatus(t('status.print.preparing'));
  let undo: (() => void) | undefined;
  try {
    await doc.fonts?.ready;
    await Promise.all([...editor.view.dom.querySelectorAll('img')].map((img) => img.decode().catch(() => { /* a broken image prints as it shows */ })));
    undo = await options.prepare?.();
  } catch (e) {
    undo?.();
    options.onStatus(t('status.print.failed', { error: (e as Error).message }));
    return;
  }
  win.addEventListener('afterprint', () => {
    undo?.();
    options.onStatus(t('status.print.done'));
  }, { once: true });
  editor.preparePrint();
  win.print();
}

/** Whether a key event is Print (Ctrl+P, or ⌘P on a Mac), with no other modifier. */
export function isPrintKey(e: KeyboardEvent): boolean {
  return (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'p';
}

// ---------------------------------------------------------------- the notes after the body (section 3.6 step 3)

/**
 * The document's footnotes, then its endnotes, laid out after the body for the print: each in order of
 * its first reference, labelled as the text labels it (`noteLabels`), its story read-only. A note no
 * reference points at is not printed. Nothing when the document has no notes; else how to take them away.
 */
export async function prepareNotes(editor: DocxEditorView, session: EditorDocument): Promise<(() => void) | undefined> {
  const [footnotes, endnotes] = await Promise.all([session.stories.footnotes(), session.stories.endnotes()]);
  const labels = decorationsKey.getState(editor.state)?.noteLabels;
  const order = new Map<string, number>();
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'footnoteRef') {
      const key = `${String(node.attrs['kind'])}:${String(node.attrs['footnoteId'])}`;
      if (!order.has(key)) order.set(key, order.size);
    }
    return true;
  });
  const printed = (notes: readonly Note[]): Note[] => notes.filter((n) => order.has(n.key)).sort((a, b) => order.get(a.key)! - order.get(b.key)!);
  const groups = [{ kind: 'footnote' as const, notes: printed(footnotes) }, { kind: 'endnote' as const, notes: printed(endnotes) }].filter((g) => g.notes.length > 0);
  const mount = editor.view.dom.parentElement;
  if (groups.length === 0 || !mount) return undefined;
  const doc = mount.ownerDocument;
  const section = doc.createElement('section');
  section.className = 'print-notes';
  section.setAttribute('aria-hidden', 'true');
  const views: StoryView[] = [];
  for (const group of groups) {
    const heading = doc.createElement('h2');
    heading.className = 'print-notes-heading';
    heading.textContent = t(group.kind === 'footnote' ? 'print.footnotes' : 'print.endnotes');
    const list = doc.createElement('ol');
    list.className = 'print-notes-list';
    for (const note of group.notes) {
      const item = doc.createElement('li');
      item.className = 'print-note';
      item.dataset['note'] = note.key;
      const label = doc.createElement('span');
      label.className = 'print-note-label';
      label.textContent = labels?.get(note.key) ?? '';
      const body = doc.createElement('div');
      body.className = 'print-note-body';
      item.append(label, body);
      list.appendChild(item);
      views.push(new StoryView(body, session, note, { label: label.textContent }));
    }
    section.append(heading, list);
  }
  mount.insertAdjacentElement('afterend', section);
  return () => {
    for (const view of views) view.destroy();
    section.remove();
  };
}
