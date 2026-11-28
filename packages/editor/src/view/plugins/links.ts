// SPDX-License-Identifier: GPL-3.0-or-later
//
// Following a link (ED-003 section 5.5, step C4): Ctrl+click (Cmd on a Mac) on a link opens an
// external address in a new tab, and puts the caret at an anchor's bookmark. A plain click
// places the caret, as in Word.
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { linkAt, bookmarkPosition, type EditorDocument } from '@docx4j/editor-model';
import { scrollBlockToTop } from '../scroll.js';

export const linksKey = new PluginKey('docx4j-links');

/** How an external address is opened: a new tab by default; the application sets its host's (ED-014 section 4.1), so the view knows nothing of it. */
let opener = (url: string): void => { window.open(url, '_blank', 'noopener'); };

/** Set how Ctrl+click opens an external address. */
export function setLinkOpener(open: (url: string) => void): void {
  opener = open;
}

export function linksPlugin(_session: EditorDocument): Plugin {
  return new Plugin({
    key: linksKey,
    props: {
      handleClick(view, pos, event) {
        if (!(event.ctrlKey || event.metaKey)) return false;
        const state = view.state.apply(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
        const link = linkAt(state);
        if (!link) return false;
        if (link.address) {
          opener(link.address);
          return true;
        }
        if (link.anchor) {
          const at = bookmarkPosition(state, link.anchor);
          if (at === undefined) return false;
          // The target's block goes to the top of the view, its content under it, not the least scroll that shows
          // the caret (ED-005 section 12.94): the focus first, since focusing may scroll on its own.
          view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
          view.focus();
          scrollBlockToTop(view, at);
          return true;
        }
        return false;
      },
    },
  });
}
