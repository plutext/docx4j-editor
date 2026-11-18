// SPDX-License-Identifier: GPL-3.0-or-later
//
// The final section lives in the store's `package` map (ED-002 section 10.1 item 2); the state's
// `doc.attrs.finalSectPr` is a projection of it. This plugin keeps the two in step both ways: a
// change in the map (a peer, an undo) becomes a transaction setting the attribute, and a
// transaction that set the attribute (`section.*` on the final section) writes the map.
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { type EditorDocument, FINAL_SECT_PR, LOCAL_ORIGIN, plainEqual } from '@docx4j/editor-model';

export const finalSectPrKey = new PluginKey('docx4j-final-sectPr');

export function finalSectPrPlugin(session: EditorDocument): Plugin {
  const map = session.store.package;
  return new Plugin({
    key: finalSectPrKey,
    view(view: EditorView) {
      const fromMap = (): void => {
        const value = map.get(FINAL_SECT_PR) ?? null;
        if (plainEqual(value, view.state.doc.attrs['finalSectPr'])) return;
        view.dispatch(view.state.tr.setDocAttribute('finalSectPr', value).setMeta(finalSectPrKey, 'from-map'));
      };
      map.observe(fromMap);
      return {
        update(v: EditorView, prev: EditorState) {
          const value = v.state.doc.attrs['finalSectPr'] as unknown;
          if (plainEqual(value, prev.doc.attrs['finalSectPr'])) return;
          if (plainEqual(value, map.get(FINAL_SECT_PR) ?? null)) return;
          session.store.ydoc.transact(() => {
            if (value === null || value === undefined) map.delete(FINAL_SECT_PR);
            else map.set(FINAL_SECT_PR, value);
          }, LOCAL_ORIGIN);
        },
        destroy() {
          map.unobserve(fromMap);
        },
      };
    },
  });
}
