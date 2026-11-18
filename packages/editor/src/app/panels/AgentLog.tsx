// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the agents in the room did (ED-005 section 8.6, step F4): the room's `agentLog`, newest first,
// each call with its time, its agent, its tool and its summary (the change report's first line, a
// reading's gist, or the refusal); a click goes to where it acted, by the paragraph id where it has
// one (which the person's edits do not shift), else by the address as it was. In the review panel,
// beside the changes the calls made. The store is subscribed through `useSyncExternalStore`
// (section 12.14).
import { useCallback, useRef, useSyncExternalStore } from 'react';
import { TextSelection } from 'prosemirror-state';
import { agentLog, onAgents, type AgentLogEntry, type EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { t as tr, formatDate } from '../i18n.js';

export function AgentLog({ editor, session, onStatus, lang }: { editor: DocxEditorView; session: EditorDocument; onStatus: (s: string) => void; lang: string }) {
  void lang;
  const store = session.store;
  const version = useRef(0);
  const subscribe = useCallback((fn: () => void) => onAgents(store, () => { version.current++; fn(); }), [store]);
  useSyncExternalStore(subscribe, () => version.current);
  const entries = agentLog(store);
  if (entries.length === 0) return null;

  const goTo = async (entry: AgentLogEntry): Promise<void> => {
    const where = entry.paraIds[0] !== undefined ? { paraId: entry.paraIds[0] } : entry.addresses[0];
    if (where === undefined) return;
    const pos = await session.positionOf(where);
    const state = editor.state;
    if (pos === undefined || pos > state.doc.content.size) { onStatus(tr('status.agentLog.gone')); return; }
    editor.view.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(Math.min(pos + 1, state.doc.content.size)))).scrollIntoView());
    editor.focus();
  };

  return (
    <section className="agent-log" data-testid="agent-log" aria-label={tr('review.agentLog')}>
      <h3>{tr('review.agentLog')}</h3>
      <ol>
        {[...entries].reverse().map((entry, i) => (
          <li key={`${entry.time}:${entries.length - i}`} data-testid="agent-log-entry" data-ok={entry.ok ? 'true' : 'false'} data-tool={entry.tool}>
            <button type="button" className="agent-log-main" disabled={entry.addresses.length === 0 && entry.paraIds.length === 0} onClick={() => void goTo(entry)}>
              <span className="meta">{formatDate(entry.time)}</span>
              <span className="who">{entry.agent}</span>
              {!entry.ok && <span className="badge refused">{tr('review.agentLog.refused')}</span>}
              <span className="summary" dir="auto">{entry.summary}</span>
            </button>
          </li>
        ))}
      </ol>
    </section>
  );
}
