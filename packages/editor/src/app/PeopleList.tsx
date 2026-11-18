// SPDX-License-Identifier: GPL-3.0-or-later
//
// The people in the room (ED-004 section 5): initials in the application bar, one per peer in
// their colour, the name and where they are as the tooltip; a click goes to their caret. Who
// joins and leaves is announced politely, with the count. An agent (ED-005 section 8.2) is badged
// and has a Stop, which any person in the room may press: written into the room, it makes the
// agent's companion leave and refuse the agent's further calls there (section 12.28).
import { Fragment, useEffect, useRef, useState } from 'react';
import { stopAgent, type EditStore } from '@docx4j/editor-model';
import type { Awareness } from 'y-protocols/awareness';
import { useT } from './i18n.js';
import { initials, peersOf, whereText, type Peer } from './presence.js';
import { authorPattern } from '../view/authors.js';

export function PeopleList({ awareness, store, onStatus, noteLabel }: {
  awareness: Awareness;
  store: EditStore;
  onStatus: (text: string) => void;
  /** The number a note's reference shows, by `kind:id` (the body view's decorations hold them). */
  noteLabel?: (key: string) => string | undefined;
}) {
  const t = useT();
  const [peers, setPeers] = useState<Peer[]>(() => peersOf(awareness, store));
  const [announcement, setAnnouncement] = useState('');
  const known = useRef(new Map<number, string>(peers.map((p) => [p.clientId, p.name])));

  useEffect(() => {
    const changed = (): void => {
      const now = peersOf(awareness, store);
      setPeers(now);
      const before = known.current;
      const after = new Map(now.map((p) => [p.clientId, p.name]));
      const joined = [...after].filter(([id]) => !before.has(id)).map(([, name]) => name);
      const left = [...before].filter(([id]) => !after.has(id)).map(([, name]) => name);
      known.current = after;
      const count = now.length + 1;
      if (joined.length > 0) setAnnouncement(t('presence.joined', { names: joined.join(', '), count }));
      else if (left.length > 0) setAnnouncement(t('presence.left', { names: left.join(', '), count }));
    };
    awareness.on('change', changed);
    changed();
    return () => { awareness.off('change', changed); };
  }, [awareness, store, t]);

  /** Scroll to the peer's caret and show its name for a moment; a caret in a story that is not on screen is said instead. */
  const goTo = (peer: Peer): void => {
    const caret = document.querySelector<HTMLElement>(`.peer-caret[data-peer="${peer.clientId}"]`);
    if (!caret) { onStatus(t('presence.notShown', { name: peer.name, where: whereText(peer.where, noteLabel) })); return; }
    caret.scrollIntoView({ block: 'center' });
    caret.classList.add('flash');
    setTimeout(() => caret.classList.remove('flash'), 1500);
  };

  return (
    <span className="people" data-testid="people" aria-label={t('presence.people', { count: peers.length + 1 })}>
      {peers.map((p) => (
        <Fragment key={p.clientId}>
          <button type="button" className={`person${p.agent ? ' agent' : ''}`} data-testid="person" data-peer={p.clientId} data-agent={p.agent ? 'true' : undefined}
            style={{ backgroundColor: p.color, backgroundImage: authorPattern(p.name, 'rgb(255 255 255 / 0.3)') }} title={t('presence.person', { name: p.name, where: whereText(p.where, noteLabel) })}
            aria-label={t('presence.person', { name: p.name, where: whereText(p.where, noteLabel) })} onClick={() => goTo(p)}>
            {initials(p.name)}
          </button>
          {p.agent && <span className="agent-badge" data-testid="agent-badge" title={t('presence.agent.title', { name: p.name })}>{t('presence.agent')}</span>}
          {p.agent && (
            <button type="button" className="link agent-stop" data-testid="agent-stop" title={t('presence.stop.title', { name: p.name })} aria-label={t('presence.stop.title', { name: p.name })}
              onClick={() => { stopAgent(store, p.name); onStatus(t('status.agent.stopped', { name: p.name })); }}>
              {t('presence.stop')}
            </button>
          )}
        </Fragment>
      ))}
      <span className="visually-hidden" role="status" aria-live="polite" data-testid="people-announcement">{announcement}</span>
    </span>
  );
}
