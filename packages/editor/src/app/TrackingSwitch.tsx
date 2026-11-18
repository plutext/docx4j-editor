// SPDX-License-Identifier: GPL-3.0-or-later
//
// The tracking switch (ED-005 section 6.6): in the application bar in every perspective, since the
// mode is the document's and matters to every edit wherever it is made. Pressed when tracking is on,
// its title naming the author. It reads the mode on each transaction of the body view, and again
// when a peer's settings part arrives (the view's `partsChanged`), when the status line says the
// mode was changed in the shared document. Turning it on is the application's (`onToggle`), which
// asks for a name first when "Your name" is empty.
import { useEffect, useRef } from 'react';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import { useEditorState } from './state.js';
import { useT } from './i18n.js';

export interface TrackingSwitchProps {
  editor: DocxEditorView;
  author: string;
  onToggle: () => void;
  onStatus: (text: string) => void;
  /** Set by the application while it changes the mode itself, so the change is not reported as a peer's. */
  local: { current: boolean };
}

export function TrackingSwitch({ editor, author, onToggle, onStatus, local }: TrackingSwitchProps) {
  const t = useT();
  useEditorState(editor);
  const on = editor.session.trackingMode === 'TrackAll';
  const previous = useRef(on);
  useEffect(() => {
    if (previous.current === on) return;
    previous.current = on;
    if (local.current) { local.current = false; return; }
    onStatus(t('status.tracking.peer', { state: on ? 'on' : 'off' }));
  }, [on, local, onStatus, t]);
  return (
    <button type="button" className={`tracking-switch${on ? ' on' : ''}`} aria-pressed={on} data-testid="tracking-switch"
      title={on ? t('tracking.switch.on.title', { author }) : t('tracking.switch.off.title')} onClick={onToggle}>
      {t('tracking.switch')}
    </button>
  );
}

/**
 * The status bar's tracking indicator, as Word's status bar shows "Track Changes: On": always in view
 * where the status line's own message passes. A click toggles the mode as the switch does. Its text
 * says the state, so it takes no `aria-pressed` (a toggle's name does not change with its state).
 */
export function TrackingIndicator({ editor, author, onToggle }: Pick<TrackingSwitchProps, 'editor' | 'author' | 'onToggle'>) {
  const t = useT();
  useEditorState(editor);
  const on = editor.session.trackingMode === 'TrackAll';
  return (
    <button type="button" className={`tracking-indicator${on ? ' on' : ''}`} data-testid="tracking-indicator"
      title={on ? t('tracking.switch.on.title', { author }) : t('tracking.switch.off.title')} onClick={onToggle}>
      {t('tracking.indicator', { state: on ? 'on' : 'off' })}
    </button>
  );
}

