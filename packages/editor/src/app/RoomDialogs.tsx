// SPDX-License-Identifier: GPL-3.0-or-later
//
// Share and Join (ED-004 section 4.4). Share shows the room's link with a Copy button and says
// what sharing means: anyone with the link can edit, and the relay's host can read what is
// shared. Join takes a pasted link.
import { useState } from 'react';
import { Dialog } from './Dialogs.js';
import { useT } from './i18n.js';
import { parseRoomLink, relayHost, type RoomLink } from './room.js';

export function ShareDialog({ url, relay, onClose }: { url: string; relay: string; onClose: () => void }) {
  const t = useT();
  const [copied, setCopied] = useState(false);
  const copy = (): void => {
    void navigator.clipboard?.writeText(url).then(() => setCopied(true), () => setCopied(false));
  };
  return (
    <Dialog title={t('share.title')} onClose={onClose} onOk={onClose}>
      <p>{t('share.anyone')}</p>
      <p data-testid="share-relay">{t('share.relay', { host: relayHost(relay) })}</p>
      <label className="field share-link">
        <span>{t('share.link')}</span>
        <input type="text" readOnly value={url} data-testid="share-link" onFocus={(e) => e.currentTarget.select()} />
      </label>
      <button type="button" data-testid="share-copy" onClick={copy}>{copied ? t('share.copied') : t('share.copy')}</button>
    </Dialog>
  );
}

export function JoinDialog({ onJoin, onClose }: { onJoin: (link: RoomLink) => void; onClose: () => void }) {
  const t = useT();
  const [text, setText] = useState('');
  const [error, setError] = useState(false);
  const ok = (): void => {
    const link = parseRoomLink(text.trim());
    if (!link) { setError(true); return; }
    onJoin(link);
  };
  return (
    <Dialog title={t('join.title')} onClose={onClose} onOk={ok}>
      <label className="field share-link">
        <span>{t('join.link')}</span>
        <input type="text" value={text} data-testid="join-link" onChange={(e) => { setText(e.target.value); setError(false); }} />
      </label>
      {error && <p className="meta" role="alert" data-testid="join-error">{t('join.notALink')}</p>}
    </Dialog>
  );
}
