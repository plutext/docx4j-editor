// SPDX-License-Identifier: GPL-3.0-or-later
//
// A failed round-trip self-check (ED-001 section 11 item 1): the save is blocked and never
// overwrites; the differing block is shown as XML on both sides (live from the export, reloaded
// from the diagnosis bytes), and the bytes can be downloaded for diagnosis under a suffixed name.
import { useEffect, useState } from 'react';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, type SelfCheckError } from '@docx4j/editor-model';
import { xmlOf } from './xml.js';
import { t, useT } from './i18n.js';

export interface SelfCheckDialogProps {
  error: SelfCheckError;
  session: EditorDocument;
  diagnosis: Uint8Array;
  name: string;
  onDiagnosis: () => void;
  onClose: () => void;
}

export function SelfCheckDialog({ error, session, diagnosis, name, onDiagnosis, onClose }: SelfCheckDialogProps) {
  useT();
  const d = error.result.difference;
  const [sides, setSides] = useState<{ live: string; reloaded: string } | null>(null);
  useEffect(() => {
    if (!d || d.index < 0) return;
    let cancelled = false;
    void (async () => {
      try {
        const live = await xmlOf(session.exportBlock(session.doc.child(d.index)));
        const pkg = await WordprocessingMLPackage.load(diagnosis);
        const body = (await pkg.getMainDocumentPart().readContents()).body;
        const el = body?.content?.[d.index];
        const reloaded = el ? await xmlOf(el) : t('selfCheck.noBlock');
        if (!cancelled) setSides({ live, reloaded });
      } catch (e) {
        if (!cancelled) setSides({ live: d.live, reloaded: `${d.reloaded}\n(${(e as Error).message})` });
      }
    })();
    return () => { cancelled = true; };
  }, [d, session, diagnosis]);
  return (
    <div className="dialog-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="dialog self-check" role="alertdialog" aria-modal="true" aria-label={t('selfCheck.label')} data-testid="self-check-dialog">
        <h2>{t('selfCheck.title')}</h2>
        <p>{t('selfCheck.note')} {d && d.index >= 0 ? t('selfCheck.block', { index: d.index }) : t('selfCheck.finalSection')}</p>
        {d && (
          <div className="self-check-diff">
            <div><h3>{t('selfCheck.live')}</h3><pre>{sides?.live ?? d.live}</pre></div>
            <div><h3>{t('selfCheck.reloaded')}</h3><pre>{sides?.reloaded ?? d.reloaded}</pre></div>
          </div>
        )}
        <div className="dialog-buttons">
          <button type="button" data-testid="self-check-download" onClick={onDiagnosis}>{t('selfCheck.download', { name })}</button>
          <button type="button" onClick={onClose} autoFocus>{t('dialog.close')}</button>
        </div>
      </div>
    </div>
  );
}
