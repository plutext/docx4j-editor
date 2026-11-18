// SPDX-License-Identifier: GPL-3.0-or-later
//
// The macro dialogs (ED-003 section 7.4): a console script saved as a macro with a name and a
// home (this browser, or the document as a custom XML part), and the source of a document macro
// shown before its first run in this browser (ED-001 section 11 item 5, written into the UI).
import { useEffect, useRef, useState } from 'react';
import type { Macro } from '@docx4j/editor-model';
import type { MacroHome } from './macros.js';
import { useT } from './i18n.js';

export function SaveMacroDialog({ script, documentAllowed, onSave, onClose }: { script: string; documentAllowed: boolean; onSave: (name: string, home: MacroHome) => void; onClose: () => void }) {
  const t = useT();
  const [name, setName] = useState('');
  const [home, setHome] = useState<MacroHome>('browser');
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { input.current?.focus(); }, []);
  return (
    <div className="dialog-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <form className="dialog" role="dialog" aria-label={t('macroDialog.save.title')} data-testid="save-macro-dialog" onSubmit={(e) => { e.preventDefault(); if (name.trim() !== '') onSave(name.trim(), home); }}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <h2>{t('macroDialog.save.title')}</h2>
        <label className="field"><span>{t('macroDialog.name')}</span><input ref={input} type="text" data-testid="macro-name" value={name} onChange={(e) => setName(e.target.value)} required /></label>
        <fieldset className="macro-home">
          <legend>{t('macroDialog.keep')}</legend>
          <label><input type="radio" name="home" value="browser" checked={home === 'browser'} onChange={() => setHome('browser')} /> {t('macroDialog.browser')}</label>
          <label><input type="radio" name="home" value="document" data-testid="macro-home-document" checked={home === 'document'} disabled={!documentAllowed} onChange={() => setHome('document')} /> {t('macroDialog.document')}{documentAllowed ? '' : t('macroDialog.document.off')}</label>
        </fieldset>
        <pre className="macro-source">{script}</pre>
        <div className="dialog-buttons">
          <button type="button" onClick={onClose}>{t('dialog.cancel')}</button>
          <button type="submit" data-testid="macro-save">{t('macroDialog.save')}</button>
        </div>
      </form>
    </div>
  );
}

export function MacroSourceDialog({ macro, onRun, onClose }: { macro: Macro; onRun: () => void; onClose: () => void }) {
  const t = useT();
  const button = useRef<HTMLButtonElement>(null);
  useEffect(() => { button.current?.focus(); }, []);
  return (
    <div className="dialog-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="dialog" role="dialog" aria-label={t('macroDialog.source.label')} data-testid="macro-source-dialog" onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <h2>{t('macroDialog.source.title', { name: macro.name })}</h2>
        <p>{t('macroDialog.source.note')}</p>
        <pre className="macro-source" data-testid="macro-source">{macro.script}</pre>
        <div className="dialog-buttons">
          <button ref={button} type="button" onClick={onClose}>{t('dialog.cancel')}</button>
          <button type="button" data-testid="macro-run-anyway" onClick={onRun}>{t('macroDialog.run')}</button>
        </div>
      </div>
    </div>
  );
}
