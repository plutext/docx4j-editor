// SPDX-License-Identifier: GPL-3.0-or-later
//
// The projection diagnostics (ED-001 section 11 item 8; section 3.6's Developer perspective):
// what is opaque and why, by kind and address, grouped (ED-003 section 8.6); a click selects the
// node. The banner's list is the same component.
import { embeddedFontStatusOf } from '../fonts.js';
import { preservedCount, type EditorDocument } from '@docx4j/editor-model';
import { listedChanges } from '../review.js';
import type { EditorSurface } from '../../view/surface.js';
import { useNotes } from './FootnotesPanel.js';
import { DiagnosticList } from './DiagnosticList.js';
import { NormalisationSection } from './NormalisationSection.js';
import { useT } from '../i18n.js';

export function DiagnosticsPanel({ editor, session }: { editor: EditorSurface; session: EditorDocument }) {
  const t = useT();
  const d = session.diagnostics;
  const notes = useNotes(session);
  const fonts = embeddedFontStatusOf(session);
  const orphans = notes ? [...notes.footnotes, ...notes.endnotes].filter((n) => n.orphan).length : 0;
  // What arrived from a shared room and could not be applied (ED-004 section 4.4): a kind of its own.
  const problems = session.inRoom ? session.roomProblems : [];
  return (
    <section className="panel" aria-label={t('tab.diagnostics')} data-testid="panel-diagnostics">
      <h2>{t('tab.diagnostics')}</h2>
      <dl>
        <dt>{t('diagnostics.opaque')}</dt><dd>{d.opaque.length}</dd>
        <dt>{t('diagnostics.tracked')}</dt><dd data-testid="diag-tracked">{listedChanges(session)}</dd>
        <dt>{t('diagnostics.preserved')}</dt><dd data-testid="diag-preserved">{preservedCount(d)}{preservedCount(d) > 0 ? ` (${Object.entries(d.tracked.preserved).map(([k, n]) => `${k} ${n}`).join(', ')})` : ''}</dd>
        <dt>{t('diagnostics.markers')}</dt><dd>{d.pageBreakMarkers}</dd>
        <dt>{t('diagnostics.comments')}</dt><dd>{d.comments}</dd>
        {fonts && (fonts.loaded.length > 0 || fonts.failed.length > 0) && <><dt>{t('diagnostics.fonts')}</dt><dd data-testid="diag-fonts">{t('diagnostics.fonts.loaded', { count: fonts.loaded.length })}{fonts.repaired.length > 0 ? t('diagnostics.fonts.repaired', { count: fonts.repaired.length, repairs: [...new Set(fonts.faces.flatMap((f) => f.repairs))].join('; ') }) : ''}{fonts.failed.length > 0 ? t('diagnostics.fonts.failed', { count: fonts.failed.length, faces: fonts.failed.join('; ') }) : ''}</dd></>}
        <dt>{t('diagnostics.footnotes')}</dt><dd data-testid="diag-footnotes">{notes ? notes.footnotes.length : '…'}</dd>
        <dt>{t('diagnostics.endnotes')}</dt><dd>{notes ? notes.endnotes.length : '…'}</dd>
        {orphans > 0 && <><dt>{t('diagnostics.orphans')}</dt><dd>{orphans}</dd></>}
      </dl>
      <h3>{t('diagnostics.byKind')}</h3>
      <DiagnosticList editor={editor} entries={d.opaque} />
      {problems.length > 0 && (
        <>
          <h3>{t('diagnostics.roomProblems', { count: problems.length })}</h3>
          <ul data-testid="diag-room-problems">
            {problems.map((p) => <li key={`${p.kind}:${p.key}`}>{t(`diagnostics.roomProblems.${p.kind}`)} <code>{p.key}</code>: {p.reason}</li>)}
          </ul>
        </>
      )}
      <NormalisationSection session={session} />
    </section>
  );
}
