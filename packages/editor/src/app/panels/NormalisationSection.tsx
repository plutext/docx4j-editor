// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Diagnostics tab's normalisation report (ED-008 section 8.8): made on the button, never at open, since it reads
// the whole package. It changes nothing; Apply waits on a file that needs it (Jason, 2026-11-05).
import { Fragment, useState } from 'react';
import type { EditorDocument } from '@docx4j/editor-model';
import { normalisationReport, type NormalisationReport } from '../normalisation.js';
import { useT } from '../i18n.js';

type Run = { session: EditorDocument } & ({ status: 'running' } | { status: 'done'; report: NormalisationReport } | { status: 'failed'; reason: string });

export function NormalisationSection({ session }: { session: EditorDocument }) {
  const t = useT();
  const [run, setRun] = useState<Run | undefined>();
  // A report belongs to the document it read: another one opened in this panel starts without it.
  const current = run?.session === session ? run : undefined;
  const start = (): void => {
    setRun({ session, status: 'running' });
    normalisationReport(session).then(
      (report) => setRun((r) => (r?.session === session ? { session, status: 'done', report } : r)),
      (e: unknown) => setRun((r) => (r?.session === session ? { session, status: 'failed', reason: e instanceof Error ? e.message : String(e) } : r)),
    );
  };
  return (
    <>
      <h3>{t('diagnostics.normalisation')}</h3>
      <p className="meta">{t('diagnostics.normalisation.about')}</p>
      <button type="button" data-testid="normalisation-run" disabled={current?.status === 'running'} onClick={start}>{t('diagnostics.normalisation.run')}</button>
      {current?.status === 'running' && <p className="meta" role="status">{t('diagnostics.normalisation.running')}</p>}
      {current?.status === 'failed' && <p className="meta" role="alert" data-testid="normalisation-failed">{t('diagnostics.normalisation.failed', { reason: current.reason })}</p>}
      {current?.status === 'done' && <ReportList report={current.report} />}
    </>
  );
}

function ReportList({ report }: { report: NormalisationReport }) {
  const t = useT();
  const notes = report.notes.reduce((n, g) => n + g.messages.length, 0);
  return (
    <div className="diagnostic-list" data-testid="normalisation-report">
      <p className="meta" data-testid="normalisation-summary">{report.values === 0 ? t('diagnostics.normalisation.clean') : t('diagnostics.normalisation.values', { count: report.values })}</p>
      {report.parts.map((part) => (
        <Fragment key={part.partName}>
          <p className="meta" data-part={part.partName}><code>{part.partName}</code> <span className="count">{part.count}</span></p>
          {part.kinds.map((k) => (
            <details key={k.kind} className="diagnostic-kind" data-kind={k.kind}>
              <summary><code>{k.kind}</code> <span className="count">{k.count}</span></summary>
              <ul>
                {k.changes.map((c) => (
                  <li key={`${c.before}\n${c.after ?? ''}`}>
                    {c.after === null ? t('diagnostics.normalisation.removal', { before: c.before }) : t('diagnostics.normalisation.change', { before: c.before, after: c.after })} <span className="count">{c.count}</span>
                  </li>
                ))}
              </ul>
            </details>
          ))}
        </Fragment>
      ))}
      {notes > 0 && <p className="meta" data-testid="normalisation-notes">{t('diagnostics.normalisation.notes', { count: notes })}</p>}
      {report.notes.map((g) => (
        <details key={g.code} className="diagnostic-kind" data-kind={g.code}>
          <summary><code>{g.code}</code> <span className="count">{g.messages.length}</span></summary>
          <ul>{g.messages.map((m, i) => <li key={i}>{m}</li>)}</ul>
        </details>
      ))}
    </div>
  );
}
