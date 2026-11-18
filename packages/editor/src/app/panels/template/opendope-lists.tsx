// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Questions and Components parts as read-only lists (ED-003 sections 4.5 and 4.6): the Data
// tab shows them under their headings, the OpenDoPE tab too. A question lists the entries whose
// `@questionID` names it; a component its IRI, with the note that fetching is a processor step.
import type { QuestionsPart, ComponentsPart, XPathsPart, EntryUse } from '@docx4j/editor-model';
import { t } from '../../i18n.js';

/** How an XPaths entry is used, as badges in the Controls tree's colours (one per kind), then who uses it. */
export function EntryUses({ uses }: { uses: EntryUse[] }) {
  if (uses.length === 0) return <span className="meta">{t('od.unused')}</span>;
  const kinds = [...new Set(uses.map((u) => u.kind))];
  return (
    <>
      {kinds.map((k) => <span key={k} className={`badge badge-${k}`}>{t(`od.use.${k}`)}</span>)}
      <span className="meta"> {uses.map((u) => u.label).join(', ')}</span>
    </>
  );
}

export function QuestionsList({ questions, xpaths, meta }: { questions: QuestionsPart; xpaths: XPathsPart | undefined; meta: string }) {
  return (
    <div className="od-part" data-testid="od-questions">
      <h4>{t('od.questions')} <span className="meta">{meta}{questions.title ? ` · ${questions.title}` : ''} · {t('od.questions.readOnly')}</span></h4>
      {questions.questions.length === 0 ? <p className="meta">{t('od.none')}</p> : (
        <ul className="od-list">
          {questions.questions.map((q) => {
            const elicits = xpaths?.entries.filter((e) => e.attrs['questionID'] === q.id).map((e) => e.id) ?? [];
            return <li key={q.id}><b>{q.id}</b> {q.text}<span className="meta"> · {q.response.kind === 'free' ? t('od.question.free') : t('od.question.choices', { count: q.response.items.length })}{elicits.length > 0 ? ` · ${t('od.question.elicits', { ids: elicits.join(', ') })}` : ''}</span></li>;
          })}
        </ul>
      )}
    </div>
  );
}

export function ComponentsList({ components, meta, users }: { components: ComponentsPart; meta: string; users?: (id: string) => string[] }) {
  return (
    <div className="od-part" data-testid="od-components">
      <h4>{t('od.components')} <span className="meta">{t('od.components.meta', { meta })}</span></h4>
      {components.components.length === 0 ? <p className="meta">{t('od.none')}</p> : (
        <ul className="od-list">
          {components.components.map((c) => {
            const u = users?.(c.id) ?? [];
            return <li key={c.id}><b>{c.id}</b> <code>{c.iri}</code>{users ? <span className="meta"> · {u.length === 0 ? t('od.unused') : t('data.usedBy', { users: u.join(', ') })}</span> : null}</li>;
          })}
        </ul>
      )}
    </div>
  );
}
