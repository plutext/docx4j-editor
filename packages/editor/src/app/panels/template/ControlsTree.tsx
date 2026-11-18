// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Controls tab (ED-003 section 4.5): every control in the document in document order with
// its nesting, each with its form, kind, title or tag, its role and its verdict; a click selects
// the control's content (`contentControl.select`); the caret's control is highlighted. Headers,
// footers and notes list their controls under their story, read-only (E2.d edits them).
import { useEffect, useMemo, useState } from 'react';
import { listControls, controlAt, isNodeTarget, type ControlSummary, type EditorDocument } from '@docx4j/editor-model';
import { useEditorState, useStoreRevision } from '../../state.js';
import { useTemplate, formLabel, kindLabel, controlLabel, roleBadge, verdictSummary } from '../../template.js';
import { t, useT, tNodes } from '../../i18n.js';
import type { TabContext } from '../../perspectives.js';

interface StoryControls {
  label: string;
  controls: ControlSummary[];
}

async function storyControls(session: EditorDocument): Promise<StoryControls[]> {
  const out: StoryControls[] = [];
  const none = (): undefined => undefined;
  for (const h of await session.stories.headers()) {
    const controls = listControls(h.doc, none);
    if (controls.length > 0) out.push({ label: t(h.kind === 'header' ? 'controls.story.header' : 'controls.story.footer', { id: h.rId }), controls });
  }
  for (const n of [...(await session.stories.footnotes()), ...(await session.stories.endnotes())]) {
    const controls = listControls(n.doc, none);
    if (controls.length > 0) out.push({ label: t(n.kind === 'footnote' ? 'controls.story.footnote' : 'controls.story.endnote', { id: String(n.id) }), controls });
  }
  return out;
}

export default function ControlsTree({ editor, session, onTab }: TabContext) {
  const lang = useT();
  const state = useEditorState(editor);
  const revision = useStoreRevision(session);
  const tpl = useTemplate(session);
  const controls = useMemo(() => (state ? listControls(state.doc, session.controlEntry) : []), [state, session, revision]);   // eslint-disable-line react-hooks/exhaustive-deps
  const current = state ? controlAt(state) : undefined;
  const [stories, setStories] = useState<StoryControls[]>([]);
  useEffect(() => {
    let cancelled = false;
    void storyControls(session).then((s) => { if (!cancelled) setStories(s); });
    return () => { cancelled = true; };
  }, [session, lang]);

  const isCurrent = (c: ControlSummary): boolean => current !== undefined && (isNodeTarget(current) ? c.pos === current.pos : c.groupId === current.groupId);
  const select = (c: ControlSummary): void => {
    editor.run('contentControl.select', c.pos !== undefined ? { pos: c.pos } : { groupId: c.groupId });
    editor.focus();
  };
  const verdict = (c: ControlSummary): string => {
    if (!tpl.loaded) return '';
    try { return verdictSummary(session.template.verdictFor(c.sdtPr)); } catch (e) { return t('verdict.error', { error: (e as Error).message }); }
  };

  const row = (c: ControlSummary, i: number, live: boolean) => {
    const badge = roleBadge(c.sdtPr);
    const processorSet = c.tag.processorSet.length > 0;
    return (
      <li key={`${c.form}-${c.pos ?? c.groupId}-${i}`} className={`control-row${isCurrent(c) && live ? ' current' : ''}`} style={{ paddingInlineStart: `${c.depth * 1.1}rem` }} aria-current={isCurrent(c) && live ? 'true' : undefined} aria-level={c.depth + 1}>
        <span className={`form form-${c.form}`}>{formLabel(c.form)}</span>
        {live
          ? <button type="button" className={`link${isCurrent(c) ? ' current' : ''}`} onClick={() => select(c)}>{controlLabel(c.sdtPr)}</button>
          : <span className="label">{controlLabel(c.sdtPr)}</span>}
        <span className="meta"> {kindLabel(c.reading.kind)}</span>
        {badge && <span className={`badge badge-${c.tag.role !== 'none' ? c.tag.role : 'bound'}`}>{badge}</span>}
        {processorSet && <span className="meta processor-set"> {c.tag.processorSet.map((p) => p.key).join(', ')}</span>}
        {live && <span className="meta verdict"> {verdict(c)}</span>}
      </li>
    );
  };

  return (
    <section className="panel controls-tree" aria-label={t('tab.controls')} data-testid="panel-controls">
      <h2>{t('tab.controls')} <span className="meta">{t('controls.meta', { count: controls.length })}</span></h2>
      {controls.length === 0
        ? <p className="meta">{t('controls.none')}</p>
        : <ul className="control-list" aria-label={t('controls.body')} data-testid="controls-list">{controls.map((c, i) => row(c, i, true))}</ul>}
      {stories.map((s) => (
        <div key={s.label} className="story-controls">
          <h3>{s.label} <span className="meta">{t('controls.story.meta')}</span></h3>
          <ul className="control-list" aria-label={t('controls.inStory', { story: s.label })}>{s.controls.map((c, i) => row(c, i, false))}</ul>
        </div>
      ))}
      <p className="meta">{tNodes('controls.atCaret', { tab: <button type="button" className="link" onClick={() => onTab('control')}>{t('tab.control')}</button> })}</p>
    </section>
  );
}
