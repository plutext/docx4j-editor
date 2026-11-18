// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Preview tab (ED-003 section 4.7): the bind step in place on a clone of the package
// (`previewInstance`), shown as a read-only story with the document's look, the controls drawn
// with their tags and their verdicts as on the page; recomputed, debounced, while the tab is in
// front, on every change to the document, the controls map, a part or the data part. Repeats
// are not expanded and dropped controls stay, dimmed: the badges say what a processor would do.
// An XHTML-bound control shows its markup converted, as the processor converts it (ED-005 section
// 12.117): the styles come from the clone's resolver through a stylesheet scoped to the preview
// (a style the conversion added, or one the document defines and the body does not use, is
// covered by no rule of the body's), and the list labels and indents from the clone's numbering.
import { useEffect, useMemo, useRef, useState } from 'react';
import { previewInstance, type Preview } from '@docx4j/editor-model';
import { StoryView } from '../../../view/StoryView.js';
import { documentCss } from '../../../view/css.js';
import { verdictsPlugin, VerdictCache } from '../../../view/plugins/verdicts.js';
import { useEditorState, useStoreRevision } from '../../state.js';
import { useTemplate } from '../../template.js';
import type { TabContext } from '../../perspectives.js';
import { t, useT } from '../../i18n.js';

const DEBOUNCE_MS = 400;

export default function PreviewPanel({ editor, session, perspective }: TabContext) {
  useT();
  const state = useEditorState(editor);
  const revision = useStoreRevision(session);
  const tpl = useTemplate(session);
  const [preview, setPreview] = useState<Preview | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [templateRevision, setTemplateRevision] = useState(0);
  const host = useRef<HTMLDivElement>(null);
  const view = useRef<StoryView | null>(null);
  const inFront = perspective.panel.tab === 'preview';
  /** What the shown preview was computed from; a change while the tab is hidden is caught up when it comes to the front. */
  const shown = useRef<{ doc: unknown; revision: number; templateRevision: number } | null>(null);

  useEffect(() => session.template.subscribe(() => setTemplateRevision((r) => r + 1)), [session]);

  useEffect(() => {
    if (!tpl.loaded || !state || !inFront) return;
    const key = { doc: state.doc, revision, templateRevision };
    if (shown.current && shown.current.doc === key.doc && shown.current.revision === key.revision && shown.current.templateRevision === key.templateRevision) return;
    let cancelled = false;
    const timer = setTimeout(() => {
      setBusy(true);
      void previewInstance(session).then((p) => {
        if (cancelled) return;
        shown.current = key;
        setPreview(p);
        setError(null);
      }, (e: unknown) => { if (!cancelled) setError((e as Error).message); }).finally(() => { if (!cancelled) setBusy(false); });
    }, DEBOUNCE_MS);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [session, state, revision, templateRevision, tpl.loaded, inFront]);

  // The story view lives as long as the host; each preview replaces its state (a new clone, new ids: a fresh view keeps the node views honest).
  useEffect(() => {
    if (!host.current || !preview) return;
    const v = new StoryView(host.current, session, preview, {
      label: t('preview.story'), part: preview.main, ids: preview.ctx.ids, controls: preview.controls, labels: preview.labels,
      ...(preview.numbering ? { numbering: preview.numbering } : {}),
      plugins: [verdictsPlugin(session, preview.controls, new VerdictCache(session))],
    });
    view.current = v;
    return () => { v.destroy(); view.current = null; };
  }, [session, preview]);

  const css = useMemo(() => (preview
    ? documentCss(session, preview.doc, [], { resolver: preview.pkg.propertyResolverOrUndefined, selector: preview.main.runFontSelectorOrUndefined, scope: '.preview-page' })
    : ''), [session, preview]);

  return (
    <section className="panel preview-panel" aria-label={t('tab.preview')} data-testid="panel-preview">
      <h2>{t('tab.preview')} <span className="meta">{t('preview.meta')}</span></h2>
      {tpl.error && <p className="error" role="alert">{tpl.error}</p>}
      {error && <p className="error" role="alert" data-testid="preview-error">{error}</p>}
      {preview && (
        <p className="meta" data-testid="preview-summary">
          {t('preview.summary', { bound: preview.result.bound, updated: preview.result.updated, skipped: preview.result.skipped })}
          {' '}{t('preview.repeats')}{busy ? t('preview.recomputing') : ''}
        </p>
      )}
      {!preview && !error && <p className="meta">{tpl.loaded ? t('preview.computing') : t('data.loading')}</p>}
      <style data-testid="preview-css">{css}</style>
      <div className="show-control-tags preview-page"><div ref={host} data-testid="preview-story" /></div>
    </section>
  );
}
