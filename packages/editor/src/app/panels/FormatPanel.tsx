// SPDX-License-Identifier: GPL-3.0-or-later
//
// Reveal formatting (Jason, 2026-09-24, after Word's): the effective run, paragraph and
// section properties at the caret, each element of the resolved `w:rPr`, `w:pPr` and `w:sectPr`
// with its attributes, from the resolver the page's stylesheet comes from.
import { useEffect, useState } from 'react';
import type { EditorDocument } from '@docx4j/editor-model';
import type { EditorSurface } from '../../view/surface.js';
import { useEditorState } from '../state.js';
import { effectiveAt, propertyRows, type PropertyRow } from '../format.js';
import { t, useT } from '../i18n.js';

interface Sections {
  font: PropertyRow[];
  paragraph: PropertyRow[];
  section: PropertyRow[];
}

export function FormatPanel({ editor, session }: { editor: EditorSurface; session: EditorDocument }) {
  useT();
  const state = useEditorState(editor);
  const [rows, setRows] = useState<Sections>({ font: [], paragraph: [], section: [] });
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!state) return;
    let cancelled = false;
    const timer = setTimeout(() => {
      void (async () => {
        try {
          const effective = effectiveAt(session, state);
          const [font, paragraph, section] = await Promise.all([propertyRows('rPr', effective.rPr), propertyRows('pPr', effective.pPr), propertyRows('sectPr', effective.sectPr)]);
          if (!cancelled) { setRows({ font, paragraph, section }); setError(null); }
        } catch (e) {
          if (!cancelled) setError((e as Error).message);
        }
      })();
    }, 100);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [session, state]);

  return (
    <section className="panel format-panel" aria-label={t('tab.format')} data-testid="panel-format">
      <h2>{t('format.heading')} <span className="meta">{t('format.meta')}</span></h2>
      {error && <p className="error" role="alert">{error}</p>}
      <Group title={t('format.font')} kind="rPr" rows={rows.font} />
      <Group title={t('format.paragraph')} kind="pPr" rows={rows.paragraph} />
      <Group title={t('format.section')} kind="sectPr" rows={rows.section} />
    </section>
  );
}

function Group({ title, kind, rows }: { title: string; kind: string; rows: PropertyRow[] }) {
  return (
    <div className="format-group" data-testid={`format-${kind}`}>
      <h3>{title} <span className="meta">w:{kind}</span></h3>
      {rows.length === 0 ? <p className="meta">{t('format.nothing')}</p> : <Rows rows={rows} />}
    </div>
  );
}

function Rows({ rows }: { rows: PropertyRow[] }) {
  return (
    <ul className="property-rows">
      {rows.map((r, i) => (
        <li key={`${r.name}-${i}`}>
          <span className="element">{r.name}</span>
          {r.attributes.map((a) => <span key={a.name} className="attribute"> <span className="attr-name">{a.name}</span>=<span className="attr-value">"{a.value}"</span></span>)}
          {r.children.length > 0 && <Rows rows={r.children} />}
        </li>
      ))}
    </ul>
  );
}
