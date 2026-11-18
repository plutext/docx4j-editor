// SPDX-License-Identifier: GPL-3.0-or-later
//
// The formatting stack at the caret (Jason, 2026-09-24): for the run, the list, the paragraph
// and the table, the layers Word resolves in order, nearest first: the direct formatting, the
// style chain the resolver climbs (a run with no `w:rStyle` carries the default character
// style; a paragraph with no `w:pStyle` the default paragraph style), then the document
// defaults. Each entry opens its XML below the lists, editable: a style or the document
// defaults are replaced in the styles part in place, direct formatting through the tree.
import { useCallback, useEffect, useState } from 'react';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { EditorDocument } from '@docx4j/editor-model';
import type { EditorSurface } from '../../view/surface.js';
import { useEditorState } from '../state.js';
import { styleChains, styleXml, parseStyle, replaceStyle, layerXml, parseLayer, replaceDirect, replaceDocDefaults, type StyleChains, type LayerKind } from '../format.js';
import { xmlSyntaxError, XmlSyntaxError } from '../xml.js';
import { t, useT } from '../i18n.js';
import { XmlEditor } from './XmlPane.js';

const EMPTY: StyleChains = {
  character: [], numbering: [], paragraph: [], table: [],
  applied: { character: undefined, paragraph: undefined, numId: undefined },
  direct: { rPr: undefined, pPr: undefined }, defaults: { rPr: undefined, pPr: undefined, present: false },
};

/** What the XML box shows: a style by id, or a layer that is not a style. */
type Selection = { kind: 'style'; id: string } | { kind: 'layer'; layer: LayerKind };

/** A layer's title (`styles.layer.<kind>`), and the entry for an absent one. */
const layerTitle = (layer: LayerKind): string => t(`styles.layer.${layer}`);

export function StylesPanel({ editor, session }: { editor: EditorSurface; session: EditorDocument }) {
  useT();
  const state = useEditorState(editor);
  const [chains, setChains] = useState<StyleChains>(EMPTY);
  const [selected, setSelected] = useState<Selection | null>(null);
  const [xml, setXml] = useState('');
  const [editing, setEditing] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [errorLine, setErrorLine] = useState<number | null>(null);
  /** Bumped after an applied edit so the chains and the XML are read again. */
  const [revision, setRevision] = useState(0);

  useEffect(() => {
    if (!state) return;
    let cancelled = false;
    const timer = setTimeout(() => {
      void styleChains(session, state).then((c) => { if (!cancelled) setChains(c); });
    }, 100);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [session, state, revision]);

  const all = [...chains.character, ...chains.numbering, ...chains.paragraph, ...chains.table];
  const currentStyle = selected?.kind === 'style' ? all.find((s) => s.styleId === selected.id) ?? session.pkg.propertyResolverOrUndefined?.getStyle(selected.id) : undefined;
  const layerValue = (layer: LayerKind): wml.RPr | wml.PPr | undefined =>
    layer === 'rPr' ? chains.direct.rPr : layer === 'pPr' ? chains.direct.pPr : layer === 'rPrDefault' ? chains.defaults.rPr : chains.defaults.pPr;

  useEffect(() => {
    if (!selected || editing !== null) return;
    let cancelled = false;
    const load = selected.kind === 'style' ? (currentStyle ? styleXml(currentStyle) : Promise.resolve('')) : layerXml(selected.layer, layerValue(selected.layer));
    void load.then((x) => { if (!cancelled) setXml(x); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected, currentStyle, chains, editing, revision]);

  const apply = useCallback(async () => {
    if (editing === null || !selected) return;
    setError(null);
    try {
      if (selected.kind === 'style') {
        const style = await parseStyle(editing);
        if (style.styleId !== selected.id) throw new Error(t('styles.idMustStay', { id: selected.id }));
        await replaceStyle(session, editor, style);
      } else {
        const value = await parseLayer(selected.layer, editing);
        if (selected.layer === 'rPr' || selected.layer === 'pPr') await replaceDirect(session, editor, selected.layer, value);
        else await replaceDocDefaults(session, editor, selected.layer, value);
      }
    } catch (e) {
      setError(e instanceof XmlSyntaxError && e.line !== undefined ? t('xml.lineError', { line: e.line, message: e.message }) : (e as Error).message);
      setErrorLine(e instanceof XmlSyntaxError ? e.line ?? null : null);
      return;
    }
    setEditing(null);
    setErrorLine(null);
    setRevision((r) => r + 1);
  }, [editing, selected, session, editor]);

  const pick = (sel: Selection): void => { setSelected(sel); setEditing(null); setError(null); setErrorLine(null); };
  const isSelected = (sel: Selection): boolean =>
    selected !== null && selected.kind === sel.kind && (sel.kind === 'style' ? (selected as { id: string }).id === sel.id : (selected as { layer: LayerKind }).layer === sel.layer);

  const layerEntry = (layer: LayerKind, present: boolean, testId: string) => (
    <li className="layer">
      <button type="button" className={`link${isSelected({ kind: 'layer', layer }) ? ' current' : ''}`} data-testid={testId} onClick={() => pick({ kind: 'layer', layer })}>
        {present ? layerTitle(layer) : t(`styles.layer.absent.${layer}`)}
      </button>
      <span className="meta"> {layer === 'rPr' || layer === 'pPr' ? `w:${layer}` : `w:docDefaults/w:${layer}`}</span>
    </li>
  );

  const list = (title: string, kind: keyof Omit<StyleChains, 'applied' | 'direct' | 'defaults'>, styles: wml.Style[], note: string | undefined, direct: LayerKind | undefined, defaults: LayerKind | undefined, empty: string) => (
    <div className="style-chain" data-testid={`styles-${kind}`}>
      <h3>{title}{note ? <span className="meta"> {note}</span> : null}</h3>
      {styles.length === 0 && !direct && !defaults
        ? <p className="meta">{empty}</p>
        : (
          <ol className="styles">
            {direct && layerEntry(direct, layerValue(direct) !== undefined, `direct-${direct}`)}
            {styles.length === 0 && <li className="meta">{empty}</li>}
            {styles.map((s, i) => (
              <li key={s.styleId ?? i}>
                <button type="button" className={`link${isSelected({ kind: 'style', id: s.styleId ?? '' }) ? ' current' : ''}`} onClick={() => pick({ kind: 'style', id: s.styleId ?? '' })}>{s.name?.val ?? s.styleId}</button>
                <span className="meta"> {s.styleId}{i === 0 ? '' : t('styles.basedOn')}{s._default ? t('styles.default') : ''}</span>
              </li>
            ))}
            {defaults && layerEntry(defaults, layerValue(defaults) !== undefined, `defaults-${defaults}`)}
          </ol>
        )}
    </div>
  );

  const title = selected?.kind === 'style' ? `${currentStyle?.name?.val ?? currentStyle?.styleId ?? selected.id}` : selected ? layerTitle(selected.layer) : '';
  const subtitle = selected?.kind === 'style' ? `w:style ${selected.id}` : selected ? (selected.layer === 'rPr' || selected.layer === 'pPr' ? t('styles.throughTree') : t('styles.inStylesPart')) : '';

  return (
    <section className="panel styles-panel" aria-label={t('tab.styles')} data-testid="panel-styles">
      <h2>{t('styles.heading')} <span className="meta">{t('styles.meta')}</span></h2>
      {list(t('styles.character'), 'character', chains.character, chains.applied.character ? undefined : chains.character.length > 0 ? t('styles.character.default') : undefined, 'rPr', 'rPrDefault', t('styles.character.none'))}
      {list(t('styles.numbering'), 'numbering', chains.numbering, chains.applied.numId ? t('styles.numbering.list', { id: chains.applied.numId }) : undefined, undefined, undefined, t('styles.numbering.none'))}
      {list(t('styles.paragraph'), 'paragraph', chains.paragraph, chains.applied.paragraph ? undefined : chains.paragraph.length > 0 ? t('styles.paragraph.default') : undefined, 'pPr', 'pPrDefault', t('styles.paragraph.none'))}
      {list(t('styles.table'), 'table', chains.table, chains.table.length > 0 && chains.table[0]?._default ? t('styles.table.default') : undefined, undefined, undefined, t('styles.table.none'))}
      {selected && (
        <div className="style-xml" data-testid="style-xml">
          <h3>{title} <span className="meta">{subtitle}</span></h3>
          <div className="pane-actions">
            {editing === null
              ? <button type="button" data-testid="style-edit" onClick={() => { setEditing(xml); setErrorLine(null); }} disabled={xml === ''}>{t('styles.edit')}</button>
              : <>
                  <button type="button" data-testid="style-apply" onClick={() => void apply()}>{t('styles.apply')}</button>
                  <button type="button" onClick={() => { setEditing(null); setError(null); setErrorLine(null); }}>{t('styles.cancel')}</button>
                </>}
          </div>
          {editing === null
            ? <pre data-testid="style-xml-text">{xml}</pre>
            : <XmlEditor value={editing} errorLine={errorLine} message={error} level="block"
                onChange={(v) => { setEditing(v); const syntax = xmlSyntaxError(v, 'block'); setErrorLine(syntax?.line ?? null); setError(syntax ? t('xml.lineError', { line: syntax.line ?? '?', message: syntax.message }) : null); }} />}
          {error && <p className="error" role="alert" data-testid="style-error">{error}</p>}
        </div>
      )}
    </section>
  );
}
