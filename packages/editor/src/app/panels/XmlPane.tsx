// SPDX-License-Identifier: GPL-3.0-or-later
//
// Reveal codes (ED-001 section 3.5 items 1 and 3; ED-002 section 6.2): breadcrumbs from the
// caret, the Open XML of the block or of the run under the caret marshalled from the tree, an
// editor that applies a change through the tree (decision 20), and the content-API script that
// would build the element. The XML is shown in Monaco read-only until Edit (ED-005 section 12.99):
// in a table, the caret's cell is marked and the other rows and cells are folded.
import { Suspense, lazy, useCallback, useEffect, useMemo, useState } from 'react';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { toApiScript, typeNameOf, segmentsOf, type Element } from '@docx4j/core-ts';
import type { EditorSurface } from '../../view/surface.js';
import { useEditorState } from '../state.js';
import { Fold } from './Fold.js';
import { t, useT } from '../i18n.js';
import { crumbsOf, blockAt, blocksIn, runUnderCaret, xmlOf, parseElement, parseElements, xmlSyntaxError, XmlSyntaxError, outlineOf, tableChainOf, cellFocusOf, type Crumb, type XmlFocus } from '../xml.js';

const XmlCodeEditor = lazy(() => import('./XmlCodeEditor.js'));

export interface XmlPaneProps {
  editor: EditorSurface;
  /** Hand a script to the console ("Open in console"). */
  onOpenInConsole?: (script: string) => void;
}

type Level = 'block' | 'run';

export function XmlPane({ editor, onOpenInConsole }: XmlPaneProps) {
  useT();
  const state = useEditorState(editor);
  const [level, setLevel] = useState<Level>('block');
  const [xml, setXml] = useState('');
  const [script, setScript] = useState('');
  const [editing, setEditing] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  /** The line the parser objects to while editing, live. */
  const [errorLine, setErrorLine] = useState<number | null>(null);
  const [crumbs, setCrumbs] = useState<Crumb[]>([]);
  const [hasRun, setHasRun] = useState(false);
  const [showScript, setShowScript] = useState(false);
  const [showSource, setShowSource] = useState(false);
  const [count, setCount] = useState(1);
  /** In a table, the caret's cell in the text and what to fold around it. */
  const [focus, setFocus] = useState<XmlFocus | null>(null);

  useEffect(() => {
    if (!state || editing !== null) return;
    const { crumbs: c, run } = crumbsOf(state);
    const blocks = blocksIn(state);
    setCrumbs(c);
    setCount(blocks.nodes.length);
    // The run level exists for a caret or a selection within one block.
    setHasRun(run !== null && blocks.nodes.length === 1);
    if (blocks.nodes.length !== 1 && level === 'run') setLevel('block');
    if (blocks.nodes.length === 0) { setXml(''); setScript(''); return; }
    let cancelled = false;
    const timer = setTimeout(() => {
      void (async () => {
        try {
          const elements = blocks.nodes.map((n) => editor.exportBlock(n));   // through the story's context in a note or a header
          const single = elements.length === 1 ? elements[0]! : undefined;
          const target: Element | Element[] = level === 'run' && run !== null && single ? runUnderCaret(editor.session, state, single)?.run ?? single : (single ?? elements);
          const xmls = Array.isArray(target) ? await Promise.all(target.map(xmlOf)) : [await xmlOf(target)];
          const s = showScript ? await toApiScript(target).catch((e: Error) => `// ${e.message}`) : '';
          const text = xmls.join('\n');
          const chain = level === 'block' && single ? tableChainOf(state) : [];
          const f = chain.length > 0 ? cellFocusOf(outlineOf(text), chain) ?? null : null;
          if (!cancelled) { setXml(text); setScript(s); setError(null); setFocus(f); }
        } catch (e) {
          if (!cancelled) setXml(t('xml.notAvailable', { error: (e as Error).message }));
        }
      })();
    }, 150);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [editor, state, level, editing, showScript]);

  const select = useCallback((crumb: Crumb) => {
    const { doc, tr } = editor.state;
    const sel = crumb.node.isTextblock ? TextSelection.create(doc, crumb.pos + 1, crumb.pos + crumb.node.nodeSize - 1)
      : crumb.node.type.spec.atom ? NodeSelection.create(doc, crumb.pos) : TextSelection.create(doc, crumb.pos + 1);
    editor.view.dispatch(tr.setSelection(sel).scrollIntoView());
    editor.focus();
  }, [editor]);

  const apply = useCallback(async () => {
    if (editing === null || !state) return;
    const block = blockAt(state);
    const blocks = blocksIn(state);
    if (!block || blocks.nodes.length === 0) return;
    setError(null);
    const isRun = level === 'run';
    let replacements: Element[];
    try {
      replacements = isRun ? [await parseElement(editing, 'run')] : await parseElements(editing, 'block');
    } catch (e) {
      setError(e instanceof XmlSyntaxError && e.line !== undefined ? t('xml.lineError', { line: e.line, message: e.message }) : (e as Error).message);
      setErrorLine(e instanceof XmlSyntaxError ? e.line ?? null : null);
      return;
    }
    const replacement = replacements[0]!;
    const blockIndex = block.index;
    const runInfo = isRun ? runUnderCaret(editor.session, state, editor.exportBlock(block.node)) : undefined;
    if (isRun && !runInfo) { setError(t('xml.noRun')); return; }
    const result = await editor.session.runOnTree(({ body }) => {
      const content = body.content;
      if (!isRun) {
        // The selected blocks are replaced as a group by whatever the text holds, one or more.
        content.splice(blocks.from, blocks.to - blocks.from + 1, ...(replacements as never[]));
        return;
      }
      const paragraph = content[blockIndex];
      if (!paragraph || typeNameOf(paragraph) !== 'org_docx4j_wml.P') throw new Error('editing a run inside a content control is not supported yet; edit the block');
      let runIndex = -1;
      let lastRun: object | undefined;
      for (const seg of segmentsOf(paragraph.value as object)) {
        if (seg.run !== lastRun) { runIndex++; lastRun = seg.run; }
        if (runIndex === runInfo!.index) { seg.runOwner[seg.runIndex] = replacement; return; }
      }
      throw new Error('the run is no longer where it was');
    });
    if (result.error) { setError((result.error as Error).message); return; }
    setEditing(null);
    setErrorLine(null);
    editor.focus();
  }, [editing, state, level, editor]);

  return (
    <Fold id="codes" className="panel xml-pane" label={t('tab.codes')} testId="panel-xml" title={t('tab.codes')} meta={count > 1 ? t('xml.blocksSelected', { count }) : undefined}>
      <nav className="crumbs" aria-label={t('xml.ancestors')}>
        <span className="crumb">body</span>
        {crumbs.map((c) => (
          <span key={c.depth}>{' > '}<button type="button" className={`crumb link${level === 'block' && c.depth === (crumbs[0]?.depth ?? 1) ? ' current' : ''}`} onClick={() => { setLevel('block'); select(c); }}>{c.label}</button></span>
        ))}
        {hasRun && <span>{' > '}<button type="button" className={`crumb link${level === 'run' ? ' current' : ''}`} data-testid="crumb-run" onClick={() => setLevel('run')}>r</button></span>}
      </nav>
      <div className="pane-actions">
        {editing === null
          ? <button type="button" data-testid="xml-edit" onClick={() => { setEditing(xml); setErrorLine(null); }} disabled={xml === '' || xml.startsWith('(')}>{t('xml.edit')}</button>
          : <>
              <button type="button" data-testid="xml-apply" onClick={() => void apply()}>{t('xml.apply')}</button>
              <button type="button" onClick={() => { setEditing(null); setError(null); setErrorLine(null); }}>{t('xml.cancel')}</button>
            </>}
        <label><input type="checkbox" checked={showScript} onChange={(e) => setShowScript(e.target.checked)} /> {t('xml.apiScript')}</label>
        <label><input type="checkbox" checked={showSource} onChange={(e) => setShowSource(e.target.checked)} /> {t('xml.factorySource')}</label>
      </div>
      {showSource && (
        <p className="meta" data-testid="factory-source-note">{t('xml.factorySource.note')}</p>
      )}
      <XmlEditor value={editing ?? xml} readOnly={editing === null} focus={editing === null ? focus : undefined} errorLine={errorLine} message={error} level={level}
        onChange={(v) => {
          if (editing === null) return;
          setEditing(v);
          const syntax = xmlSyntaxError(v, level);
          setErrorLine(syntax?.line ?? null);
          setError(syntax ? t('xml.lineError', { line: syntax.line ?? '?', message: syntax.message }) : null);
        }} />
      {error && <p className="error" role="alert" data-testid="xml-error">{error}</p>}
      {showScript && (
        <div className="api-script">
          <h3>{t('xml.script.title')}</h3>
          <pre data-testid="api-script">{script}</pre>
          <div className="pane-actions">
            <button type="button" onClick={() => void navigator.clipboard?.writeText(script)}>{t('xml.copy')}</button>
            {onOpenInConsole && <button type="button" data-testid="open-in-console" onClick={() => onOpenInConsole(script)}>{t('xml.openInConsole')}</button>}
          </div>
        </div>
      )}
    </Fold>
  );
}

/**
 * The XML shown and edited, in Monaco (ED-005 section 3.8): its find and replace inside the editor,
 * colouring, folding, and the line the parser objects to marked. Read-only until Edit, with the
 * caret's cell in focus in a table (section 12.99). Loaded on first use.
 */
export function XmlEditor({ value, readOnly = false, focus, errorLine, onChange, level, message }: { value: string; readOnly?: boolean; focus?: XmlFocus | null; errorLine: number | null; onChange: (v: string) => void; level: Level; message?: string | null }) {
  useT();
  const problems = useMemo(() => (errorLine !== null ? [{ line: errorLine, message: message ?? t('xml.lineError', { line: errorLine, message: '' }) }] : []), [errorLine, message]);
  return (
    <Suspense fallback={<p className="meta">{t('tab.loading')}</p>}>
      <XmlCodeEditor value={value} readOnly={readOnly} focus={focus} onChange={onChange} problems={problems} testId="xml-editor" ariaLabel={level === 'run' ? t('xml.editor.run') : t('xml.editor.block')} />
    </Suspense>
  );
}
