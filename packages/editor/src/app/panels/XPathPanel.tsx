// SPDX-License-Identifier: GPL-3.0-or-later
//
// Select by XPath (ED-005 section 3.5): a tab of the Developer perspective. An expression over the
// main document part's XML as the edit layer has it (or a notes, header or footer part's), run
// on Enter and never per keystroke, since a run exports the edit layer and marshals the part; the
// hits listed with their address, the element's name and a snippet; a click selects where the hit
// projects (in the body, or in the note's or header's view). "Open in console" writes the script
// that selects the same, a replace by XPath being a console script (section 10). The engine is
// the browser's own XPath 1.0 (`document.evaluate`, through core-ts's default engine) unless
// "XPath 3.1" is ticked, which loads FontoXPath, the Template perspective's engine: over the
// 255-page document `//w:p[w:pPr/w:pStyle]` takes 0.6 s with the first and 25 s with the second
// in Chromium (ED-005 section 12.1), so 3.1 is the option, not the default.
import { useCallback, useEffect, useState } from 'react';
import { NodeSelection, TextSelection, type Selection } from 'prosemirror-state';
import { DefaultXPathEngine } from '@docx4j/core-ts';
import { selectScript, type EditorDocument, type XPathHit } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { surfaces, type EditorSurface } from '../../view/surface.js';
import { loadTemplate } from '../template.js';
import { useT } from '../i18n.js';

/** The browser's own XPath 1.0, whatever engine the package has installed. */
const native = new DefaultXPathEngine();

export interface XPathPanelProps {
  editor: DocxEditorView;
  session: EditorDocument;
  onConsoleScript: (script: string) => void;
  onActivate?: (surface: EditorSurface) => void;
}

/** Select what a hit places at in a surface: a paragraph's text, a run's text, an atom, the start of a table or a row. */
export function selectAt(surface: EditorSurface, pos: number): void {
  const { doc } = surface.state;
  const at = Math.min(pos, doc.content.size);
  const node = doc.nodeAt(at);
  let sel: Selection;
  if (node?.isTextblock) sel = TextSelection.create(doc, at + 1, at + node.nodeSize - 1);
  else if (node?.isText) sel = TextSelection.create(doc, at, at + node.nodeSize);
  else if (node && NodeSelection.isSelectable(node)) sel = NodeSelection.create(doc, at);
  else sel = TextSelection.near(doc.resolve(node && !node.isLeaf ? at + 1 : at));
  surface.reveal(sel.from, sel.to);
  surface.view.dispatch(surface.state.tr.setSelection(sel).scrollIntoView());
  surface.focus();
}

export default function XPathPanel({ editor, session, onConsoleScript, onActivate }: XPathPanelProps) {
  const t = useT();
  const [xpath31, setXpath31] = useState(false);
  const [expression, setExpression] = useState('//w:p[w:pPr/w:pStyle]');
  const [part, setPart] = useState('');
  const [parts, setParts] = useState<string[]>([]);
  const [hits, setHits] = useState<XPathHit[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [running, setRunning] = useState(false);
  const [ms, setMs] = useState<number | null>(null);

  useEffect(() => {
    let cancelled = false;
    void session.stories.selectableParts().then((p) => { if (!cancelled) setParts(p); });
    return () => { cancelled = true; };
  }, [session]);

  const run = useCallback(async () => {
    if (expression.trim() === '' || running) return;
    setRunning(true);
    setError(null);
    const start = performance.now();
    try {
      if (xpath31) await loadTemplate(session);
      const engine = xpath31 ? session.pkg.xpathEngine : native;
      setHits(await session.selectXPath(expression, { engine, ...(part === '' ? {} : { part }) }));
      setMs(Math.round(performance.now() - start));
    } catch (e) {
      setHits(null);
      setError((e as Error).message);
    } finally {
      setRunning(false);
    }
  }, [expression, part, running, session, xpath31]);

  const go = (hit: XPathHit): void => {
    if (!hit.story) { selectAt(editor, hit.pos); return; }
    const target = [...surfaces.values()].find((s) => s.session === session && s.story?.key === hit.story!.key);
    if (!target) return;
    selectAt(target, hit.pos);
    onActivate?.(target);
  };

  const addressOf = (hit: XPathHit): string => `${hit.story ? hit.story.key : 'body'}/${hit.block}`;

  return (
    <section className="panel xpath-panel" aria-label={t('tab.xpath')} data-testid="panel-xpath">
      <h2>{t('tab.xpath')}</h2>
      <div className="xpath-form">
        <input type="text" dir="ltr" spellCheck={false} data-testid="xpath-expression" aria-label={t('xpath.expression')} value={expression}
          onChange={(e) => setExpression(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); void run(); } }} />
        <select data-testid="xpath-part" aria-label={t('xpath.part')} value={part} onChange={(e) => setPart(e.target.value)}>
          <option value="">{t('xpath.part.main')}</option>
          {parts.map((p) => <option key={p} value={p}>{p}</option>)}
        </select>
        <label title={t('xpath.version.title')}><input type="checkbox" data-testid="xpath-31" checked={xpath31} onChange={(e) => setXpath31(e.target.checked)} /> {t('xpath.version')}</label>
        <button type="button" data-testid="xpath-run" disabled={running} onClick={() => void run()}>{running ? t('xpath.running') : t('xpath.run')}</button>
        <button type="button" data-testid="xpath-console" onClick={() => onConsoleScript(selectScript(expression))}>{t('xml.openInConsole')}</button>
      </div>
      {error !== null && <p className="error" role="alert" data-testid="xpath-error">{error}</p>}
      {hits !== null && <p className="meta" data-testid="xpath-count" aria-live="polite">{t('xpath.count', { count: hits.length, ms: ms ?? 0 })}</p>}
      {hits !== null && hits.length > 0 && (
        <div className="table-box">
          <table data-testid="xpath-hits">
            <thead><tr><th>{t('xpath.address')}</th><th>{t('xpath.name')}</th><th>{t('xpath.snippet')}</th></tr></thead>
            <tbody>
              {hits.slice(0, 500).map((h, i) => (
                <tr key={i} title={h.path} className={h.exact ? '' : 'approximate'}>
                  <td><button type="button" className="link" data-testid="xpath-hit" onClick={() => go(h)}>{addressOf(h)}</button></td>
                  <td><code>{h.name}</code></td>
                  <td className="text" dir="auto">{h.snippet}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {hits.length > 500 && <p className="meta">{t('xpath.more', { count: hits.length - 500 })}</p>}
        </div>
      )}
    </section>
  );
}
