// SPDX-License-Identifier: GPL-3.0-or-later
//
// One part's XML in Monaco with Apply, Revert and Format (ED-005 sections 4.3 and 12.62): the parts panel's editor,
// and since section 12.64 the Data tab's XML view of a custom XML part, whose context menu adds Bind and Copy XPath.
import { Suspense, lazy, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { Part } from '@docx4j/core-ts';
import { checkPartXml, applyPartXml, type EditorDocument, type PartEditing, type PartXmlProblem, type AppliedPart } from '@docx4j/editor-model';
import { t, useT } from '../i18n.js';
import { pretty, formatXml } from '../xml.js';
import type { XmlEditorApi, XmlContextAction } from './XmlCodeEditor.js';

const XmlCodeEditor = lazy(() => import('./XmlCodeEditor.js'));

/**
 * The text a part is edited as: a typed part pretty-printed (whitespace between its elements means
 * nothing), an XML part as it is (its text nodes are data). Line ends are a line feed, as Monaco
 * keeps them and as an XML parser reads any other (XML 1.0 section 2.11), so the text shown is the
 * text compared.
 */
export async function editableText(part: Part, editing: PartEditing): Promise<string> {
  const xml = (await (part as unknown as { getXml(): Promise<string> }).getXml()).replace(/\r\n?/g, '\n');
  return editing.mode === 'typed' ? pretty(xml) : xml;
}

/** A problem said as the pane says it, with its line when the check found one. */
function problemText(p: PartXmlProblem): string {
  const where = p.line !== undefined ? t('parts.problem.line', { line: p.line }) : '';
  switch (p.kind) {
    case 'element': return `${t('parts.problem.element', { name: p.name ?? '', path: p.path ?? '' })}${where}`;
    case 'attribute': return `${t('parts.problem.attribute', { name: p.name ?? '', path: p.path ?? '' })}${where}`;
    case 'value': return `${t('parts.problem.value', { message: p.message ?? '' })}${where}`;
    default: return `${t('parts.problem.parse', { message: p.message ?? '' })}${where}`;
  }
}

interface PartEditorProps {
  session: EditorDocument;
  part: Part;
  editing: PartEditing;
  initial: string;
  onApplied: (result: AppliedPart) => void;
  /** Entries for the editor's context menu, given the offset and the text (the Data tab's Bind and Copy XPath). */
  contextActions?: readonly XmlContextAction[];
  /** The editor's test id; `part-xml` in the parts panel. */
  testId?: string;
}

/**
 * One part's XML in Monaco with Apply and Revert (ED-005 section 4.3). Apply checks first and
 * applies only a text with no problem; Revert puts back the text the pane opened with, which is
 * another Apply. A change to the part from elsewhere (a peer in the room, the Template
 * perspective) reloads the text when nothing is being edited, and otherwise is said, with the
 * choice of reloading (the text here lost) or applying over it (section 4.5).
 */
export function PartEditor({ session, part, editing, initial, onApplied, contextActions, testId = 'part-xml' }: PartEditorProps) {
  useT();
  const [text, setText] = useState(initial);
  const [problems, setProblems] = useState<PartXmlProblem[]>([]);
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [remote, setRemote] = useState<string | null>(null);
  /** The part's text as last read or applied here: what "edited" is measured from. */
  const seen = useRef(initial);
  /** Revert's text: the part as the pane opened it. */
  const opened = useRef(initial);
  /** The editor's text as it is now, set as it changes (a click can come before React renders the change). */
  const current = useRef(text);
  /** The XML editor's own actions, once it is loaded: Format. */
  const xmlApi = useRef<XmlEditorApi | null>(null);
  const [formattable, setFormattable] = useState(false);
  const onApi = useCallback((api: XmlEditorApi | null) => { xmlApi.current = api; setFormattable(api !== null); }, []);
  /**
   * Format (Jason, 2026-10-04): the text indented in the editor as one edit, which Ctrl+Z undoes; nothing
   * is written until Apply. A part Word wrote on one line (`stylesWithEffects.xml`) can then be read.
   */
  const format = useCallback(async () => {
    try {
      formatXml(current.current);
    } catch (e) {
      setStatus(t('parts.format.failed', { error: (e as Error).message }));
      return;
    }
    await xmlApi.current?.format();
    setStatus(t('parts.format.done'));
  }, []);
  const edited = useCallback((v: string) => { current.current = v; setText(v); }, []);

  // A peer's edit arriving, or a part edit made elsewhere in this session (the Template perspective's
  // Data tab, a console script), which bumps the package map's parts revision: read the part again.
  useEffect(() => {
    const reread = (): void => {
      void editableText(part, editing).then((now) => {
        if (now === seen.current) return;
        if (current.current === seen.current) { seen.current = now; edited(now); setRemote(null); } else setRemote(now);
      });
    };
    const off = session.onPartsReceived(reread);
    session.store.package.observe(reread);
    return () => { off(); session.store.package.unobserve(reread); };
  }, [session, part, editing, edited]);

  const apply = useCallback(async (xml: string) => {
    setBusy(true);
    setStatus(null);
    try {
      const found = await checkPartXml(part, xml);
      setProblems(found);
      if (found.length > 0) { setStatus(t('parts.refused', { count: found.length })); return; }
      const result = await applyPartXml(session, part, xml);
      seen.current = xml;
      setRemote(null);
      onApplied(result);
      setStatus(result.dangling.length === 0 ? t('parts.applied')
        : t('parts.applied.dangling', { names: result.dangling.map((d) => t(`parts.dangling.${d.kind}`, { id: d.id, count: d.count })).join(', ') }));
    } catch (e) {
      setStatus(t('parts.applyFailed', { error: (e as Error).message }));
    } finally {
      setBusy(false);
    }
  }, [part, session, onApplied]);

  const markers = useMemo(() => problems.filter((p) => p.line !== undefined).map((p) => ({ line: p.line!, ...(p.column !== undefined ? { column: p.column } : {}), message: problemText(p) })), [problems]);

  return (
    <div className="part-editor" data-testid="part-editor">
      <div className="pane-actions">
        <button type="button" data-testid="part-apply" disabled={busy || text === seen.current} onClick={() => void apply(current.current)} title={t('parts.apply.title')}>{t('parts.apply')}</button>
        <button type="button" data-testid="part-revert" disabled={busy || (text === opened.current && seen.current === opened.current)}
          onClick={() => { edited(opened.current); setProblems([]); if (seen.current !== opened.current) void apply(opened.current); }} title={t('parts.revert.title')}>{t('parts.revert')}</button>
        <button type="button" data-testid="part-format" disabled={busy || !formattable} onClick={() => void format()} title={t('parts.format.title')}>{t('parts.format')}</button>
        <span className="meta">{editing.mode === 'typed' ? t('parts.mode.typed') : t('parts.mode.bytes')}</span>
      </div>
      {remote !== null && (
        <p className="banner" role="status" data-testid="part-remote">
          {t('parts.remote')}{' '}
          <button type="button" data-testid="part-reload" onClick={() => { seen.current = remote; edited(remote); setRemote(null); setProblems([]); }}>{t('parts.remote.reload')}</button>{' '}
          <button type="button" data-testid="part-apply-mine" onClick={() => void apply(current.current)}>{t('parts.remote.applyMine')}</button>
        </p>
      )}
      <Suspense fallback={<p className="meta">{t('tab.loading')}</p>}>
        <XmlCodeEditor value={text} onChange={edited} problems={markers} testId={testId} ariaLabel={t('parts.editor', { name: part.partName.toString() })} api={onApi} {...(contextActions ? { actions: contextActions } : {})} />
      </Suspense>
      {problems.length > 0 && (
        <ul className="error part-problems" role="alert" data-testid="part-problems">
          {problems.map((p, i) => <li key={i}>{problemText(p)}</li>)}
        </ul>
      )}
      {status !== null && <p className="meta" role="status" data-testid="part-status">{status}</p>}
    </div>
  );
}
