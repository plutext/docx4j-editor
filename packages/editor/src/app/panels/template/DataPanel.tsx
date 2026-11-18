// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Data tab (ED-003 section 4.5): the custom XML parts as trees, the data part first and
// named as such (or "undetermined" with the rule that failed), the OpenDoPE parts under their
// own heading with their typed reads, Microsoft's well-known parts collapsed. An XPath box
// evaluates an expression against the selected part and highlights the nodes it selects; a
// data node drags onto a control in the document to bind it (plugins/controls.ts reads the
// drag); the tree is edited in place (a child or a sibling added, an element duplicated, renamed
// or removed, attributes added, edited and removed, a leaf's text), through the DOM behind the
// engine's node wrappers with the part told by `touch()`: the sample data the preview evaluates
// against, written back to the part (no undo for a part edit, section 9 item 7). The OpenDoPE
// parts are read-only here; their own tab edits them.
import { useCallback, useEffect, useMemo, useState, type DragEvent, type ReactNode } from 'react';
import type { CustomXmlPart, CustomXmlNode } from '@docx4j/core-ts';
import { xhtmlKindOf, isOpenDopePart, controlToBind, controlAt, sdtPrOf, sdtId, applyBoundRefresh, dataBindRefusal, bindsPicture, bindsFragment, isFlatOpc, pictureFactsOf, partEditing, COVER_PAGE_NS, DOCUMENT_MANAGEMENT_NS, CORE_PROPERTIES_NS, BIBLIOGRAPHY_NS, type EditorDocument, type Operand, type EntryUse } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../../view/DocxEditorView.js';
import { BIND_MIME, type BindDrag } from '../../../view/plugins/controls.js';
import { refreshXhtmlControls } from '../../../view/xhtmlbind.js';
import { refreshBoundAtCaret } from '../../../view/boundrefresh.js';
import { refreshPictureControls } from '../../../view/picturebind.js';
import { refreshFragmentControls } from '../../../view/fragmentbind.js';
import { Fold } from '../Fold.js';
import { PartEditor, editableText } from '../PartXmlEditor.js';
import type { XmlContextAction } from '../XmlCodeEditor.js';
import { elementPathAt } from '../../xml.js';
import { useEditorState, useStoreRevision } from '../../state.js';
import { useTemplate, partLabel, operandText, imageDataUrl, useFragmentFacts, fragmentSummary } from '../../template.js';
import type { TabContext } from '../../perspectives.js';
import { QuestionsList, ComponentsList, EntryUses } from './opendope-lists.js';
import { t, useT, isRightToLeft } from '../../i18n.js';

const WELL_KNOWN = new Set([COVER_PAGE_NS, DOCUMENT_MANAGEMENT_NS, CORE_PROPERTIES_NS, BIBLIOGRAPHY_NS]);

/**
 * An edit of a part's DOM: the part is marked for the save and for the application's dirty state, the
 * template view forgets its reads, and every control bound to a node whose value changed shows it, as
 * Word's mapping does (ED-005 section 12.55), out of the undo history like the part edit itself.
 */
function partEdited(session: EditorDocument, part: CustomXmlPart, editor: DocxEditorView, onStatus: (s: string) => void): void {
  session.markPartChanged(part.part.partName.toString());
  session.template.invalidate();
  applyBoundRefresh(session, editor.state);
  // What the refreshes could not do (a floating picture, a node that is no image, markup dropped) reaches the
  // status line, each note once (ED-005 section 12.88).
  void Promise.all([refreshXhtmlControls(editor.view, session), refreshPictureControls(editor.view, session), refreshFragmentControls(editor, session)])
    .then((results) => { for (const note of new Set(results.flatMap((r) => r.notes))) onStatus(note); });
}

/**
 * The Data tab's bind of a node (the tree's bind, the XML view's context menu; ED-005 sections 12.56 to 12.58): the
 * control at the caret bound to it, as a drop on it would; with no control there, or inside one that holds controls, a
 * plain text control made there and bound in one step, as Word's XML Mapping Pane inserts one; a selection that cannot
 * hold one is refused, saying why, as is a node mapped to a rich text control.
 */
function bindNode(editor: DocxEditorView, session: EditorDocument, part: CustomXmlPart, node: CustomXmlNode, onStatus: (s: string) => void): void {
  const data = dragData(part, node);
  const xpath = node.xpath;
  const refusal = dataBindRefusal(editor.state, data);
  // A node holding image data makes a picture control (ED-005 section 12.117): where it cannot go, the refusal says so in its own words.
  const picture = bindsPicture(editor.state, data);
  const fragment = bindsFragment(editor.state, data);
  if (refusal) { onStatus(t(picture && (refusal === 'cells' || refusal === 'paragraphs' || refusal === 'content') ? 'data.bind.refusedPicture' : `data.bind.refused.${refusal}`)); return; }
  const xhtml = xhtmlKindOf(node.text) !== 'none';
  if (controlToBind(editor.state)) {
    const bound = editor.run('contentControl.bind', data);
    onStatus(bound ? t('status.drop.bound', { xpath }) : t('status.drop.notBound', { xpath }));
    // The control bound shows its node at once: markup converted, a picture control's image placed, a Word document filled
    // (ED-005 sections 12.88 and 12.124), that control alone.
    if (bound) void refreshBoundAtCaret(editor, session).then((notes) => { for (const note of notes) onStatus(note); });
    return;
  }
  const made = editor.run('contentControl.insertBound', data);
  onStatus(made ? t(xhtml ? 'data.bind.insertedXhtml' : picture ? 'data.bind.insertedPicture' : fragment ? 'data.bind.insertedFragment' : 'data.bind.inserted', { xpath }) : t('status.drop.notBound', { xpath }));
  // An XHTML node's markup converted into the new control (ED-005 section 12.79); an image node's picture placed in it (section 12.117).
  if (made && xhtml) void refreshXhtmlControls(editor.view, session).then((r) => { for (const note of r.notes) onStatus(note); });
  if (made && fragment) {
    // The new control filled with its Word document at once, as Word fills it on open (ED-005 proposal 46, part 3).
    const target = controlAt(editor.state);
    const id = target ? sdtId(sdtPrOf(target, session.controlEntry)) : undefined;
    void refreshFragmentControls(editor, session, id !== undefined ? { ids: new Set([id]) } : {});
  }
  if (made && picture) {
    // The new control alone, where the command left the caret: another control bound to the node says nothing here.
    const target = controlAt(editor.state);
    const id = target ? sdtId(sdtPrOf(target, session.controlEntry)) : undefined;
    void refreshPictureControls(editor.view, session, undefined, id !== undefined ? new Set([id]) : undefined).then((r) => { for (const note of r.notes) onStatus(note); });
  }
}

/** A node's XPath to the clipboard (ED-005 section 12.64), the status line saying so, with the prefixes it needs when it has any. */
async function copyXPath(node: CustomXmlNode, onStatus: (s: string) => void): Promise<void> {
  try {
    await navigator.clipboard.writeText(node.xpath);
    onStatus(node.prefixMappings ? t('data.copy.donePrefixes', { xpath: node.xpath, prefixes: node.prefixMappings }) : t('data.copy.done', { xpath: node.xpath }));
  } catch (e) {
    onStatus(t('data.copy.failed', { error: (e as Error).message }));
  }
}

/** The Data tab's view of a part, remembered in this browser. */
type PartView = 'tree' | 'xml';
const VIEW_KEY = 'docx4j-editor.data.view';
function storedView(): PartView {
  try { return localStorage.getItem(VIEW_KEY) === 'xml' ? 'xml' : 'tree'; } catch { return 'tree'; }
}

export default function DataPanel({ editor, session, onStatus }: TabContext) {
  useT();
  const state = useEditorState(editor);
  const revision = useStoreRevision(session);
  const tpl = useTemplate(session);
  const template = session.template;
  const [selectedId, setSelectedId] = useState<string>('');
  const [expression, setExpression] = useState('');
  const [prefixes, setPrefixes] = useState('');
  const [result, setResult] = useState<{ nodes: { xpath: string; text: string }[]; error?: string } | null>(null);
  const [hits, setHits] = useState<Set<string>>(() => new Set());

  const parts = tpl.loaded ? template.parts : [];
  const data = tpl.loaded ? template.dataPart : undefined;
  const od = tpl.loaded ? template.openDopeParts : undefined;
  const odParts = od ? [od.xpaths, od.conditions, od.components, od.questions, od.answers].filter((p): p is CustomXmlPart => p !== undefined) : [];
  const others = parts.filter((p) => p !== data?.part && !odParts.includes(p));
  const selected = parts.find((p) => p.id === selectedId) ?? data?.part ?? parts[0];
  useEffect(() => { if (data?.part && selectedId === '') setSelectedId(data.part.id); }, [data?.part, selectedId]);

  const evaluate = (): void => {
    if (!selected) return;
    try {
      const nodes = selected.selectNodes(expression, prefixes || undefined);
      setResult({ nodes: nodes.map((n) => ({ xpath: n.xpath, text: n.text })) });
      setHits(new Set(nodes.map((n) => n.xpath)));
    } catch (e) {
      setResult({ nodes: [], error: (e as Error).message });
      setHits(new Set());
    }
  };

  const users = useMemo(() => {
    if (!tpl.loaded || !state) return { entry: (): EntryUse[] => [], condition: (): string[] => [] };
    void revision;
    return { entry: (id: string) => template.usesOfEntry(id, state.doc), condition: (id: string) => template.usersOfCondition(id, state.doc) };
  }, [tpl.loaded, state, template, revision]);

  if (!tpl.loaded) {
    return <section className="panel data-panel" aria-label={t('tab.data')} data-testid="panel-data"><h2>{t('tab.data')}</h2>{tpl.error ? <p className="error" role="alert">{tpl.error}</p> : <p className="meta">{t('data.loading')}</p>}</section>;
  }

  return (
    <section className="panel data-panel" aria-label={t('tab.data')} data-testid="panel-data">
      <h2>{t('tab.data')} <span className="meta">{t('data.meta', { count: parts.length })}</span></h2>
      <div className="xpath-box" data-testid="xpath-box">
        <label>{t('data.xpath')} <input type="text" aria-label={t('data.xpath')} data-testid="xpath-expression" value={expression} onChange={(e) => setExpression(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); evaluate(); } }} placeholder="/invoice/items/item" /></label>
        <label>{t('data.prefixes')} <input type="text" aria-label={t('data.prefixes.label')} data-testid="xpath-prefixes" value={prefixes} onChange={(e) => setPrefixes(e.target.value)} placeholder={t('data.prefixes.placeholder')} /></label>
        <label>{t('data.part')}
          <select aria-label={t('data.part')} data-testid="xpath-part" value={selected?.id ?? ''} onChange={(e) => setSelectedId(e.target.value)}>
            {parts.map((p) => <option key={p.id} value={p.id}>{data?.part === p ? t('data.dataPart', { part: partLabel(p) }) : partLabel(p)}</option>)}
          </select>
        </label>
        <button type="button" data-testid="xpath-evaluate" onClick={evaluate}>{t('data.evaluate')}</button>
        {result && (
          <div className="xpath-result" data-testid="xpath-result">
            {result.error
              ? <span className="error">{result.error}</span>
              : <>
                  <span>{t('data.result.nodes', { count: result.nodes.length })}</span>
                  <ul>{result.nodes.slice(0, 20).map((n) => <li key={n.xpath}><code>{n.xpath}</code> <span className="meta">{n.text.length > 60 ? `${n.text.slice(0, 60)}…` : n.text}</span></li>)}</ul>
                  {result.nodes.length > 20 && <span className="meta">{t('data.result.more', { count: result.nodes.length - 20 })}</span>}
                </>}
          </div>
        )}
      </div>
      <Fold id="data.part" level={3} title={t('data.dataPart.title')} meta={data?.part ? partLabel(data.part) : t('data.dataPart.undetermined')}>
        {data?.part
          ? <PartTree session={session} editor={editor} part={data.part} label={t('data.dataPart.rule', { part: partLabel(data.part), rule: data.rule, reason: data.reason })} hits={selected === data.part ? hits : undefined} open onStatus={onStatus} testId="data-part" editable />
          : <>
              <p className="error" data-testid="data-part-undetermined">{t('data.dataPart.undeterminedReason', { reason: data?.reason ?? '' })}</p>
              <NewDataPart session={session} onStatus={onStatus} onAdded={(id) => setSelectedId(id)} />
            </>}
      </Fold>
      {(odParts.length > 0) && (
        <Fold id="data.opendope" level={3} title={t('data.opendope.title')} meta={`${odParts.length}`}>
          {od?.xpaths && template.xpaths && (
            <div className="od-part" data-testid="od-xpaths">
              <h4>{t('data.opendope.xpaths')} <span className="meta">{t('data.opendope.xpaths.meta', { part: partLabel(od.xpaths), version: template.xpaths.version ?? t('data.opendope.version.none'), mode: template.mode })}{template.xpaths.authoringTool ? ` · ${template.xpaths.authoringTool} ${template.xpaths.authoringToolVersion ?? ''}` : ''}</span></h4>
              <ul className="od-list">
                {template.xpaths.entries.map((e) => {
                  const u = users.entry(e.id);
                  return <li key={e.id}><b>{e.id}</b> <code>{e.xpath}</code>{e.attrs['name'] ? <span className="meta"> {e.attrs['name']}</span> : null} <EntryUses uses={u} /></li>;
                })}
              </ul>
            </div>
          )}
          {od?.conditions && template.conditions && (
            <div className="od-part" data-testid="od-conditions">
              <h4>{t('data.opendope.conditions')} <span className="meta">{partLabel(od.conditions)}</span></h4>
              <ul className="od-list">
                {template.conditions.conditions.map((c) => {
                  const u = users.condition(c.id);
                  return <li key={c.id}><b>{c.id}</b> <code>{operandText(c.operand as Operand, template.xpaths)}</code>{c.attrs['name'] ? <span className="meta"> {c.attrs['name']}</span> : null}<span className="meta"> · {u.length === 0 ? t('data.unused') : t('data.usedBy', { users: u.join(', ') })}</span></li>;
                })}
              </ul>
            </div>
          )}
          {od?.questions && template.questions && <QuestionsList questions={template.questions} xpaths={template.xpaths} meta={partLabel(od.questions)} />}
          {od?.components && template.components && <ComponentsList components={template.components} meta={partLabel(od.components)} />}
          {template.problems.length > 0 && <ul className="warnings" data-testid="od-problems">{template.problems.map((p) => <li key={`${p.requirement}-${p.message}`}>{p.message} ({p.requirement})</li>)}</ul>}
        </Fold>
      )}
      {others.length > 0 && (
        <Fold id="data.others" level={3} title={t('data.others.title')} meta={`${others.length}`}>
          {others.map((p) => (
            <PartTree key={p.id} session={session} editor={editor} part={p} label={isOpenDopePart(p) ? t('data.others.opendope', { part: partLabel(p) }) : WELL_KNOWN.has(p.namespaceUri) ? t('data.others.microsoft', { part: partLabel(p) }) : partLabel(p)}
              hits={selected === p ? hits : undefined} open={!WELL_KNOWN.has(p.namespaceUri) && !isOpenDopePart(p)} onStatus={onStatus} editable={!isOpenDopePart(p)} />
          ))}
        </Fold>
      )}
    </section>
  );
}

// ---------------------------------------------------------------- a data part made here

/**
 * A document with no data part gets one (section 11.2 item 31): an empty root element named
 * here (in a namespace when one is given), or XML pasted or taken from a file. The part is
 * added to the package as a part edit (no undo); rule 3 names it when it is the only candidate,
 * and the tree above then takes it, ready for children and for bindings.
 */
function NewDataPart({ session, onStatus, onAdded }: { session: EditorDocument; onStatus: (s: string) => void; onAdded: (id: string) => void }) {
  const [root, setRoot] = useState('data');
  const [namespace, setNamespace] = useState('');
  const [xml, setXml] = useState('');
  const add = (text: string, what: string): void => {
    try {
      const part = session.template.addDataPart(text);
      onAdded(part.id);
      onStatus(t('data.new.added', { what, part: partLabel(part) }));
    } catch (e) { onStatus((e as Error).message); }
  };
  const create = (): void => {
    const name = root.trim();
    if (!NAME.test(name)) return onStatus(t('data.new.enterRoot'));
    const ns = namespace.trim();
    add(ns === '' ? `<${name}/>` : `<${name} xmlns="${ns.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"/>`, `<${name}>`);
  };
  const fromFile = (file: File | undefined): void => {
    if (!file) return;
    file.text().then((text) => add(text, file.name), (e: Error) => onStatus(e.message));
  };
  return (
    <div className="new-data-part" data-testid="new-data-part">
      <h4>{t('data.new.title')}</h4>
      <label>{t('data.new.root')} {<input type="text" aria-label={t('data.new.root')} data-testid="new-data-root" value={root} onChange={(e) => setRoot(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); create(); } }} placeholder="invoice" />}</label>
      <label>{t('data.new.namespace')} <input type="text" aria-label={t('data.new.namespace')} data-testid="new-data-namespace" value={namespace} onChange={(e) => setNamespace(e.target.value)} placeholder={t('data.new.namespace.none')} /></label>
      <button type="button" data-testid="new-data-create" onClick={create}>{t('data.new.create')}</button>
      <label>{t('data.new.paste')} <textarea aria-label={t('data.new.paste.label')} data-testid="new-data-xml" rows={4} value={xml} onChange={(e) => setXml(e.target.value)} placeholder="<invoice><customer/></invoice>" /></label>
      <button type="button" data-testid="new-data-add-xml" disabled={xml.trim() === ''} onClick={() => add(xml, t('data.new.theXml'))}>{t('data.new.add')}</button>
      <label>{t('data.new.file')} <input type="file" aria-label={t('data.new.file.label')} data-testid="new-data-file" accept=".xml,text/xml,application/xml" onChange={(e) => { fromFile(e.target.files?.[0]); e.target.value = ''; }} /></label>
    </div>
  );
}

// ---------------------------------------------------------------- one part as a tree

/**
 * The arrow keys walk a tree (ED-001 section 11 item 9's checklist over Template, ED-003 section
 * 8.6): Down and Up between the nodes' names in document order (a collapsed subtree is not in the
 * DOM, so the DOM order is the visible order), Right expands a closed node or steps into an open
 * one, Left collapses an open node or steps out to the parent, Home and End to the tree's ends.
 * The rows keep their own state; the walk drives their twisties.
 */
export function walkTree(e: React.KeyboardEvent<HTMLElement>): void {
  const target = e.target as HTMLElement;
  if (!target.classList.contains('node-name')) return;
  const names = Array.from(e.currentTarget.querySelectorAll<HTMLElement>('.node-name'));
  const i = names.indexOf(target);
  if (i < 0) return;
  const item = target.closest<HTMLElement>('[role="treeitem"]');
  const twisty = item?.querySelector<HTMLButtonElement>(':scope > .tree-row > button.twisty');
  const expanded = item?.getAttribute('aria-expanded');
  let next: HTMLElement | undefined;
  // In a right-to-left chrome the tree opens towards the left: the two horizontal arrows swap (WAI-ARIA's tree pattern).
  const key = isRightToLeft(e.currentTarget) && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') ? (e.key === 'ArrowLeft' ? 'ArrowRight' : 'ArrowLeft') : e.key;
  switch (key) {
    case 'ArrowDown': next = names[i + 1]; break;
    case 'ArrowUp': next = names[i - 1]; break;
    case 'Home': next = names[0]; break;
    case 'End': next = names[names.length - 1]; break;
    case 'ArrowRight':
      if (expanded === 'false') twisty?.click();
      else if (expanded === 'true') next = names[i + 1];
      break;
    case 'ArrowLeft':
      if (expanded === 'true') twisty?.click();
      else next = item?.parentElement?.closest<HTMLElement>('[role="treeitem"]')?.querySelector<HTMLElement>(':scope > .tree-row > .node-name') ?? undefined;
      break;
    default: return;
  }
  e.preventDefault();
  next?.focus();
}

function PartTree({ session, editor, part, label, hits, open, onStatus, testId, editable }: { session: EditorDocument; editor: DocxEditorView; part: CustomXmlPart; label: string; hits: Set<string> | undefined; open: boolean; onStatus: (s: string) => void; testId?: string; editable: boolean }) {
  const [expanded, setExpanded] = useState(open);
  // An editable part is shown as a tree or, for editing the XML itself, in Monaco (Jason, ED-005 section 12.64).
  const [view, setView] = useState<PartView>(() => (editable ? storedView() : 'tree'));
  const choose = (v: PartView): void => { setView(v); try { localStorage.setItem(VIEW_KEY, v); } catch { /* the choice is not kept */ } };
  /** The node last selected in the tree, which Copy XPath copies (kept when the focus moves to the button). */
  const [current, setCurrent] = useState<CustomXmlNode | null>(null);
  const root = part.documentElement;
  const id = testId ?? 'part';
  return (
    <div className="part-tree" data-testid={testId}>
      <div className="tree-row">
        <button type="button" className="twisty" aria-label={expanded ? t('data.tree.collapse') : t('data.tree.expand')} aria-expanded={expanded} onClick={() => setExpanded((e) => !e)}>{expanded ? '▾' : '▸'}</button>
        <span className="part-label">{label}</span>
        {expanded && view === 'tree' && (
          <button type="button" className="link copy-xpath" data-testid={`${id}-copy-xpath`} disabled={!current} title={t('data.node.copy.title')}
            aria-label={current ? t('data.node.copy.label', { xpath: current.xpath }) : t('data.node.copy')} onClick={() => { if (current) void copyXPath(current, onStatus); }}>{t('data.node.copy')}</button>
        )}
        {editable && expanded && (
          <span className="view-switch" role="group" aria-label={t('data.view.label')}>
            <button type="button" className="link" aria-pressed={view === 'tree'} data-testid={`${id}-view-tree`} onClick={() => choose('tree')}>{t('data.view.tree')}</button>
            <button type="button" className="link" aria-pressed={view === 'xml'} data-testid={`${id}-view-xml`} onClick={() => choose('xml')}>{t('data.view.xml')}</button>
          </span>
        )}
      </div>
      {expanded && (view === 'xml' && editable
        ? <PartXml session={session} editor={editor} part={part} onStatus={onStatus} testId={`${id}-xml`} />
        : <ul role="tree" aria-label={label} onKeyDown={walkTree}><NodeRow session={session} editor={editor} part={part} node={root} depth={0} hits={hits} onStatus={onStatus} editable={editable} onSelect={setCurrent} /></ul>)}
    </div>
  );
}

/**
 * A part's XML in Monaco with Apply, Revert and Format, as the parts panel edits it (ED-005 section 12.64); an Apply
 * refreshes the bound controls, as an edit in the tree does. The context menu's Bind and Copy XPath act on the element
 * (or attribute) at the cursor, in the part as applied: with unapplied changes, the positions would not be the part's.
 */
function PartXml({ session, editor, part, onStatus, testId }: { session: EditorDocument; editor: DocxEditorView; part: CustomXmlPart; onStatus: (s: string) => void; testId: string }) {
  useT();
  const editing = useMemo(() => partEditing(part.part), [part]);
  const [initial, setInitial] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    void editableText(part.part, editing).then((text) => { if (!cancelled) setInitial(text); });
    return () => { cancelled = true; };
  }, [part, editing]);
  /** The node at an offset of the editor's text, when that text is the part's; else the status line says why not. */
  const nodeAt = useCallback(async (offset: number, text: string): Promise<CustomXmlNode | undefined> => {
    if (text !== await editableText(part.part, editing)) { onStatus(t('data.xml.notApplied')); return undefined; }
    const where = elementPathAt(text, offset);
    const element = where ? part.selectSingleNode(`/${where.path.map((i) => `*[${i}]`).join('/')}`) : undefined;
    if (!where || !element) { onStatus(t('data.xml.noNode')); return undefined; }
    return (where.attribute ? element.attributes.find((a) => (a.node as Attr).name === where.attribute) : undefined) ?? element;
  }, [part, editing, onStatus]);
  const actions = useMemo<XmlContextAction[]>(() => [
    { id: 'bind', label: t('data.xml.bind'), run: (offset, text) => { void nodeAt(offset, text).then((node) => { if (node) bindNode(editor, session, part, node, onStatus); }); } },
    { id: 'copyXPath', label: t('data.xml.copy'), run: (offset, text) => { void nodeAt(offset, text).then((node) => { if (node) void copyXPath(node, onStatus); }); } },
  ], [editor, part, onStatus, nodeAt]);
  if (initial === null) return <p className="meta">{t('tab.loading')}</p>;
  return <PartEditor session={session} part={part.part} editing={editing} initial={initial} testId={testId} contextActions={actions}
    onApplied={() => { applyBoundRefresh(session, editor.state); void Promise.all([refreshXhtmlControls(editor.view, session), refreshPictureControls(editor.view, session), refreshFragmentControls(editor, session)]).then((results) => { for (const note of new Set(results.flatMap((r) => r.notes))) onStatus(note); }); }} />;
}

function dragData(part: CustomXmlPart, node: CustomXmlNode): BindDrag {
  return { xpath: node.xpath, prefixMappings: node.prefixMappings, storeItemID: part.id };
}

/** The drag type a node travels under within the tree (a move); the bind type travels beside it for a drop on the page. */
const MOVE_MIME = 'application/x-docx4j-move';

type Zone = 'before' | 'after' | 'into';

/** An XML name: `item`, `ns0:item`. */
const NAME = /^[A-Za-z_][\w.-]*(:[A-Za-z_][\w.-]*)?$/;

/** The DOM element behind a node wrapper (the wrapper's `node` is the DOM node; the tree is edited through the DOM and the part told). */
function elementOf(node: CustomXmlNode): Element {
  return node.node as Element;
}

/** A new element beside or under `context`, in its namespace with its prefix, so a data part's children stay in the part's namespace. */
function newElement(context: Element, name: string): Element {
  const doc = context.ownerDocument;
  if (name.includes(':')) {
    const prefix = name.slice(0, name.indexOf(':'));
    const ns = context.lookupNamespaceURI(prefix);
    if (!ns) throw new Error(t('data.edit.undeclaredPrefix', { prefix }));
    return doc.createElementNS(ns, name);
  }
  const ns = context.namespaceURI;
  const prefix = context.prefix;
  return ns ? doc.createElementNS(ns, prefix ? `${prefix}:${name}` : name) : doc.createElement(name);
}

type Form = { kind: 'child'; name: string; value: string; position: 'last' | 'first' } | { kind: 'before' | 'after'; name: string; value: string } | { kind: 'rename'; name: string } | { kind: 'attribute'; name: string; value: string } | { kind: 'attributeValue'; name: string; value: string } | { kind: 'value'; value: string };

function NodeRow({ session, editor, part, node, depth, hits, onStatus, editable, onSelect }: { session: EditorDocument; editor: DocxEditorView; part: CustomXmlPart; node: CustomXmlNode; depth: number; hits: Set<string> | undefined; onStatus: (s: string) => void; editable: boolean; onSelect: (node: CustomXmlNode) => void }) {
  const children = node.childElements;
  const [expanded, setExpanded] = useState(depth < 3);
  const [form, setForm] = useState<Form | null>(null);
  const [selected, setSelected] = useState(false);
  const [zone, setZone] = useState<Zone | null>(null);
  const xpath = node.xpath;
  const hit = hits?.has(xpath) ?? false;
  const leaf = children.length === 0;
  const el = elementOf(node);
  // A leaf holding image data is shown as a thumbnail beside its value (ED-005 section 12.117), and binds as a picture control.
  const text = leaf ? node.text : '';
  const image = useMemo(() => (text === '' ? undefined : pictureFactsOf(text.trim())), [text]);
  // A leaf holding a Word document as Flat OPC is said for what it is, read-only, and binds as Word's mapped rich text (ED-005
  // proposal 46): its value is tens of kilobytes of escaped XML, which the value box showed raw and let be typed over.
  const fragment = text !== '' && isFlatOpc(text);
  const fragmentFacts = useFragmentFacts(fragment ? text : undefined);
  const onDragStart = (e: DragEvent<HTMLElement>): void => {
    e.dataTransfer.setData(BIND_MIME, JSON.stringify(dragData(part, node)));
    if (editable && node.parentNode) e.dataTransfer.setData(MOVE_MIME, JSON.stringify({ storeItemID: part.id, xpath, prefixMappings: node.prefixMappings }));
    e.dataTransfer.setData('text/plain', xpath);
    e.dataTransfer.effectAllowed = 'linkMove';
  };
  /** Where a drop on this row would put the dragged node: the upper quarter before it, the lower quarter after it, the middle inside it as the last child. */
  const zoneOf = (e: DragEvent<HTMLElement>): Zone => {
    const rect = e.currentTarget.getBoundingClientRect();
    const y = (e.clientY - rect.top) / Math.max(1, rect.height);
    return y < 0.25 ? 'before' : y > 0.75 ? 'after' : 'into';
  };
  const onDragOver = (e: DragEvent<HTMLElement>): void => {
    if (!editable || !Array.from(e.dataTransfer.types).includes(MOVE_MIME)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const z = zoneOf(e);
    if (z !== zone) setZone(z);
  };
  const onDrop = (e: DragEvent<HTMLElement>): void => {
    const data = e.dataTransfer.getData(MOVE_MIME);
    setZone(null);
    if (!data) return;
    e.preventDefault();
    e.stopPropagation();
    try {
      const from = JSON.parse(data) as { storeItemID: string; xpath: string; prefixMappings: string };
      if (from.storeItemID !== part.id) return fail(t('data.edit.ownPart'));
      const source = part.selectSingleNode(from.xpath, from.prefixMappings || undefined);
      if (!source) return fail(t('data.edit.gone', { xpath: from.xpath }));
      const src = elementOf(source);
      if (src === el || src.contains(el)) return fail(t('data.edit.intoItself'));
      const z = zoneOf(e);
      if (z === 'into') el.appendChild(src);
      else if (z === 'before') el.parentNode!.insertBefore(src, el);
      else el.parentNode!.insertBefore(src, el.nextSibling);
      if (z === 'into') setExpanded(true);
      done(t(z === 'into' ? 'data.edit.movedInto' : z === 'before' ? 'data.edit.movedBefore' : 'data.edit.movedAfter', { from: from.xpath, to: xpath }));
    } catch (err) { fail((err as Error).message); }
  };
  const move = (by: -1 | 1): void => {
    const sibling = by < 0 ? el.previousElementSibling : el.nextElementSibling;
    if (!sibling || !el.parentNode) return;
    el.parentNode.insertBefore(el, by < 0 ? sibling : sibling.nextSibling);
    done(t(by < 0 ? 'data.edit.movedUp' : 'data.edit.movedDown', { xpath }));
  };
  const done = (what: string): void => {
    part.touch();
    partEdited(session, part, editor, onStatus);
    setForm(null);
    onStatus(t('data.edit.done', { what }));
  };
  const fail = (message: string): void => { onStatus(message); };

  const apply = (): void => {
    if (!form) return;
    try {
      switch (form.kind) {
        case 'child': case 'before': case 'after': {
          const name = form.name.trim();
          if (!NAME.test(name)) return fail(t('data.edit.enterElement'));
          const parent = form.kind === 'child' ? el : el.parentElement;
          if (!parent) return fail(t('data.edit.noSiblings'));
          const made = newElement(form.kind === 'child' ? el : parent, name);
          if (form.value !== '') made.textContent = form.value;
          if (form.kind === 'child') { if (form.position === 'first') el.insertBefore(made, el.firstChild); else el.appendChild(made); setExpanded(true); }
          else parent.insertBefore(made, form.kind === 'before' ? el : el.nextSibling);
          return done(form.kind === 'child' ? t('data.edit.addedChild', { name, position: form.position, xpath }) : t(form.kind === 'before' ? 'data.edit.addedBefore' : 'data.edit.addedAfter', { name, xpath }));
        }
        case 'rename': {
          const name = form.name.trim();
          if (!NAME.test(name)) return fail(t('data.edit.enterElement'));
          if (name === el.tagName) { setForm(null); return; }
          const made = newElement(el, name);
          for (const a of Array.from(el.attributes)) made.setAttributeNS(a.namespaceURI, a.name, a.value);
          while (el.firstChild) made.appendChild(el.firstChild);
          if (!el.parentNode) return fail(t('data.edit.cannotRename'));
          el.parentNode.replaceChild(made, el);
          return done(t('data.edit.renamed', { xpath, name }));
        }
        case 'attribute': case 'attributeValue': {
          const name = form.name.trim();
          if (!NAME.test(name)) return fail(t('data.edit.enterAttribute'));
          if (name.includes(':')) {
            const ns = el.lookupNamespaceURI(name.slice(0, name.indexOf(':')));
            if (!ns) return fail(t('data.edit.undeclaredPrefix', { prefix: name.slice(0, name.indexOf(':')) }));
            el.setAttributeNS(ns, name, form.value);
          } else el.setAttribute(name, form.value);
          return done(t('data.edit.attributeSet', { name, xpath }));
        }
        case 'value': {
          if (form.value === node.text) { setForm(null); return; }
          node.text = form.value;
          return done(t('data.edit.valueSet', { xpath, value: form.value }));
        }
      }
    } catch (e) { fail((e as Error).message); }
  };
  const duplicate = (): void => {
    if (!el.parentNode) return fail(t('data.edit.cannotDuplicate'));
    el.parentNode.insertBefore(el.cloneNode(true), el.nextSibling);
    done(t('data.edit.duplicated', { xpath }));
  };
  const remove = (): void => {
    if (!node.parentNode) return fail(t('data.edit.cannotRemove'));
    node.delete();
    partEdited(session, part, editor, onStatus);
    onStatus(t('data.edit.removed', { xpath }));
  };
  const removeAttribute = (name: string): void => {
    el.removeAttribute(name);
    done(t('data.edit.attributeRemoved', { name, xpath }));
  };
  /**
   * The keyboard's bind (section 8.6's checklist): the control at the caret bound to this node, as a drop
   * on it would. With no control there, or inside one that holds controls (rich text, a group, a
   * repeating section), a plain text control made at the caret, or around a selection within a
   * paragraph, and bound in one step, as Word's XML Mapping Pane inserts one (ED-005 sections 12.56 and
   * 12.57); a selection that cannot hold one is refused, saying why, as is a node mapped to a rich text
   * control, which Word maps to nothing else (section 12.58).
   */
  const bind = (): void => bindNode(editor, session, part, node, onStatus);

  const field = (label: string, value: string, onChange: (v: string) => void, testId: string, placeholder?: string, autoFocus = false): ReactNode => (
    <input type="text" aria-label={label} data-testid={testId} value={value} placeholder={placeholder} autoFocus={autoFocus} onChange={(e) => onChange(e.target.value)}
      onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); apply(); } else if (e.key === 'Escape') setForm(null); }} />
  );

  return (
    <li role="treeitem" aria-expanded={leaf ? undefined : expanded} aria-selected={selected} aria-label={leaf ? t('data.node.leafLabel', { name: node.baseName, value: node.text === '' ? t('data.node.empty') : fragment ? fragmentSummary(fragmentFacts) : node.text }) : node.baseName}
      className={`data-node${hit ? ' hit' : ''}${selected ? ' selected' : ''}`} data-xpath={xpath}>
      <div className={`tree-row${zone ? ` drop-${zone}` : ''}`} onClick={() => { setSelected(true); onSelect(node); }} onBlur={(e) => { if (!e.currentTarget.parentElement?.contains(e.relatedTarget as Node | null)) setSelected(false); }}
        onDragOver={onDragOver} onDragLeave={() => setZone(null)} onDrop={onDrop}>
        {leaf ? <span className="twisty" aria-hidden="true" /> : <button type="button" className="twisty" tabIndex={-1} aria-label={expanded ? t('data.tree.collapse') : t('data.tree.expand')} onClick={() => setExpanded((x) => !x)}>{expanded ? '▾' : '▸'}</button>}
        {/* The root's name is the tree's tab stop; the arrow keys reach the rest (walkTree), and a selected node keeps its stop so Tab comes back to it. */}
        <span className="node-name" draggable onDragStart={onDragStart} title={t('data.node.title', { xpath })} tabIndex={depth === 0 || selected ? 0 : -1} onFocus={() => { setSelected(true); onSelect(node); }}
          onKeyDown={(e) => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') { e.preventDefault(); void copyXPath(node, onStatus); } }}>{node.baseName}</span>
        {node.attributes.map((a) => (
          <span key={a.baseName} className="attribute">
            {' '}<button type="button" className="attr link" title={editable ? t('data.node.editAttribute', { name: a.baseName }) : undefined} disabled={!editable} onClick={() => setForm({ kind: 'attributeValue', name: (a.node as Attr).name, value: a.nodeValue })}>
              <span className="attr-name">@{(a.node as Attr).name}</span>=<span className="attr-value">"{a.nodeValue}"</span>
            </button>
            {editable && <button type="button" className="link attr-remove" aria-label={t('data.node.removeAttribute', { name: (a.node as Attr).name, xpath })} title={t('data.node.removeAttribute.title')} onClick={() => removeAttribute((a.node as Attr).name)}>×</button>}
          </span>
        ))}
        {image && (
          <img className="node-thumbnail" data-testid="node-thumbnail" src={imageDataUrl(image.contentType, text)} alt={t('picture.alt', { xpath })}
            title={image.format === 'SVG' ? t('picture.sizeSvg') : t('picture.size', { format: image.format, width: image.widthPx, height: image.heightPx })} />
        )}
        {fragment && <span className="node-fragment" data-testid="node-fragment" title={t('data.node.fragment.title')}>{fragmentSummary(fragmentFacts)}</span>}
        {leaf && !fragment && (form?.kind === 'value'
          ? field(t('data.node.valueOf', { xpath }), form.value, (v) => setForm({ kind: 'value', value: v }), 'node-value-input', undefined, true)
          : <button type="button" className="node-value" title={editable ? t('data.node.editValue') : undefined} disabled={!editable} onClick={() => setForm({ kind: 'value', value: node.text })}>{node.text === '' ? <span className="meta">{t('data.node.empty')}</span> : node.text}</button>)}
        {form?.kind === 'value' && <button type="button" className="link" onClick={apply}>{t('data.node.ok')}</button>}
        <span className="node-actions">
          <button type="button" className="link" title={t(image ? 'data.node.bind.titlePicture' : fragment ? 'data.node.bind.titleFragment' : 'data.node.bind.title')} aria-label={t(image ? 'data.node.bind.labelPicture' : fragment ? 'data.node.bind.labelFragment' : 'data.node.bind.label', { xpath })} onClick={bind}>{t('data.node.bind')}</button>
          {editable && <>
            <button type="button" className="link" title={t('data.node.child.title')} aria-label={t('data.node.child.label', { xpath })} onClick={() => setForm({ kind: 'child', name: '', value: '', position: 'last' })}>{t('data.node.child')}</button>
            {node.parentNode && <button type="button" className="link" title={t('data.node.before.title')} aria-label={t('data.node.before.label', { xpath })} onClick={() => setForm({ kind: 'before', name: el.tagName, value: '' })}>{t('data.node.before')}</button>}
            {node.parentNode && <button type="button" className="link" title={t('data.node.after.title')} aria-label={t('data.node.after.label', { xpath })} onClick={() => setForm({ kind: 'after', name: el.tagName, value: '' })}>{t('data.node.after')}</button>}
            {node.parentNode && <button type="button" className="link" title={t('data.node.up.title')} aria-label={t('data.node.up.label', { xpath })} disabled={!el.previousElementSibling} onClick={() => move(-1)}>↑</button>}
            {node.parentNode && <button type="button" className="link" title={t('data.node.down.title')} aria-label={t('data.node.down.label', { xpath })} disabled={!el.nextElementSibling} onClick={() => move(1)}>↓</button>}
            {node.parentNode && <button type="button" className="link" title={t('data.node.duplicate.title')} aria-label={t('data.node.duplicate.label', { xpath })} onClick={duplicate}>{t('data.node.duplicate')}</button>}
            <button type="button" className="link" title={t('data.node.attribute.title')} aria-label={t('data.node.attribute.label', { xpath })} onClick={() => setForm({ kind: 'attribute', name: '', value: '' })}>{t('data.node.attribute')}</button>
            {node.parentNode && <button type="button" className="link" title={t('data.node.rename.title')} aria-label={t('data.node.rename.label', { xpath })} onClick={() => setForm({ kind: 'rename', name: el.tagName })}>{t('data.node.rename')}</button>}
            {node.parentNode && <button type="button" className="link" title={t('data.node.remove.title')} aria-label={t('data.node.remove.label', { xpath })} onClick={remove}>×</button>}
          </>}
        </span>
      </div>
      {form && form.kind !== 'value' && (
        <div className="tree-row node-form" data-testid="node-form">
          {form.kind === 'child' && <>
            <select aria-label={t('data.form.position')} data-testid="node-form-position" value={form.position} onChange={(e) => setForm({ ...form, position: e.target.value as 'last' | 'first' })}><option value="last">{t('data.form.lastChild')}</option><option value="first">{t('data.form.firstChild')}</option></select>
            {field(t('data.form.newName'), form.name, (v) => setForm({ ...form, name: v }), 'node-form-name', t('data.form.elementName'), true)} = {field(t('data.form.newValue'), form.value, (v) => setForm({ ...form, value: v }), 'node-form-value', t('data.form.valueOptional'))}</>}
          {(form.kind === 'before' || form.kind === 'after') && <>{t(form.kind === 'before' ? 'data.form.before' : 'data.form.after', { name: node.baseName })} {field(t('data.form.newName'), form.name, (v) => setForm({ ...form, name: v }), 'node-form-name', t('data.form.elementName'), true)} = {field(t('data.form.newValue'), form.value, (v) => setForm({ ...form, value: v }), 'node-form-value', t('data.form.valueOptional'))}</>}
          {form.kind === 'rename' && <>{t('data.form.renameTo')} {field(t('data.form.newNameRename'), form.name, (v) => setForm({ kind: 'rename', name: v }), 'node-form-name', t('data.form.elementName'), true)}</>}
          {form.kind === 'attribute' && <>@{field(t('data.form.attributeName'), form.name, (v) => setForm({ ...form, name: v }), 'node-form-name', t('data.form.name'), true)} = {field(t('data.form.attributeValue'), form.value, (v) => setForm({ ...form, value: v }), 'node-form-value', t('data.form.value'))}</>}
          {form.kind === 'attributeValue' && <>@{form.name} = {field(t('data.form.valueOfAttribute', { name: form.name }), form.value, (v) => setForm({ ...form, value: v }), 'node-form-value', t('data.form.value'), true)}</>}
          <button type="button" data-testid="node-form-apply" onClick={apply}>{t('data.form.apply')}</button>
          <button type="button" className="link" onClick={() => setForm(null)}>{t('data.form.cancel')}</button>
        </div>
      )}
      {!leaf && expanded && <ul role="group">{children.map((c, i) => <NodeRow key={`${c.baseName}-${i}`} session={session} editor={editor} part={part} node={c} depth={depth + 1} hits={hits} onStatus={onStatus} editable={editable} onSelect={onSelect} />)}</ul>}
    </li>
  );
}
