// SPDX-License-Identifier: GPL-3.0-or-later
//
// The parts as a tree (ED-001 section 3.6's Developer row; ED-002 section 6.4, reworked
// 2026-09-24 at Jason's request): the package relationships at the root, each relationship a
// branch to the part it targets and that part's own relationships below it, external targets
// as leaves; parts no relationship reaches are listed last. A click on a part shows, between
// the tree and the part's content, the XML of its relationships part when it has one, then the
// part's content: an XML part as it would be saved (the source bytes of a part nobody
// unmarshalled, the marshalled tree of one somebody did, so the pane says which), an image as
// the image, any other binary part as its size. Since E4.b (ED-005 section 4) a package part is
// edited here in Monaco with Apply and Revert: styles, numbering, settings, theme, fonts and
// `docProps` typed and checked that nothing would be lost, a custom XML or untyped part as XML
// that parses; the projected parts, the comments, the relationships and the binaries stay read
// only, each saying where it is edited instead. A part edit has no undo: Revert is the pane's.
import { Suspense, lazy, useEffect, useMemo, useState } from 'react';
import { XmlPart, DefaultXmlPart, BinaryPart, RelationshipsPart, isExternal, type Part } from '@docx4j/core-ts';
import { partEditing, type EditorDocument, type PartEditing } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { t, useT } from '../i18n.js';
import { pretty } from '../xml.js';
import { PartEditor, editableText } from './PartXmlEditor.js';

const XmlCodeEditor = lazy(() => import('./XmlCodeEditor.js'));

type Kind = 'xml' | 'binary' | 'other';

export interface PartNode {
  /** The part name, or the external target. */
  name: string;
  /** The relationship that led here; none at the root. */
  relId?: string;
  relType?: string;
  external: boolean;
  part?: Part;
  kind: Kind;
  live: boolean;
  contentType: string;
  /** The part appeared earlier in the tree (a theme two parts share): its branches are not repeated. */
  repeated: boolean;
  children: PartNode[];
}

function kindOf(part: Part): Kind {
  return part instanceof XmlPart || part instanceof DefaultXmlPart ? 'xml' : part instanceof BinaryPart ? 'binary' : 'other';
}

/** Whether a part is written from its tree at save, not from its source bytes. */
function isLive(part: Part): boolean {
  return part instanceof XmlPart && part.isUnmarshalled;
}

/** The last segment of a relationship type: `officeDocument`, `styles`, `hyperlink`. */
export function relTypeLabel(type: string | undefined): string {
  if (!type) return '';
  return type.slice(type.lastIndexOf('/') + 1);
}

function shortContentType(ct: string): string {
  return ct.replace(/^application\/vnd\.openxmlformats-(officedocument|package)\./, '').replace(/^application\/vnd\.(ms-|openxmlformats-)/, '');
}

/** The tree from the package relationships; `seen` keeps a shared part from being expanded twice. */
export function partsTree(session: EditorDocument): { root: PartNode[]; orphans: PartNode[] } {
  const pkg = session.pkg;
  const seen = new Set<string>();
  const visit = (rp: RelationshipsPart | undefined): PartNode[] => {
    if (!rp) return [];
    const out: PartNode[] = [];
    for (const rel of rp.list) {
      if (isExternal(rel)) {
        out.push({ name: rel.target, relId: rel.id, relType: rel.type, external: true, kind: 'other', live: false, contentType: '', repeated: false, children: [] });
        continue;
      }
      const part = rp.getPart(rel);
      const name = part ? part.partName.toString() : rp.resolveTarget(rel).toString();
      if (!part) {
        out.push({ name, relId: rel.id, relType: rel.type, external: false, kind: 'other', live: false, contentType: '(missing)', repeated: false, children: [] });
        continue;
      }
      const repeated = seen.has(name);
      seen.add(name);
      out.push({
        name, relId: rel.id, relType: rel.type, external: false, part, kind: kindOf(part), live: isLive(part),
        contentType: part.contentType, repeated, children: repeated ? [] : visit(part.relationshipsPart),
      });
    }
    return out;
  };
  const root = visit(pkg.relationshipsPart);
  const orphans: PartNode[] = [];
  for (const part of pkg.parts) {
    const name = part.partName.toString();
    if (seen.has(name) || part instanceof RelationshipsPart) continue;
    orphans.push({ name, external: false, part, kind: kindOf(part), live: isLive(part), contentType: part.contentType, repeated: false, children: [] });
  }
  orphans.sort((a, b) => a.name.localeCompare(b.name));
  return { root, orphans };
}

interface Content {
  /** The part the content was read from: the pane shows it only for that part (the read is asynchronous, the selection is not). */
  part?: Part;
  text?: string;
  size?: number;
  image?: string;
  rels?: string;
}

export function PartsPanel({ session, editor }: { session: EditorDocument; editor?: DocxEditorView }) {
  useT();
  /** Bumped by an Apply and by a part edit arriving: the tree says which parts are live. */
  const [revision, setRevision] = useState(0);
  useEffect(() => session.onPartsReceived(() => setRevision((r) => r + 1)), [session]);
  const tree = useMemo(() => { void revision; return partsTree(session); }, [session, revision]);
  const [selected, setSelected] = useState<PartNode | null>(null);
  const [content, setContent] = useState<Content>({});
  const [collapsed, setCollapsed] = useState<Set<string>>(() => new Set());

  useEffect(() => { setSelected(null); setCollapsed(new Set()); }, [session]);

  const editing: PartEditing | undefined = selected?.part ? partEditing(selected.part) : undefined;

  useEffect(() => {
    if (!selected?.part) { setContent({}); return; }
    const part = selected.part;
    let url: string | undefined;
    let cancelled = false;
    setContent({});
    void (async () => {
      const next: Content = { part };
      const rp = part.relationshipsPart;
      if (rp) {
        try { next.rels = pretty(new TextDecoder().decode(await rp.getBytes())); } catch (e) { next.rels = t('xml.notAvailable', { error: (e as Error).message }); }
      }
      const bytes = await part.getBytes();
      next.size = bytes.length;
      const mode = partEditing(part);
      if (selected.kind === 'xml') next.text = mode.mode !== 'none' ? await editableText(part, mode) : pretty(new TextDecoder().decode(bytes));
      else if (part.contentType.startsWith('image/') && typeof URL.createObjectURL === 'function') {
        url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: part.contentType }));
        next.image = url;
      }
      if (!cancelled) setContent(next);
    })();
    return () => { cancelled = true; if (url) URL.revokeObjectURL(url); };
  }, [selected]);

  const toggle = (key: string): void => setCollapsed((c) => { const n = new Set(c); if (n.has(key)) n.delete(key); else n.add(key); return n; });

  const renderNodes = (nodes: PartNode[], path: string): React.ReactNode => (
    <ul role="group">
      {nodes.map((n, i) => {
        const key = `${path}/${n.relId ?? i}`;
        const open = !collapsed.has(key);
        const branch = n.children.length > 0;
        return (
          <li key={key} role="treeitem" aria-expanded={branch ? open : undefined} aria-selected={selected === n} className={`${n.external ? 'external' : ''}${n.repeated ? ' repeated' : ''}`}>
            <div className="tree-row">
              {branch
                ? <button type="button" className="twisty" aria-label={open ? t('parts.collapse') : t('parts.expand')} onClick={() => toggle(key)}>{open ? '▾' : '▸'}</button>
                : <span className="twisty" aria-hidden="true" />}
              {n.part
                ? <button type="button" className={`link${selected === n ? ' current' : ''}`} onClick={() => setSelected(n)}>{n.name}</button>
                : <span className="target">{n.name}</span>}
              <span className="meta">
                {n.relId ? `${n.relId} · ${relTypeLabel(n.relType)}` : ''}
                {n.external ? t('parts.external') : ''}
                {n.repeated ? t('parts.repeated') : ''}
                {n.live ? t('parts.live') : ''}
              </span>
            </div>
            {branch && open && renderNodes(n.children, key)}
          </li>
        );
      })}
    </ul>
  );

  return (
    <section className="panel parts-panel" aria-label={t('tab.parts')} data-testid="panel-parts">
      <h2>{t('tab.parts')} <span className="meta">{t('parts.meta')}</span></h2>
      <div className="parts-tree" role="tree" aria-label={t('parts.tree')} data-testid="parts-tree">
        {renderNodes(tree.root, '')}
        {tree.orphans.length > 0 && (
          <>
            <h3>{t('parts.orphans')}</h3>
            {renderNodes(tree.orphans, 'orphan')}
          </>
        )}
      </div>
      {selected?.part && content.part === selected.part && content.rels !== undefined && (
        <div className="part-content part-rels" data-testid="part-rels">
          <h3>{selected.part.relationshipsPart?.partName.toString()} <span className="meta">{t('parts.rels', { name: selected.name })}</span></h3>
          <Suspense fallback={<p className="meta">{t('tab.loading')}</p>}>
            <XmlCodeEditor value={content.rels} readOnly compact testId="part-rels-xml" ariaLabel={t('parts.editor', { name: selected.part.relationshipsPart?.partName.toString() ?? '' })} />
          </Suspense>
        </div>
      )}
      {selected?.part && (
        <div className="part-content" data-testid="part-content">
          <h3>{selected.name} <span className="meta">{t(selected.live ? 'parts.bytes.live' : 'parts.bytes.file', { type: shortContentType(selected.contentType), size: content.size ?? '…' })}</span></h3>
          {content.part !== selected.part
            ? <p className="meta">{t('tab.loading')}</p>
            : editing && editing.mode !== 'none' && content.text !== undefined
            ? <PartEditor key={selected.name} session={session} part={selected.part} editing={editing} initial={content.text}
                onApplied={() => { editor?.partsChanged(); setRevision((r) => r + 1); }} />
            : (
              <>
                {editing?.mode === 'none' && <p className="meta" data-testid="part-readonly">{t(`parts.readonly.${editing.reason}`)}</p>}
                {content.text !== undefined && (
                  // Read-only Monaco (Jason, 2026-10-04): the editable parts' colours, find and folding, and only the lines in
                  // view drawn, which a 4 MB main part needs (ED-005 section 12.63).
                  <Suspense fallback={<p className="meta">{t('tab.loading')}</p>}>
                    <XmlCodeEditor value={content.text} readOnly testId="part-readonly-xml" ariaLabel={t('parts.editor', { name: selected.name })} />
                  </Suspense>
                )}
                {content.image && <img src={content.image} alt={selected.name} />}
              </>
            )}
        </div>
      )}
    </section>
  );
}
