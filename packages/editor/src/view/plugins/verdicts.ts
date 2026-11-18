// SPDX-License-Identifier: GPL-3.0-or-later
//
// Verdict badges (ED-003 section 4.7): for every control with an OpenDoPE role, or a Word
// repeating section, what a processor would do, as node decorations the stylesheet draws as
// badges (`data-verdict` on a block or run control, `data-sdt-verdict` on the first cell of a
// row or cell group) on the page and on the preview alike. One decoration per control with a
// role, never one per block (ED-002 section 10.5 item 10). A verdict is one XPath per control
// and per condition leaf, so they are cached by the control's properties object and dropped
// when the template view changes (a part edit, the data part, the load).
import { touches } from './labels.js';
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { t } from '../../app/i18n.js';
import { LANGUAGE_CHANGED } from './decorations.js';
import { listControls, rowPosition, type EditorDocument, type Verdict, type ControlLookup } from '@docx4j/editor-model';

export const verdictsKey = new PluginKey<DecorationSet>('docx4j-verdicts');

/** The meta a transaction carries when the template view changed (the view dispatches it on `template.subscribe`). */
export const TEMPLATE_CHANGED = 'docx4j-template-changed';

/** The badge's text: what the processor does with the control. */
export function badgeText(v: Verdict): string | undefined {
  if (v.error) return t('view.verdict.error', { error: v.error.length > 60 ? `${v.error.slice(0, 60)}…` : v.error });
  if (v.kept !== undefined) return v.kept ? t('verdict.kept') : t('verdict.dropped');
  if (v.count !== undefined) return v.count === 0 ? t('view.verdict.removed') : t('verdict.instances', { count: v.count });
  if (v.positionPredicate !== undefined) return t('verdict.position', { position: v.positionPredicate });
  if (v.role === 'component') return v.id ? t('view.verdict.component', { id: v.id }) : t('view.verdict.componentBare');
  return undefined;
}

export class VerdictCache {
  private cache = new WeakMap<object, Verdict>();

  constructor(private readonly session: EditorDocument) {}

  clear(): void {
    this.cache = new WeakMap();
  }

  verdictFor(sdtPr: unknown): Verdict | undefined {
    if (!sdtPr || typeof sdtPr !== 'object' || !this.session.template.loaded) return undefined;
    let v = this.cache.get(sdtPr);
    if (!v) {
      try { v = this.session.template.verdictFor(sdtPr); } catch (e) { v = { role: 'none', warnings: [], error: (e as Error).message }; }
      this.cache.set(sdtPr, v);
    }
    return v;
  }
}

export function verdictDecorations(doc: PMNode, controls: ControlLookup, cache: VerdictCache): DecorationSet {
  const decos: Decoration[] = [];
  for (const c of listControls(doc, controls)) {
    const v = cache.verdictFor(c.sdtPr);
    const text = v ? badgeText(v) : undefined;
    if (!text) continue;
    const attrs: Record<string, string> = { 'data-verdict': text };
    if (v!.kept !== undefined) attrs['data-verdict-kept'] = String(v!.kept);
    if (v!.count === 0) attrs['data-verdict-kept'] = 'false';
    if (v!.error) attrs['data-verdict-error'] = 'true';
    if (c.pos !== undefined) {
      const node = doc.nodeAt(c.pos)!;
      decos.push(Decoration.node(c.pos, c.pos + node.nodeSize, attrs, { verdict: text }));
    } else if (c.tablePos !== undefined && c.rows) {
      // The first cell of the group's first row (a row group), or of its cells (a cell group).
      const table = doc.nodeAt(c.tablePos)!;
      const row = table.child(c.rows[0]);
      const rowPos = rowPosition(table, c.tablePos, c.rows[0]);
      const first = c.cells ? c.cells[0] : 0;
      let cellPos = rowPos + 1;
      for (let i = 0; i < first; i++) cellPos += row.child(i).nodeSize;
      const cellAttrs: Record<string, string> = { 'data-sdt-verdict': text };
      if (attrs['data-verdict-kept']) cellAttrs['data-verdict-kept'] = attrs['data-verdict-kept']!;
      decos.push(Decoration.node(cellPos, cellPos + row.child(first).nodeSize, cellAttrs, { verdict: text }));
    }
  }
  return DecorationSet.create(doc, decos);
}

/** The badges on a document: the body (the controls map for groups) or the preview (its own lookup). */
export function verdictsPlugin(session: EditorDocument, controls: ControlLookup, cache: VerdictCache = new VerdictCache(session)): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: verdictsKey,
    state: {
      init: (_config, state: EditorState) => verdictDecorations(state.doc, controls, cache),
      apply: (tr, set, _old, state) => {
        if (tr.getMeta(TEMPLATE_CHANGED)) cache.clear();
        if (tr.getMeta(LANGUAGE_CHANGED)) return verdictDecorations(state.doc, controls, cache);
        // A keystroke outside every control changes no verdict: the whole-document listing runs only when one may have (ED-003 section 8.2).
        const inControl = (n: PMNode): boolean => n.type.name === 'sdtBlock' || n.type.name === 'sdtInline' || (n.type.name === 'table_row' && Array.isArray(n.attrs['sdt']) && (n.attrs['sdt'] as unknown[]).length > 0) || (n.type.name === 'table_cell' && n.attrs['sdt'] !== null && n.attrs['sdt'] !== undefined);
        if (tr.getMeta(TEMPLATE_CHANGED)) return verdictDecorations(state.doc, controls, cache);
        if (!tr.docChanged) return set;
        return touches(tr, inControl) ? verdictDecorations(state.doc, controls, cache) : set.map(tr.mapping, tr.doc);
      },
    },
    props: {
      decorations: (state) => verdictsKey.getState(state) ?? DecorationSet.empty,
    },
  });
}
