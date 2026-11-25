// SPDX-License-Identifier: GPL-3.0-or-later
//
// The editing surface (ED-002 section 5.2; ED-001 section 2's view layer): one ProseMirror
// `EditorView` over a session's Yjs store, bound to the command registry. Framework-free, and
// internal to the application in E1 (ED-002 section 8 item 6): not exported, no custom element.
// Since E2.d (ED-003 section 6.4) it is the body's `EditorSurface`, the owner of the document
// stylesheet, and the host of the editable header and footer views at the page boundaries.
import { proofingPlugin } from './proofing.js';
import { EditorState, Plugin, type Transaction } from 'prosemirror-state';
import { AttrStep } from 'prosemirror-transform';
import type { Node as PMNode } from 'prosemirror-model';
import type { NodeViewConstructor, MarkViewConstructor } from 'prosemirror-view';
import { keymap } from 'prosemirror-keymap';
import { gapCursor } from 'prosemirror-gapcursor';
import { ySyncPlugin, yUndoPlugin, ySyncPluginKey, initProseMirrorDoc } from 'y-prosemirror';
import type { Transaction as YTransaction } from 'yjs';
import { type EditorDocument, sessionPlugin, FINAL_SECT_PR } from '@docx4j/editor-model';
import { wordKeymap, editorBaseKeymap } from './plugins/keymap.js';
import { finalSectPrPlugin } from './plugins/finalSectPr.js';
import { invisibleAtomsPlugin } from './plugins/invisibleAtoms.js';
import { decorationsPlugin, type BoundaryMount, type PageSide } from './plugins/decorations.js';
import { pastePlugin, pasteOptionsOf } from './paste.js';
import { labelsPlugin, PARTS_CHANGED } from './plugins/labels.js';
import { tabStopsPlugin, TABS_REMEASURE } from './plugins/tabstops.js';
import { linksPlugin } from './plugins/links.js';
import { tablesPlugins } from './plugins/tables.js';
import { columnSelectPlugin } from './plugins/columnselect.js';
import { tableConditionsPlugin, tableConditionsKey, TABLE_STYLES_CHANGED } from './plugins/tableconditions.js';
import { controlsPlugin, CONTROLS_CHANGED, type ControlsPluginOptions } from './plugins/controls.js';
import { textInputPlugin } from './plugins/textinput.js';
import { verdictsPlugin, TEMPLATE_CHANGED } from './plugins/verdicts.js';
import { findPlugin } from './plugins/find.js';
import { markupPlugin, markupKey, setMarkupView, revealsHidden, type MarkupView } from './markup.js';
import { PrintPages } from './print.js';
import { AuthorStyles } from './authors.js';
import { currentChangePlugin } from './plugins/review.js';
import { nodeViews, markViews } from './nodeviews/index.js';
import { documentCss, stylesInUse, tableConditionsIn } from './css.js';
import { EditableStoryView } from './StoryView.js';
import { SurfaceBase, surfaces, type SurfaceOptions, type StylesheetOwner, type EditorSurface } from './surface.js';
import { t } from '../app/i18n.js';

export type { Listener, EditorSurface } from './surface.js';

export interface DocxEditorViewOptions extends SurfaceOptions {
  /** Headers and footers at the page boundaries (ED-002 section 7.4); on unless a test says otherwise. */
  headersAndFooters?: boolean;
  /** What happens after a data node is dropped on a control (ED-003 section 4.5's bind by drag). */
  onDrop?: ControlsPluginOptions['onDrop'];
}

/**
 * The footer of the page that ends at a boundary and the header of the one that starts, as Word
 * picks them (`Stories.headerFooterFor`), each an editable story over the part's fragment with
 * the page's `PAGE` and `NUMPAGES` (ED-003 section 6.4: one fragment per part, a binding per
 * boundary). The parts are read on first use; a widget torn down before they arrive mounts nothing.
 */
function boundaryStories(body: DocxEditorView, options: DocxEditorViewOptions): BoundaryMount {
  const session = body.session;
  return (slots, info) => {
    const views: EditableStoryView[] = [];
    let cancelled = false;
    void (async () => {
      const total = String(info.total());
      const mount = async (slot: HTMLElement, side: PageSide, kind: 'header' | 'footer'): Promise<void> => {
        // The variant follows the number the page shows (its parity), not the physical count (ED-003 section 6.4).
        const stories = await session.stories.headerFooterFor(side.section, side.numeric, side.firstOfSection, info.sections);
        const story = kind === 'header' ? stories.header : stories.footer;
        if (cancelled || !story) return;
        const el = document.createElement('div');
        el.className = `boundary-story ${kind} variant-${stories.variant}`;
        el.dataset['rid'] = story.rId;
        el.dataset['page'] = String(side.page);
        slot.appendChild(el);
        views.push(new EditableStoryView(el, session, story, {
          id: `${story.key}@${side.page}`,
          part: story.part,
          fields: { PAGE: side.number, NUMPAGES: total },
          label: t(kind === 'header' ? 'view.headerOfPage' : 'view.footerOfPage', { page: side.number }),
          stylesheet: body,
          ...(options.shortcuts ? { shortcuts: options.shortcuts } : {}),
          ...(options.onActivate ? { onActivate: options.onActivate } : {}),
          ...(options.onDestroy ? { onDestroy: options.onDestroy } : {}),
          ...(options.onPaste ? { onPaste: options.onPaste } : {}),
          ...(options.onMarkdownOffer ? { onMarkdownOffer: options.onMarkdownOffer } : {}),
          ...(options.onUntracked ? { onUntracked: options.onUntracked } : {}),
          ...(options.onRefused ? { onRefused: options.onRefused } : {}),
        }));
      };
      if (info.above) await mount(slots.above, info.above, 'footer');
      if (info.below) await mount(slots.below, info.below, 'header');
    })();
    return () => { cancelled = true; for (const v of views) v.destroy(); views.length = 0; };
  };
}

/**
 * y-prosemirror's undo plugin over the session's manager, with a view whose destroy removes the
 * listeners the plugin added and leaves the manager alone. UPSTREAM-4: the plugin's own destroy
 * calls `undoManager.destroy()`, so remaking the body view for the same session (sharing it
 * before E3.d, the headers-and-footers option) left the session with no undo at all: nothing
 * typed afterwards reached the stack (ED-004 section 10.8 item 1).
 */
function undoPluginKeeping(manager: EditorDocument['undoManager']): Plugin {
  const plugin = yUndoPlugin({ undoManager: manager });
  const makeView = plugin.spec.view!;
  return new Plugin({
    ...plugin.spec,
    view(view) {
      const added: [string, (...args: never[]) => void][] = [];
      const on = manager.on.bind(manager);
      manager.on = ((name: string, fn: (...args: never[]) => void) => { added.push([name, fn]); return on(name as never, fn as never); }) as typeof manager.on;
      try { makeView(view); } finally { manager.on = on as typeof manager.on; }
      return { destroy: () => { for (const [name, fn] of added) manager.off(name as never, fn as never); } };
    },
  });
}

export class DocxEditorView extends SurfaceBase implements StylesheetOwner {
  private readonly styleElement: HTMLStyleElement;
  /** One rule per author the tracked changes name (ED-005 section 6.3). */
  readonly authors: AuthorStyles;

  constructor(container: HTMLElement, session: EditorDocument, options: DocxEditorViewOptions = {}) {
    super('body', container, session, undefined, options);
    const { doc, mapping } = initProseMirrorDoc(session.store.body, session.schema);
    // The root node's attributes are not in the fragment (ED-002 section 10.1 item 2): the final
    // section comes from the package map, here and in the plugin that keeps the two in step.
    const withAttrs = session.schema.topNodeType.create({ finalSectPr: session.store.package.get(FINAL_SECT_PR) ?? null }, doc.content);
    // A keystroke here goes to the store under the sync plugin's origin: the user's history too.
    session.undoManager.addTrackedOrigin(ySyncPluginKey);
    const state = EditorState.create({
      doc: withAttrs,
      plugins: [
        ySyncPlugin(session.store.body, { mapping }),
        // The one view that installs y-prosemirror's undo plugin, wrapped so that its view's
        // destroy leaves the session's manager alive (UPSTREAM-4); the story views bind undo without it.
        undoPluginKeeping(session.undoManager),
        sessionPlugin(session),
        finalSectPrPlugin(session),
        // The spelling option and the document's language, on the surface for the browser's spell checker.
        proofingPlugin(session),
        invisibleAtomsPlugin(),
        // Before the keymaps: Backspace and Delete next to content the view hides (ED-005 section 6.4).
        markupPlugin(),
        keymap(Object.fromEntries(Object.entries(options.shortcuts ?? {}).map(([key, fn]) => [key, () => { fn(); return true; }]))),
        keymap(wordKeymap()),
        keymap(editorBaseKeymap()),
        // Before the table plugins: a click on a table's top border selects the column, never starts a cell drag (ED-005 section 12.107).
        columnSelectPlugin(),
        ...tablesPlugins(),
        // A caret after a table or a block control that ends the body, a cell or a control, where no paragraph follows
        // (Word adds one on open; the editor keeps the file as it is): Enter or typing there makes the paragraph.
        gapCursor(),
        tableConditionsPlugin(session),
        labelsPlugin(session),
        tabStopsPlugin(session),
        linksPlugin(session),
        decorationsPlugin(session, options.headersAndFooters === false ? {} : { boundary: boundaryStories(this, options) }),
        controlsPlugin(session, options.onDrop ? { onDrop: options.onDrop } : {}),
        // Text input narrowed to what changed: a composition inside a marked run keeps its caret (ED-013 section 9).
        textInputPlugin(),
        verdictsPlugin(session, session.controlEntry),
        pastePlugin(session, pasteOptionsOf(options)),
        findPlugin(session),
        currentChangePlugin(),
        ...(options.plugins ?? []),
      ],
    });
    this.styleElement = document.createElement('style');
    this.styleElement.dataset['docx4jEditor'] = 'document';
    this.styleElement.textContent = documentCss(session, withAttrs, this.storyDocs());
    this.covered = stylesInUse(session, withAttrs, this.storyDocs());
    container.ownerDocument.head.appendChild(this.styleElement);
    this.authors = new AuthorStyles(container.ownerDocument);
    container.classList.add('docx-editor', 'body');
    this.mount(state, {
      nodeViews: nodeViews(session, (fn) => this.revoke.push(fn), { authors: this.authors }) as Record<string, NodeViewConstructor>,
      markViews: markViews(session, this.authors) as Record<string, MarkViewConstructor>,
      // A function, so the name follows the language; the document's own direction is its paragraphs' (w:bidi), never the chrome's.
      attributes: () => ({ role: 'textbox', 'aria-multiline': 'true', 'aria-label': t('view.document'), dir: 'ltr' }),
    });
    // Remote changes (a peer, an undo from the store) arrive as transactions through the same path.
    // The controls map changed by something other than this view (an undo, a tree run, a peer):
    // an empty transaction so the current-control marks, the row labels and the panels refresh.
    const onControls = (_events: unknown, txn: YTransaction): void => {
      if (txn.origin !== ySyncPluginKey && !this.destroyed) this.view.dispatch(this.view.state.tr.setMeta(CONTROLS_CHANGED, true));
    };
    session.store.controls.observeDeep(onControls);
    this.revoke.push(() => session.store.controls.unobserveDeep(onControls));
    // The template view changed (its load, a part write, the data part edited): the verdict badges recompute.
    this.revoke.push(session.template.subscribe(() => { if (!this.destroyed) this.view.dispatch(this.view.state.tr.setMeta(TEMPLATE_CHANGED, true)); }));
    // A notes or header part read: the stylesheet takes the styles the stories use (ED-003 section 6.4).
    this.revoke.push(session.stories.subscribe(() => { if (!this.destroyed) this.refreshCss(); }));
    // A part edit that arrived from the room (a peer's style, list or settings): the look follows (ED-005 section 4.5).
    this.revoke.push(session.onPartsReceived(() => this.partsChanged()));
    // Each section on its own page when the document prints (ED-006 section 3.3).
    this.printPages = new PrintPages(session, () => this.view.state.doc, container.ownerDocument);
  }

  /** The print rules for the document as it is now: `beforeprint` writes them, and the Print command before it prints. */
  preparePrint(): void {
    this.printPages.prepare();
  }
  private readonly printPages: PrintPages;

  /**
   * A part the document's look depends on changed (ED-005 section 4.4): an Apply in the parts
   * panel, or a peer's edit arriving. The stylesheet is made again, and every surface of the
   * session walks its labels and computes its decorations again, the document being unchanged.
   */
  partsChanged(): void {
    if (this.destroyed) return;
    this.refreshCss();
    this.printPages?.refreshSettings();
    for (const s of [...surfaces.values()]) if (s.session === this.session) s.view.dispatch(s.state.tr.setMeta(PARTS_CHANGED, true).setMeta(TABLE_STYLES_CHANGED, true));
  }

  protected afterDocChanged(tr: Transaction, next: EditorState): void {
    // A style the stylesheet has no rule for yet (a picked style, a pasted run style), or a class of table conditions: regenerated.
    this.refreshIfIntroduced(tr, next.doc, tableConditionsKey.getState(next)?.changedKeys);
  }

  /** The style ids the stylesheet has rules for: the document's and the read stories' at generation. */
  private covered: { paragraph: Set<string>; character: Set<string> };

  /** The classes of table conditions the stylesheet has rules for (ED-005 section 12.65). */
  private coveredTables = new Set<string>();

  /** The current documents of the stories read so far, whose styles the stylesheet covers too. */
  private storyDocs(): PMNode[] {
    const stories = this.session.stories;
    return stories.loadedStories().map((s) => stories.currentDoc(s));
  }

  /** Regenerate the document stylesheet (after a section or style change, or a style coming into use). */
  refreshCss(): void {
    const doc = this.view.state.doc;
    const extra = this.storyDocs();
    this.styleElement.textContent = documentCss(this.session, doc, extra);
    this.covered = stylesInUse(this.session, doc, extra);
    const tables = tableConditionsIn(this.session, [doc, ...extra]);
    this.coveredTables = new Set([...tables.cells.keys(), ...tables.paragraphs.keys()]);
    // New rules move text, and tabs with it: every tab is measured again after the paint.
    if (!this.destroyed) this.view.dispatch(this.view.state.tr.setMeta(TABS_REMEASURE, true));
  }

  refreshIfIntroduced(tr: Transaction, doc: PMNode, tableKeys?: ReadonlySet<string>): void {
    if (this.introducesStyle(tr, doc) || [...(tableKeys ?? [])].some((k) => !this.coveredTables.has(k))) this.refreshCss();
  }

  /** Whether the transaction put a paragraph or run style on the page that the stylesheet does not cover: only the changed ranges are read. */
  private introducesStyle(tr: Transaction, doc: PMNode): boolean {
    const ranges: [number, number][] = [];
    tr.steps.forEach((step, i) => {
      const rest = tr.mapping.slice(i + 1);
      step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => ranges.push([rest.map(newStart, -1), rest.map(newEnd, 1)]));
      if (step instanceof AttrStep) { const pos = rest.map(step.pos, -1); ranges.push([pos, pos + 1]); }
    });
    for (const [from, to] of ranges) {
      let found = false;
      doc.nodesBetween(Math.max(0, from), Math.min(to, doc.content.size), (node) => {
        if (found) return false;
        if (node.type.name === 'paragraph') {
          const id = node.attrs['pStyle'] as string | null;
          if (id && !this.covered.paragraph.has(id)) found = true;
        }
        for (const m of node.marks) if (m.type.name === 'runStyle' && typeof m.attrs['id'] === 'string' && !this.covered.character.has(m.attrs['id'])) found = true;
        return !found;
      });
      if (found) return true;
    }
    return false;
  }

  /** Which view of the tracked changes the body shows (ED-005 section 6.4): All markup unless chosen. */
  get markupView(): MarkupView {
    return markupKey.getState(this.view.state)?.view ?? 'all';
  }

  set markupView(view: MarkupView) {
    if (this.destroyed || view === this.markupView) return;
    this.view.dispatch(setMarkupView(this.view.state.tr, view).scrollIntoView());
  }

  /**
   * Before a navigation to a range (a review panel's entry, a find hit, an XPath hit, a comment's
   * text): when the current view hides it, switch to All markup so that what is selected can be
   * seen (ED-005 section 6.4). True when it switched, for the caller's status line.
   */
  override reveal(from: number, to: number = from): boolean {
    if (this.destroyed || !revealsHidden(this.view.state, from, to)) return false;
    this.markupView = 'all';
    this.options.onRevealed?.();
    return true;
  }

  override destroy(): void {
    super.destroy();
    this.styleElement.remove();
    this.printPages?.destroy();
    this.authors.destroy();
  }
}

export type { EditorSurface as Surface };
