// SPDX-License-Identifier: GPL-3.0-or-later
//
// A story on screen (ED-002 section 7.1; ED-003 section 6.4): a footnote, a comment or a header
// projected with the body's schema, rendered with the body's node views, mark views, script-font
// decorations and generated stylesheet, so it looks like the body it belongs to. `StoryView` is
// the read-only form over a plain document (comments until E4, the preview); `EditableStoryView`
// is an `EditorSurface` over the story's Yjs fragment (a note, a header, a footer), with the
// body's keymaps, tables, labels, tab stops and links, and undo through the session's manager
// rather than y-prosemirror's undo plugin (UPSTREAM-4). The stylesheet is the body view's: both
// expect a `DocxEditorView` over the same session in the same document.
import { EditorState, type Plugin, type Transaction } from 'prosemirror-state';
import { EditorView, type NodeViewConstructor, type MarkViewConstructor } from 'prosemirror-view';
import { gapCursor } from 'prosemirror-gapcursor';
import { ySyncPlugin, ySyncPluginKey, initProseMirrorDoc } from 'y-prosemirror';
import type { Node as PMNode } from 'prosemirror-model';
import { type EditorDocument, type Story, type LiveLabel, sessionPlugin, storyPlugin } from '@docx4j/editor-model';
import { decorationsPlugin } from './plugins/decorations.js';
import { invisibleAtomsPlugin } from './plugins/invisibleAtoms.js';
import { surfaceKeymaps } from './plugins/keymap.js';
import { labelsPlugin, fixedLabelsPlugin } from './plugins/labels.js';
import { proofingPlugin } from './proofing.js';
import { tabStopsPlugin } from './plugins/tabstops.js';
import { linksPlugin } from './plugins/links.js';
import { tablesPlugins } from './plugins/tables.js';
import { columnSelectPlugin } from './plugins/columnselect.js';
import { tableConditionsPlugin, tableConditionsKey } from './plugins/tableconditions.js';
import { controlsPlugin } from './plugins/controls.js';
import { textInputPlugin } from './plugins/textinput.js';
import { pastePlugin, pasteOptionsOf } from './paste.js';
import { findPlugin } from './plugins/find.js';
import { nodeViews, markViews, type NodeViewOptions } from './nodeviews/index.js';
import { SurfaceBase, type SurfaceOptions, type StylesheetOwner } from './surface.js';

export interface StoryViewOptions extends NodeViewOptions {
  /** An accessible name for the region (a note's number, a comment's author). */
  label?: string;
  /** Extra plugins (the preview's verdict badges). */
  plugins?: Plugin[];
  /** The story's list labels, computed outside the view (the preview's, from its clone's numbering); none otherwise. */
  labels?: Map<PMNode, LiveLabel>;
}

export class StoryView {
  readonly view: EditorView;
  private readonly revoke: (() => void)[] = [];

  constructor(readonly container: HTMLElement, readonly session: EditorDocument, story: Story, options: StoryViewOptions = {}) {
    const state = EditorState.create({
      doc: story.doc,
      plugins: [sessionPlugin(session), ...(options.labels ? [fixedLabelsPlugin(options.labels)] : []), decorationsPlugin(session, { story: true, ...(options.fields ? { fields: options.fields } : {}), ...(options.ids ? { ids: options.ids } : {}) }), ...(options.plugins ?? [])],
    });
    container.classList.add('docx-editor', 'story-view');
    this.view = new EditorView(container, {
      state,
      editable: () => false,
      nodeViews: nodeViews(session, (fn) => this.revoke.push(fn), options) as Record<string, NodeViewConstructor>,
      markViews: markViews(session) as Record<string, MarkViewConstructor>,
      // Focusable though not editable: a click or the Tab key lands in it, so the status line can say why typing does nothing.
      attributes: { role: 'region', 'aria-readonly': 'true', tabindex: '0', dir: 'ltr', ...(options.label ? { 'aria-label': options.label } : {}) },
    });
  }

  get state(): EditorState {
    return this.view.state;
  }

  /** Show another story in place (a re-read after an edit elsewhere). */
  update(story: Story): void {
    this.view.updateState(EditorState.create({ doc: story.doc, plugins: this.view.state.plugins }));
  }

  destroy(): void {
    this.view.destroy();
    for (const fn of this.revoke) fn();
  }
}

export interface EditableStoryViewOptions extends SurfaceOptions {
  /** The surface's id in the registry: the story's key, with the mount when the same story is shown more than once. */
  id?: string;
  /** The part whose relationships resolve the story's images (a header's own). */
  part?: NodeViewOptions['part'];
  /** Field results known for the story (`PAGE`, `NUMPAGES` at a boundary). */
  fields?: Record<string, string>;
  /** An accessible name (the note's number, the header's page). */
  label?: string;
  /** The body view, whose stylesheet the story shares. */
  stylesheet: StylesheetOwner;
}

/**
 * An editable story: the same commands, keys and plugins as the body over the story's fragment
 * (`Stories.fragmentOf`, made on the first binding), so typing in a footnote or a header goes to
 * the store and the session's undo manager as typing in the body does. Several views may bind
 * one fragment (a header at every boundary that shows it) and converge through Yjs.
 */
export class EditableStoryView extends SurfaceBase {
  private readonly stylesheet: StylesheetOwner;

  constructor(container: HTMLElement, session: EditorDocument, story: Story, options: EditableStoryViewOptions) {
    super(options.id ?? story.key, container, session, story, options);
    this.stylesheet = options.stylesheet;
    const fragment = session.stories.fragmentOf(story);
    const { doc, mapping } = initProseMirrorDoc(fragment, session.schema);
    session.undoManager.addTrackedOrigin(ySyncPluginKey);
    // The keymaps, and the test of the keys they bind that the extension's key gate asks (ED-014 section 4.4).
    const keys = surfaceKeymaps(options.shortcuts);
    const state = EditorState.create({
      doc,
      plugins: [
        ySyncPlugin(fragment, { mapping }),
        // UPSTREAM-4: no yUndoPlugin here; `history.undo` and `history.redo` reach the session's manager directly.
        sessionPlugin(session),
        storyPlugin(story),
        proofingPlugin(session),
        invisibleAtomsPlugin(),
        ...keys.plugins,
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
        decorationsPlugin(session, { story: true, ...(options.fields ? { fields: options.fields } : {}) }),
        controlsPlugin(session, {}),
        textInputPlugin(),
        pastePlugin(session, pasteOptionsOf(options)),
        findPlugin(session),
        ...(options.plugins ?? []),
      ],
    });
    container.classList.add('docx-editor', 'story-view', 'editable');
    this.mount(state, {
      nodeViews: nodeViews(session, (fn) => this.revoke.push(fn), { ...(options.part ? { part: options.part } : {}), ...(options.fields ? { fields: options.fields } : {}) }) as Record<string, NodeViewConstructor>,
      markViews: markViews(session) as Record<string, MarkViewConstructor>,
      attributes: { role: 'textbox', 'aria-multiline': 'true', dir: 'ltr', ...(options.label ? { 'aria-label': options.label } : {}) },
    }, keys.binds);
  }

  protected afterDocChanged(tr: Transaction, next: EditorState): void {
    this.stylesheet.refreshIfIntroduced(tr, next.doc, tableConditionsKey.getState(next)?.changedKeys);
  }

  refreshCss(): void {
    this.stylesheet.refreshCss();
  }
}
