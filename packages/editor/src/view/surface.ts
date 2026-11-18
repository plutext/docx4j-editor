// SPDX-License-Identifier: GPL-3.0-or-later
//
// What every editing surface shares (ED-003 section 6.4): the body view and the editable story
// views (a footnote, an endnote, a header, a footer) are one `EditorSurface`, so the toolbar,
// the palette, the dialogs and the panels act on whichever one last took focus, the "active
// surface" the application keeps. The base class holds the command bridge, the transaction
// dispatch that carries a command's control-map writes into the same Yjs transaction, the
// listeners, and the focus report; the subclasses build their plugin stacks.
import type { PasteOptions } from './paste.js';
import { EditorState, TextSelection, type Plugin, type Transaction } from 'prosemirror-state';
import type { StepMap } from 'prosemirror-transform';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorView, type DirectEditorProps, type NodeViewConstructor } from 'prosemirror-view';
import { scrollSelectionWhileUnfocused } from './scroll.js';
import { installSelectionBar } from './selectionbar.js';
import { refreshXhtmlControls } from './xhtmlbind.js';
import { refreshPictureControls } from './picturebind.js';
import { ySyncPluginKey, type ProsemirrorBinding } from 'y-prosemirror';
import { presencePlugin } from './presence.js';
import { markSelectionPlugin } from './plugins/markselection.js';
import { optionalBreaksPlugin } from './plugins/optionalbreaks.js';
import { verticalMotionPlugin } from './plugins/verticalmotion.js';
import { narrowRemoteReplaces } from './narrow.js';
import { syncFormattingOnWrite } from './textformat.js';
import type { Element } from '@docx4j/core-ts';
import { LANGUAGE_CHANGED } from './plugins/decorations.js';
import { onLanguageChange } from '../app/i18n.js';
import { type EditorDocument, type Story, commandNamed, asCommand, taggedDispatch, effectiveFormatting, storeWritesOf, writeStoreWrite, resolutionWriteOf, writeResolution, recordTransaction, markComposed, UNTRACKED_META, syncBoundAfter, controlLockRefusal, applyControlRules, type LockRefusal, type SelectionFormatting, type StoreWrite, type ResolutionWrite } from '@docx4j/editor-model';

export type Listener = (state: EditorState, transactions: readonly Transaction[]) => void;

export interface SurfaceOptions {
  /** Extra plugins after the surface's own (a test, a panel). */
  plugins?: Plugin[];
  /**
   * Application shortcuts (the command palette) bound inside ProseMirror's keymap rather than on
   * `window`, so they run after ProseMirror has read a pending DOM selection change: a click
   * followed at once by a shortcut then acts on the clicked position, not the previous one.
   */
  shortcuts?: Record<string, () => void>;
  /** The surface took focus: the application makes it the active one. */
  onActivate?: (surface: EditorSurface) => void;
  /** The surface is going away: the application falls back to the body when it was the active one. */
  onDestroy?: (surface: EditorSurface) => void;
  /** After a paste from Word or another application (ED-003 section 8.3): the report and its status line. */
  onPaste?: PasteOptions['onPaste'];
  /** A plain paste that reads as Markdown, and the end of its offer (ED-005 section 12.66). */
  onMarkdownOffer?: PasteOptions['onMarkdownOffer'];
  /**
   * With tracking on, an edit that ran untracked (ED-005 section 5.4): the command's name, or
   * `story` for an edit in a note, a header or a footer. The application says so in the status line.
   */
  onUntracked?: (what: string) => void;
  /** An edit a content control's lock refused (ED-005 section 12.60): the application says why in the status line. */
  onRefused?: (reason: LockRefusal) => void;
  /** A navigation into content the markup view hid switched the body to All markup (ED-005 section 6.4): the status line says so. */
  onRevealed?: () => void;
}

/** The document stylesheet's owner (the body view): a story with a style the stylesheet lacks asks for a refresh. */
export interface StylesheetOwner {
  refreshCss(): void;
  /** Regenerate when the transaction put a style on the page the stylesheet has no rule for, or a class of table conditions (`tableKeys`, ED-005 section 12.65). */
  refreshIfIntroduced(tr: Transaction, doc: PMNode, tableKeys?: ReadonlySet<string>): void;
}

export interface EditorSurface {
  /** `body`, or the story's key with the mount (`footnote:3`, `header:/word/header1.xml@2`). */
  readonly id: string;
  readonly container: HTMLElement;
  readonly session: EditorDocument;
  readonly view: EditorView;
  readonly state: EditorState;
  /** The story this surface edits; undefined for the body. */
  readonly story: Story | undefined;
  /**
   * Runs a command by name over this surface: an undo step of its own, unless `joinUndo`, which leaves Yjs's
   * capture window open so the command joins the step before it (a fit that follows an import, ED-007
   * section 8.1 item 14).
   */
  run(name: string, args?: unknown, options?: { joinUndo?: boolean }): boolean;
  runAsync(name: string, args?: unknown): Promise<boolean>;
  formatting(): SelectionFormatting;
  subscribe(fn: Listener): () => void;
  focus(): void;
  hasFocus(): boolean;
  caretAt(pos: number, to?: number): void;
  /**
   * Before a navigation to a range: when the markup view hides it, switch to All markup so that
   * what is selected can be seen (ED-005 section 6.4). The body's; a story has no markup view.
   */
  reveal(from: number, to?: number): boolean;
  /** The element one block of this surface stands for now (reveal codes). Pure. */
  exportBlock(node: PMNode): Element;
  /** Regenerate the document stylesheet (a section or style change, a style coming into use). */
  refreshCss(): void;
  destroy(): void;
}

/** Every live surface by id: `window.docx4jSurfaces` for the browser tests, the application's fallback search. */
export const surfaces = new Map<string, EditorSurface>();

export abstract class SurfaceBase implements EditorSurface {
  view!: EditorView;
  protected readonly listeners = new Set<Listener>();
  protected readonly revoke: (() => void)[] = [];
  protected destroyed = false;

  protected constructor(readonly id: string, readonly container: HTMLElement, readonly session: EditorDocument, readonly story: Story | undefined, protected readonly options: SurfaceOptions) {}

  /** Make the view over the state with the shared dispatch, report focus, and register. */
  protected mount(state: EditorState, props: Omit<DirectEditorProps, 'state' | 'dispatchTransaction'>): void {
    // Presence (ED-004 section 5) from the start: added before the view exists, so joining and
    // leaving a room never change the plugin list (presence.ts says why that matters).
    const withPresence = state.reconfigure({ plugins: [...state.plugins, presencePlugin(this.session), markSelectionPlugin(), optionalBreaksPlugin(), verticalMotionPlugin()] });
    this.view = new EditorView(this.container, {
      ...props,
      state: withPresence,
      dispatchTransaction: (incoming) => {
        // A transaction for a surface already destroyed (a timer set before it went: the presence
        // plugin's redraw, a measuring pass) is dropped.
        if (this.destroyed) return;
        let tr = incoming;
        let untracked: string | undefined;
        if (incoming.getMeta('composition') !== undefined || this.view.composing) {
          // An IME composition (ED-005 section 5.4): marking its text while it is composed breaks
          // it in Chrome, so its transactions go as they are and it is recorded when it is over.
          if (incoming.docChanged) {
            this.composition ??= { start: this.view.state.doc, maps: [] };
            this.composition.maps.push(...incoming.mapping.maps);
            this.scheduleComposed();
          }
        } else {
          // A content control's lock refuses the edit whole, before the tracker; what was typed into the page is drawn
          // back. Its placeholder and its remove-when-edited act after the tracker (ED-005 section 12.60).
          const refused = controlLockRefusal(incoming, this.view.state);
          if (refused) {
            this.view.updateState(this.view.state);
            this.options.onRefused?.(refused);
            return;
          }
          const steps = incoming.steps.length;
          const recorded = recordTransaction(incoming, this.view.state);
          tr = applyControlRules(recorded.tr, this.view.state, steps);
          untracked = recorded.untracked;
          // An edit before a composition is recorded (within 50 ms of its end) joins its maps; being recorded itself, it compares equal.
          if (this.composition && tr.docChanged) this.composition.maps.push(...tr.mapping.maps);
        }
        const writes = storeWritesOf(tr);
        const resolution = resolutionWriteOf(tr);
        const next = this.view.state.apply(tr);
        if (writes.length > 0 || resolution) this.updateWithControlWrites(next, writes, resolution); else this.view.updateState(next);
        if (tr.scrolledIntoView) scrollSelectionWhileUnfocused(this.view);
        if (tr.docChanged) this.afterDocChanged(tr, next);
        // Word's two-way mapping (ED-005 section 12.55): an edit, or an undo, inside a bound control in the body writes its
        // node, and the controls bound to that node follow once this dispatch is over. A peer's change is left: its page wrote.
        if (tr.docChanged && this.story === undefined) {
          const sync = tr.getMeta(ySyncPluginKey) as { isChangeOrigin?: boolean; isUndoRedoOperation?: boolean } | undefined;
          if (!sync?.isChangeOrigin || sync.isUndoRedoOperation) {
            syncBoundAfter(this.session, tr, () => this.view.state, (fn) => queueMicrotask(() => {
              if (this.destroyed) return;
              fn();
              // The XHTML-bound controls follow the nodes a text control wrote (ED-005 section 12.75); unchanged values cost nothing.
              void refreshXhtmlControls(this.view, this.session);
              void refreshPictureControls(this.view, this.session);
            }));
          }
        }
        for (const fn of this.listeners) fn(next, [tr]);
        if (untracked) this.options.onUntracked?.(untracked);
      },
      handleDOMEvents: {
        ...props.handleDOMEvents,
        // Typing over a selection with an IME: the selection is deleted, tracked, before the
        // composition starts, so what it replaces is kept and marked (section 5.4).
        compositionstart: (view, event) => {
          if (!view.state.selection.empty && this.session.tracking() && !this.story) view.dispatch(view.state.tr.deleteSelection());
          return props.handleDOMEvents?.compositionstart?.(view, event) ?? false;
        },
      },
    });
    // UPSTREAM-9: a colleague's keystroke becomes a step over the blocks it changed, not the whole document (narrow.ts).
    narrowRemoteReplaces(this.view);
    // UPSTREAM-15: formatting the sync plugin's diff misjudges (bold taken off as a formatChange comes) written too (textformat.ts).
    syncFormattingOnWrite(this.view);
    const onFocus = (): void => { if (!this.destroyed) this.options.onActivate?.(this); };
    this.view.dom.addEventListener('focus', onFocus);
    this.revoke.push(() => this.view.dom.removeEventListener('focus', onFocus));
    // Word's selection bar: a click in the page's left margin selects the line beside it (selectionbar.ts).
    this.revoke.push(installSelectionBar(this.view, this.container));
    surfaces.set(this.id, this);
    this.revoke.push(() => { if (surfaces.get(this.id) === this) surfaces.delete(this.id); });
    this.revoke.push(onLanguageChange(() => { if (!this.destroyed) this.relocalise(); }));
  }

  /**
   * The application's language changed (ED-003 section 12.2): the decorations are computed again
   * (the page labels, the section breaks) and every node view made again (the breaks, the
   * tooltips). New node view functions are what makes ProseMirror redraw: it compares them by
   * identity.
   */
  relocalise(): void {
    this.view.dispatch(this.view.state.tr.setMeta(LANGUAGE_CHANGED, true));
    const views = this.view.props.nodeViews;
    if (views) this.view.setProps({ nodeViews: Object.fromEntries(Object.entries(views).map(([name, make]) => [name, ((...args: Parameters<NodeViewConstructor>) => make(...args)) as NodeViewConstructor])) });
  }

  /** After a transaction that changed the document: the body checks the stylesheet, a story asks the body to. */
  protected abstract afterDocChanged(tr: Transaction, next: EditorState): void;

  /** An IME composition not yet recorded: the document when it began, and the step maps since. */
  private composition: { start: PMNode; maps: StepMap[] } | undefined;
  private composedTimer: ReturnType<typeof setTimeout> | undefined;

  /** After ProseMirror has ended the composition (20 ms after compositionend), its text is recorded as a transaction would have been. */
  private scheduleComposed(): void {
    if (this.composedTimer !== undefined) clearTimeout(this.composedTimer);
    this.composedTimer = setTimeout(() => {
      this.composedTimer = undefined;
      if (this.destroyed || !this.composition) return;
      if (this.view.composing) { this.scheduleComposed(); return; }
      const { start, maps } = this.composition;
      this.composition = undefined;
      const tr = markComposed(this.view.state.tr, start, maps, this.story ? undefined : this.session.tracking());
      if (tr.docChanged) this.view.dispatch(tr.setMeta(UNTRACKED_META, true));
    }, 50);
  }

  /**
   * A command's control-map writes (ED-003 section 4.2: a row or cell group's properties) go
   * into the same Yjs transaction as the sync plugin's update of the fragment, so a wrapped group
   * and its properties are one update and one undo step, as the headless runner makes them.
   * UPSTREAM-3: y-prosemirror offers no hook into that transaction, so this holds the binding's
   * mutex, opens the transaction under the sync plugin's origin, updates the view (the plugin's
   * own update is muted by the mutex) and syncs the document itself, which is what the plugin
   * would have done; the fragment's observer then fires inside the mutex, as it does for every
   * local change, and is ignored.
   */
  private updateWithControlWrites(next: EditorState, writes: readonly StoreWrite[], resolution?: ResolutionWrite): void {
    const store = this.session.store;
    const binding = (ySyncPluginKey.getState(next) as { binding?: ProsemirrorBinding } | undefined)?.binding;
    if (!binding) {
      // No binding yet (the sync plugin's view is not initialised): the two transactions coalesce in the undo manager's capture window.
      if (resolution) store.ydoc.transact(() => writeResolution(store, resolution), ySyncPluginKey);
      this.view.updateState(next);
      store.ydoc.transact(() => { for (const w of writes) writeStoreWrite(store, w); }, ySyncPluginKey);
      return;
    }
    binding.mux(() => {
      store.ydoc.transact(() => {
        // A resolution's record first, so that its state vector is what the resolver saw (ED-005 section 12.15).
        if (resolution) writeResolution(store, resolution);
        for (const w of writes) writeStoreWrite(store, w);
        this.view.updateState(next);
        binding._prosemirrorChanged(next.doc);
      }, ySyncPluginKey);
    });
  }

  get state(): EditorState {
    return this.view.state;
  }

  /**
   * Run a registered command through the view. A command with a `prepare` half (a list started
   * or restarted: a part write first) runs when that half resolves; true then means it was
   * started, and `runAsync` answers whether it applied.
   */
  run(name: string, args?: unknown, options: { joinUndo?: boolean } = {}): boolean {
    const spec = commandNamed(name);
    if (!spec) throw new Error(`no command named ${name}`);
    if (!options.joinUndo) this.ownUndoStep(spec);
    return asCommand(spec, args)(this.view.state, this.view.dispatch, this.view);
  }

  /**
   * A command given from the surface (a toolbar button, the palette, a panel's button, the find bar)
   * is an undo step of its own, as the headless runner makes one (`runCommand`): otherwise one given
   * within Yjs's capture window (500 ms) of the last edit joined that edit's step, and two Accepts in
   * quick succession were undone together (ED-005 section 12.41). Keys keep Yjs's grouping, as before:
   * Enter, Tab and Backspace typed among other keys join the typing's step (the third review point,
   * section 12.47).
   */
  private ownUndoStep(spec: { typing?: boolean }): void {
    if (!spec.typing) this.session.undoManager.stopCapturing();
  }

  /** `run`, awaiting a command's `prepare` half; the answer is whether the command applied. */
  async runAsync(name: string, args?: unknown): Promise<boolean> {
    const spec = commandNamed(name);
    if (!spec) throw new Error(`no command named ${name}`);
    const prepared = spec.prepare ? await spec.prepare(this.view.state, args) : args;
    if (this.destroyed) return false;
    this.ownUndoStep(spec);
    return spec.run(prepared)(this.view.state, taggedDispatch(spec.name, this.view.dispatch.bind(this.view)), this.view);
  }

  /** The selection's effective formatting, for the toolbar. */
  formatting(): SelectionFormatting {
    return effectiveFormatting(this.view.state);
  }

  subscribe(fn: Listener): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  focus(): void {
    this.view.focus();
  }

  hasFocus(): boolean {
    return this.view.hasFocus();
  }

  /** Put the caret at a document position, or select a range (panels that select, tests). */
  caretAt(pos: number, to: number = pos): void {
    const { doc } = this.view.state;
    const clamp = (p: number): number => Math.max(0, Math.min(p, doc.content.size));
    this.view.dispatch(this.view.state.tr.setSelection(TextSelection.create(doc, clamp(pos), clamp(to))).scrollIntoView());
    this.view.focus();
  }

  reveal(_from: number, _to?: number): boolean {
    return false;
  }

  exportBlock(node: PMNode): Element {
    return this.story ? this.session.stories.exportBlock(this.story, node) : this.session.exportBlock(node);
  }

  abstract refreshCss(): void;

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.composedTimer !== undefined) clearTimeout(this.composedTimer);
    this.options.onDestroy?.(this);
    this.view.destroy();
    for (const fn of this.revoke) fn();
    this.listeners.clear();
  }
}
