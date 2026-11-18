// SPDX-License-Identifier: GPL-3.0-or-later
//
// Perspectives as plugins (ED-001 section 3.6 and decision 18; ED-002 section 5.4; ED-003
// section 4.9): one document, one model, one command set; a perspective is an arrangement of the
// side panel and a toolbar, and it hides, never removes: every command stays in the palette. A
// perspective contributes the side panel's tabs (id, label, what to render), which toolbar it
// wants, its panel preset, and whether a document suggests it. The interface is internal, with
// no app store and no stability promise to third parties (decision 18). The presets register
// here from `presets.tsx`: Edit, Review (E4.d), Template and Developer. The side panel is
// one column with tabs, as a browser's developer tools are (Jason, 2026-09-24); a perspective
// says whether it is open and which tab shows, and a user's own choice overrides it.
import type { ReactNode } from 'react';
import type { EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import type { EditorSurface } from '../view/surface.js';

export type PerspectiveId = string;

export type TabId = string;

/** A script handed to the console (Console.tsx): `seq` makes the same text twice two hand-overs. */
export interface ConsoleScript {
  text: string;
  mode: 'replace' | 'insert';
  seq: number;
}

/** What a tab renders with: the view and the session, and the shell's callbacks. */
export interface TabContext {
  /** The body view. */
  editor: DocxEditorView;
  /** The surface that last took focus (the body, or a note or header being edited; ED-003 section 6.4): reveal codes and the Format and Styles tabs follow it. */
  active: EditorSurface;
  session: EditorDocument;
  perspective: { id: PerspectiveId; panel: PanelState };
  /** A script handed to the console: the reveal-codes pane's API script replaces what is there, the palette's command line is inserted at the caret. */
  consoleScript: ConsoleScript | null;
  onConsoleScript: (script: string) => void;
  /** Save a console script as a macro (ED-003 section 7.4): the shell asks the name and the home. */
  onSaveMacro: (script: string) => void;
  /** The status line of the application bar. */
  onStatus: (text: string) => void;
  /** Bring another tab of the panel to the front. */
  onTab: (tab: TabId) => void;
  /** Turn the document's tracking on through the switch's path, a name asked for first when none is set (ED-005 section 6.7's suggestion). */
  onTrackChanges?: () => void;
}

export interface TabSpec {
  id: TabId;
  label: string;
  render: (ctx: TabContext) => ReactNode;
}

export interface PanelState {
  /** Whether the side panel shows. */
  open: boolean;
  /** The tab in front. */
  tab: TabId;
}

/** The formatting toolbar a perspective shows: everything, the paragraph and character basics, or none. */
export type ToolbarPreset = 'full' | 'basic' | 'none';

export interface Perspective {
  id: PerspectiveId;
  label: string;
  /** The side panel as the perspective presents it; a user's own arrangement overrides it. */
  panel: PanelState;
  toolbar: ToolbarPreset;
  /** The content control group in the toolbar (insert kinds, wrap rows, delete, unwrap, bind). */
  controls?: boolean;
  /** Draw content controls with their tags and coloured gutters on the page (editor.css `.show-control-tags`). */
  controlTags?: boolean;
  /** Draw the formatting marks (a paragraph mark's pilcrow, a cell's end, a tab's arrow) on the page (editor.css `.show-marks`, ED-005 section 12.80). */
  formattingMarks?: boolean;
  /** The side panel's tabs; none means the panel has nothing to show and stays closed. */
  tabs: readonly TabSpec[];
  /** Whether a document suggests this perspective (the banner offers it; the user's choice persists). */
  suggest?: (session: EditorDocument) => boolean;
}

const registry: Perspective[] = [];

export function registerPerspective(perspective: Perspective): Perspective {
  if (registry.some((p) => p.id === perspective.id)) throw new Error(`perspective ${perspective.id} registered twice`);
  registry.push(perspective);
  return perspective;
}

/** Every registered perspective, in registration order (the application bar's order). */
export function perspectives(): readonly Perspective[] {
  return registry;
}

export function perspectiveNamed(id: PerspectiveId): Perspective | undefined {
  return registry.find((p) => p.id === id);
}

const KEY = 'docx4j-editor.perspective';
const PANEL_KEY = 'docx4j-editor.panel.';
const WIDTH_KEY = 'docx4j-editor.panelWidth';

export function loadPerspective(): PerspectiveId {
  try {
    const v = localStorage.getItem(KEY);
    if (v && perspectiveNamed(v)) return v;
  } catch { /* storage may be unavailable */ }
  return registry[0]?.id ?? 'edit';
}

export function savePerspective(id: PerspectiveId): void {
  try { localStorage.setItem(KEY, id); } catch { /* ignore */ }
}

function presetOf(id: PerspectiveId): PanelState {
  const p = perspectiveNamed(id);
  return p ? { ...p.panel } : { open: false, tab: '' };
}

/** The side panel as a perspective shows it: the user's arrangement when they changed it, else the preset; a tab the perspective lacks falls back to the preset's. */
export function loadPanel(id: PerspectiveId): PanelState {
  const preset = presetOf(id);
  const tabs = perspectiveNamed(id)?.tabs ?? [];
  const isTab = (v: unknown): v is TabId => typeof v === 'string' && tabs.some((t) => t.id === v);
  try {
    const v = localStorage.getItem(PANEL_KEY + id);
    if (v) {
      const parsed = JSON.parse(v) as Partial<PanelState>;
      return { open: typeof parsed.open === 'boolean' ? parsed.open : preset.open, tab: isTab(parsed.tab) ? parsed.tab : preset.tab };
    }
  } catch { /* ignore */ }
  return preset;
}

export function savePanel(id: PerspectiveId, state: PanelState): void {
  try { localStorage.setItem(PANEL_KEY + id, JSON.stringify(state)); } catch { /* ignore */ }
}

export function resetPanel(id: PerspectiveId): PanelState {
  try { localStorage.removeItem(PANEL_KEY + id); } catch { /* ignore */ }
  return presetOf(id);
}

export const PANEL_MIN_WIDTH = 280;
export const PANEL_DEFAULT_WIDTH = 460;

/** The side panel's width in pixels, one for every perspective: a user's own or the default. */
export function loadPanelWidth(): number {
  try {
    const v = Number(localStorage.getItem(WIDTH_KEY));
    if (Number.isFinite(v) && v >= PANEL_MIN_WIDTH) return v;
  } catch { /* ignore */ }
  return PANEL_DEFAULT_WIDTH;
}

export function savePanelWidth(width: number): void {
  try { localStorage.setItem(WIDTH_KEY, String(Math.round(width))); } catch { /* ignore */ }
}
