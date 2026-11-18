// SPDX-License-Identifier: GPL-3.0-or-later
//
// The application shell (ED-002 section 5.4): the Edit and Developer perspectives over one
// session and one view. No application state lives in React: the session and the view are the
// state; React holds what is open on screen (the perspective, the panels, a dialog).
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as Y from 'yjs';
import { EditorDocument, SelfCheckError, documentStyles, documentFonts, preservedCount, APPLICATION_NAME, IMPORT_ORIGIN, ROOM_ORIGIN, AGENTS_ORIGIN, staleBoundControls, applyBoundRefresh, stateFor, compareWithBytes, type StyleEntry, type StaleBoundControls, type TextComparison, type SaveFormat } from '@docx4j/editor-model';
import { refreshXhtmlControls } from '../view/xhtmlbind.js';
import { refreshBoundAtCaret } from '../view/boundrefresh.js';
import { setSpellcheck } from '../view/proofing.js';
import { refreshPictureControls } from '../view/picturebind.js';
import { refreshFragmentControls } from '../view/fragmentbind.js';
import type { IndexeddbPersistence } from 'y-indexeddb';
import { DocxEditorView, type DocxEditorViewOptions } from '../view/DocxEditorView.js';
import { keepPlace, restorePlace, type KeptPlace } from '../view/place.js';
import { surfaces, type EditorSurface } from '../view/surface.js';
import { Toolbar } from './Toolbar.js';
import { Palette } from './Palette.js';
import { SelfCheckDialog } from './SelfCheckDialog.js';
import { PageSizeDialog, TrackingNameDialog, DiscardReloadDialog, DiscardCloseDialog } from './Dialogs.js';
import { editCauseOf, deepStack, followInput, logEditCause, type EditCause } from './editcause.js';
import { TrackingSwitch, TrackingIndicator } from './TrackingSwitch.js';
import { ZoomControl } from './ZoomControl.js';
import { loadZoom, saveZoom, stepZoom, clampZoom } from './zoom.js';
import { PAGE_SIZES, loadPageSize, savePageSize, isPageSize, loadMargins, saveMargins, isMarginPreset, newDocument, templateKind, documentFromTemplate, documentFromMarkdown, type PageSizePaper, type MarginPreset } from './newdoc.js';
import { DiagnosticList } from './panels/DiagnosticList.js';
import { perspectives, perspectiveNamed, loadPerspective, savePerspective, loadPanel, savePanel, resetPanel, loadPanelWidth, savePanelWidth, type PerspectiveId, type PanelState, type TabId, type ConsoleScript } from './perspectives.js';
import './presets.js';
import { DeveloperPanel, Splitter } from './panels/DeveloperPanel.js';
import { FootnotesPanel } from './panels/FootnotesPanel.js';
import { CommentsMargin } from './panels/CommentsMargin.js';
import { reasonOf, hasFileSystemAccess, pickFile, fromFile, handleOfDrop, readHandle, saveAsDownload, saveAs, saveInPlace, download, diagnosisName, isMarkdownName, formatOfName, OPEN_ACCEPT, type OpenedFile, type SourceFile } from './files.js';
import type { EditorView } from 'prosemirror-view';
import { formatAsMarkdown } from '../view/markdown.js';
import { printDocument, prepareNotes, isPrintKey } from './print.js';
import { fitTableColumns, fitTablesIn } from '../view/tablefit.js';
import { isInTable } from 'prosemirror-tables';
import { describePaste, describeDropped, markdownOfferOf } from '../view/paste.js';
import { newSessionId, persistedStore, putSession, deleteSession, setCurrentSessionId, recoverable, holdSession, sessionInUse, type SessionRecord } from './persistence.js';
import { OPTION_COMPANION, companionUrl, effectiveCompanion, probeCompanionInfo, companionRelay, inviteAgent, paginateWith, type CompanionInfo } from './companion.js';
import { RoomConnection, RoomRefusedError, newRoomLink, parseRoomLink, relayAddress, relayHost, relayOption, roomLinkUrl, setRelayOption, type RoomLink, type RoomStatus } from './room.js';
import { ShareDialog, JoinDialog } from './RoomDialogs.js';
import { PeopleList } from './PeopleList.js';
import { localUser, nameOption, setNameOption } from './presence.js';
import { setPresence } from '../view/presence.js';
import { decorationsKey } from '../view/plugins/decorations.js';
import { loadEmbeddedFonts, unloadEmbeddedFonts, describeEmbeddedFonts } from './fonts.js';
import { browserMacros, putBrowserMacro, deleteBrowserMacro, documentMacros, putDocumentMacro, deleteDocumentMacro, documentMacrosEnabled, isTrusted, trust, runMacro, describeRun, type MacroHome } from './macros.js';
import { SaveMacroDialog, MacroSourceDialog } from './MacroDialogs.js';
import type { Macro, LockRefusal } from '@docx4j/editor-model';
import type { PaletteMacro, PaletteAction } from './Palette.js';
import { chooseMarkupView, listedChanges, stepChange } from './review.js';
import { askOption, setAskOption, OPTION_ASK_KEY, OPTION_ASK_BASE_URL } from './ask.js';
import { t, tNodes, useT, useLanguage, directionOf, availableLanguages, languageOption, setLanguage, nativeName, isDraft, matchLanguage, SOURCE_LANGUAGE } from './i18n.js';
import { diagnosticSummary, perspectiveLabel, tabLabel, untrackedStatus, lockRefusedStatus } from './labels.js';
import { FindBar, type FindNavigation } from './FindBar.js';
import { watchFile } from './watch.js';
import { placeMenuPane, placeOpenMenus } from './menus.js';
import { putHandover, takeHandover, parseHandoverHash } from './persistence.js';

interface Open {
  id: string;
  name: string;
  session: EditorDocument;
  styles: StyleEntry[];
  fonts: string[];
  handle?: FileSystemFileHandle;
  lastModified: number;
  persistence: IndexeddbPersistence;
  /** Lets go of the session: until then the other tabs read it as open here, and offer nothing (ED-011 section 3.6). */
  release: () => void;
  bytes: Uint8Array;
  /** The room the document is shared in, when it is (ED-004 section 4.3). */
  room?: RoomConnection;
  /** Whether this peer joined the room rather than shared it: its bytes are the room's base file. */
  joined?: boolean;
  /** The disk version opened in a new tab (ED-011 section 3.6): shown as "name (on disk)" until any save. */
  onDisk?: true;
  /** The Markdown file the document was made from, watched in place of a handle of its own, until Save As gives it one (ED-011 section 3.8). */
  source?: SourceFile;
}

/**
 * The handle the watcher follows: a Markdown source's, else the document's own. Save As ends a source, so never both. A Flat OPC
 * file is written and never read (ED-005 section 12.122), so it is not watched: Reload would open it. Save in place still refuses
 * one changed on disk.
 */
function watchedHandle(open: Open): FileSystemFileHandle | undefined {
  if (open.source) return open.source.handle;
  return open.handle && formatOfName(open.handle.name) === 'docx' ? open.handle : undefined;
}

/** A status sentence with what a Markdown conversion dropped, where it dropped anything. */
function withDropped(status: string, dropped: string): string {
  return dropped ? `${status} ${t('status.disk.dropped', { dropped })}` : status;
}

/** The open record after a save of any kind: a disk version's "(on disk)" goes (ED-011 section 3.6). */
function saved(open: Open, changes: Partial<Open> = {}): Open {
  const { onDisk: _onDisk, ...rest } = open;
  return { ...rest, ...changes };
}

/** The session record of an open document, for autosave and recovery; while dirty, with what first made it so (ED-011 section 9.9). */
function recordOf(open: Open, dirty: boolean, firstEdit?: EditCause | null): SessionRecord {
  return { id: open.id, name: open.name, bytes: open.bytes, dirty, updated: Date.now(), ...(open.room ? { room: open.room.link } : {}), ...(open.joined ? { joined: true } : {}), ...(dirty && firstEdit ? { firstEdit } : {}) };
}

/** A promise that gives up after `ms`: a relay that never answers leaves the provider retrying for ever. */
function within<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(t('status.room.unreachable'))), ms);
    promise.then((v) => { clearTimeout(timer); resolve(v); }, (e: unknown) => { clearTimeout(timer); reject(e as Error); });
  });
}

/** The page's own address without its fragment, which a room link replaces. */
function siteAddress(): string {
  return location.href.replace(/#.*$/, '');
}

const OPTION_APP = 'docx4j-editor.writeApplication';
const OPTION_SAVE_IN_PLACE = 'docx4j-editor.saveInPlace';
const OPTION_HEADERS = 'docx4j-editor.headersAndFooters';
/** Formatting marks in the perspectives that do not draw them of themselves (ED-005 section 12.80). */
const OPTION_MARKS = 'docx4j-editor.formattingMarks';
/** Close asks before it discards unsaved changes; off by default, Close then discarding at once as it always has (Jason, 2026-11-10). */
const OPTION_ASK_BEFORE_CLOSE = 'docx4j-editor.askBeforeClose';
/** The browser's spell checker on the editing surfaces; on by default, as a browser has it (Jason, 2026-11-11; view/proofing.ts). */
const OPTION_SPELLCHECK = 'docx4j-editor.spellcheck';

function option(key: string, fallback: boolean): boolean {
  try { const v = localStorage.getItem(key); return v === null ? fallback : v === 'true'; } catch { return fallback; }
}

/** The browser's own best match, for the Options entry that follows it. */
function browserMatch(): string {
  const prefs = typeof navigator === 'undefined' ? [] : navigator.languages ?? [];
  return matchLanguage(prefs, availableLanguages()) ?? SOURCE_LANGUAGE;
}

export function App() {
  const tr = useT();
  const lang = useLanguage();
  const [languageChoice, setLanguageChoice] = useState<string | null>(() => languageOption());
  const [open, setOpen] = useState<Open | undefined>();
  const [editor, setEditor] = useState<DocxEditorView | null>(null);
  /** Bound controls whose content is older than their data, counted at open and offered once (ED-005 section 12.89). */
  const [stale, setStale] = useState<{ session: EditorDocument; counts: StaleBoundControls } | null>(null);
  /**
   * The open file written on disk since it was opened (ED-011 section 3.3): the offer to reload, or to keep,
   * with the disk version's text compared to this document's (section 3.4; `comparison` undefined while it
   * runs, `null` when the disk version could not be read).
   */
  const [diskChange, setDiskChange] = useState<{ session: EditorDocument; lastModified: number; comparison?: TextComparison | null } | null>(null);
  /** The confirmation before a reload that discards unsaved changes (ED-011 section 9.9). */
  const [discardAsk, setDiscardAsk] = useState(false);
  /** The question before Close discards unsaved changes, when `askBeforeClose` is on. */
  const [closeAsk, setCloseAsk] = useState(false);
  /** The surface that last took focus (ED-003 section 6.4): the body, or a note or header being edited; the toolbar and the palette act on it. */
  const [active, setActive] = useState<EditorSurface | null>(null);
  const [status, setStatus] = useState<string>('');
  /** The surface whose plain paste read as Markdown, while its offer stands (ED-005 section 12.66). */
  const [markdownOffer, setMarkdownOffer] = useState<EditorView | null>(null);
  /** The fit's offer to break the words it had to break mid-word at their periods and capitals (ED-007 section 8.1 item 13). */
  const [breakOffer, setBreakOffer] = useState<{ surface: EditorSurface; words: string[] } | null>(null);
  const [over, setOver] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [perspective, setPerspective] = useState<PerspectiveId>(() => loadPerspective());
  const [panel, setPanelState] = useState<PanelState>(() => loadPanel(loadPerspective()));
  const [panelWidth, setPanelWidth] = useState<number>(() => loadPanelWidth());
  const [workspaceEl, setWorkspaceEl] = useState<HTMLDivElement | null>(null);
  const [palette, setPalette] = useState(false);
  const [details, setDetails] = useState(false);
  const [selfCheck, setSelfCheck] = useState<{ error: SelfCheckError; diagnosis: Uint8Array; format: SaveFormat } | null>(null);
  const [recovery, setRecovery] = useState<SessionRecord | null>(null);
  const [writeApp, setWriteApp] = useState(() => option(OPTION_APP, true));
  const [inPlace, setInPlace] = useState(() => option(OPTION_SAVE_IN_PLACE, false));
  const [askBeforeClose, setAskBeforeClose] = useState(() => option(OPTION_ASK_BEFORE_CLOSE, false));
  const [spellcheck, setSpellcheckOption] = useState(() => { const on = option(OPTION_SPELLCHECK, true); setSpellcheck(on); return on; });
  const [headersAndFooters, setHeadersAndFooters] = useState(() => option(OPTION_HEADERS, true));
  const [formattingMarks, setFormattingMarks] = useState(() => option(OPTION_MARKS, false));
  /** The companion's address (an option; empty for none) and what its probe answered: Re-paginate shows only with `paginate` among the tools (ED-003 section 6.5). */
  const [companion, setCompanion] = useState<string>(() => companionUrl());
  const [companionInfo, setCompanionInfo] = useState<CompanionInfo | null>(null);
  const companionTools = companionInfo?.tools ?? null;
  /** The companion probed: the option, else this page's own `/companion` (a page the companion serves; ED-005 section 12.28). */
  const companionBase = effectiveCompanion(companion);
  const [repaginating, setRepaginating] = useState(false);
  const [mountEl, setMountEl] = useState<HTMLDivElement | null>(null);
  // The page's zoom (ED-005 section 12.110): CSS `zoom` over the page and its margin, never the chrome; kept per browser.
  const [zoom, setZoomState] = useState(loadZoom);
  const setZoom = useCallback((z: number) => { const next = clampZoom(z); setZoomState(next); saveZoom(next); }, []);
  const [pageAreaEl, setPageAreaEl] = useState<HTMLDivElement | null>(null);
  // Ctrl+wheel over the page zooms the page, as in Word, in place of the browser's whole-page zoom; a native listener,
  // since React registers wheel as passive and the browser's zoom could not be prevented from it.
  useEffect(() => {
    if (!pageAreaEl) return;
    const onWheel = (e: WheelEvent): void => {
      if (!e.ctrlKey || e.deltaY === 0) return;
      e.preventDefault();
      setZoomState((z) => { const next = stepZoom(z, e.deltaY < 0 ? 1 : -1); saveZoom(next); return next; });
    };
    pageAreaEl.addEventListener('wheel', onWheel, { passive: false });
    return () => pageAreaEl.removeEventListener('wheel', onWheel);
  }, [pageAreaEl]);
  const [consoleScript, setConsoleScript] = useState<ConsoleScript | null>(null);
  const consoleSeq = useRef(0);
  /** The perspectives the open document suggests (ED-003 section 4.9), once its custom XML parts are read; none once declined. */
  const [suggested, setSuggested] = useState<{ session: EditorDocument; ids: PerspectiveId[] } | null>(null);
  /** The page size for new documents (an option); the first New asks. */
  const [pageSize, setPageSize] = useState<PageSizePaper | undefined>(() => loadPageSize());
  const [margins, setMargins] = useState<MarginPreset>(() => loadMargins());
  // The open document's default tab stop, in points, as typed (ED-003 section 8.5); written on commit.
  const [defaultTabStop, setDefaultTabStop] = useState<string>('');
  const [askPageSize, setAskPageSize] = useState(false);
  /** A template opened: offered for editing as it is, or as the basis of a new document; dismissed per session. */
  const [templateOffer, setTemplateOffer] = useState<EditorDocument | null>(null);
  /** Macros (ED-003 section 7.4): the browser's list, the open document's, the script being saved, the document macro being shown before its first run. */
  const [browserList, setBrowserList] = useState<Macro[]>(() => browserMacros());
  const [docMacros, setDocMacros] = useState<{ session: EditorDocument; list: Macro[] } | null>(null);
  const [macroSave, setMacroSave] = useState<string | null>(null);
  const [macroShow, setMacroShow] = useState<Macro | null>(null);
  const [askKey, setAskKey] = useState(() => askOption(OPTION_ASK_KEY));
  const [askBaseUrl, setAskBaseUrl] = useState(() => askOption(OPTION_ASK_BASE_URL));
  /** Co-editing (ED-004 section 4): the relay option, the room's status, the link to show, a link waiting for the user's yes, the Join dialog. */
  const [relay, setRelay] = useState<string>(() => relayOption());
  const [roomStatus, setRoomStatus] = useState<RoomStatus | null>(null);
  const [shareUrl, setShareUrl] = useState<string | null>(null);
  const [joinOffer, setJoinOffer] = useState<RoomLink | null>(null);
  const [joinDialog, setJoinDialog] = useState(false);
  const [, setProblemsSeen] = useState(0);
  /** Your name in a shared room (ED-004 section 5); empty: a guest. */
  const [userName, setUserName] = useState<string>(() => nameOption());
  /** The find bar (ED-005 section 3.4): open, with its replacement row, and the last Ctrl+F or Ctrl+H with the selection's text. */
  const [findOpen, setFindOpen] = useState(false);
  const [findReplace, setFindReplace] = useState(false);
  const [findRequest, setFindRequest] = useState<{ seq: number; text?: string }>({ seq: 0 });
  const findNav = useRef<FindNavigation | null>(null);
  const openRef = useRef<Open | undefined>(undefined);
  openRef.current = open;
  const editorRef = useRef(editor);
  editorRef.current = editor;
  /** The caret and view to find again in the document about to open (ED-011 section 3.5), taken when its view has mounted. */
  const placeRef = useRef<KeptPlace | null>(null);
  const activeRef = useRef<EditorSurface | null>(null);
  activeRef.current = active ?? editor;
  const dirtyRef = useRef(false);
  dirtyRef.current = dirty;
  /** What first counted as an edit since the open or the last save (ED-011 section 9.9): kept in the session's record and a log that outlives it. */
  const firstEditRef = useRef<EditCause | null>(null);
  /** The status line's "Saved …" while it is true: the next edit takes it away, unless another message replaced it first. */
  const savedStatusRef = useRef<string | null>(null);
  /** The page's last input, for that record: followed for the application's life. */
  const inputRef = useRef<ReturnType<typeof followInput> | null>(null);
  useEffect(() => { const follower = followInput(); inputRef.current = follower; return () => { follower.stop(); if (inputRef.current === follower) inputRef.current = null; }; }, []);
  /** The modification time Save in place gave the file: the watcher's report of it is the editor's own write, not a change. */
  const ownWriteRef = useRef<number | null>(null);
  /** True while the table fit of a Markdown open writes: the import's work, not an edit, so Reload stays on offer (ED-011 section 3.8). */
  const importFitRef = useRef(false);

  /** The connection for a room, its status followed in the application bar and a refusal said in the status line. */
  const connectRoom = useCallback((store: EditorDocument['store'], link: RoomLink, create = false): RoomConnection => new RoomConnection(store, link, {
    create,
    user: localUser(),
    onStatus: (s) => {
      setRoomStatus(s === 'closed' ? null : s);
      if (s === 'refused') setStatus(t('status.room.refused', { host: relayHost(link.relay) }));
      else if (s === 'unknown') setStatus(t('status.room.unknown', { host: relayHost(link.relay) }));
      else if (s === 'full') setStatus(t('status.room.full'));
    },
  }), []);

  const closeCurrent = useCallback(async (clean: boolean) => {
    const cur = openRef.current;
    if (!cur) return;
    cur.room?.close();
    cur.persistence.destroy();
    unloadEmbeddedFonts(cur.session);
    if (clean) await deleteSession(cur.id);
    cur.release();   // after the delete: a tab checking now must not find the record of a session closed clean
    setOpen(undefined);
    setFindOpen(false);
    setDefaultTabStop('');
    setDirty(false);
    firstEditRef.current = null;
  }, []);

  /** Open bytes into a fresh session, or recover a persisted one. */
  const openBytes = useCallback(async (file: OpenedFile, recover?: SessionRecord, onDisk = false) => {
    performance.mark('docx4j:open');
    await closeCurrent(false);
    const id = recover?.id ?? newSessionId();
    const release = holdSession(id);
    const { store, persistence } = await persistedStore(id);
    performance.mark('docx4j:store');
    // A session recovered in a room: a joiner's bytes are the room's base file, opened over its store as a join (ED-004 section 4.5).
    let session: EditorDocument;
    try {
      // A file Word opens only through a repair is read as Word reads it (Jason, 2026-11-09; ED-008 section 8.8 install item 18 (h)).
      session = await EditorDocument.open(file.bytes, { store, writeApplication: writeApp, repair: 'word', ...(recover?.joined ? { mode: 'join' as const } : {}) });
    } catch (e) {
      persistence.destroy();
      release();
      throw e;
    }
    performance.mark('docx4j:session');
    const [styles, fonts] = await Promise.all([documentStyles(session), documentFonts(session)]);
    performance.mark('docx4j:styles');
    // Recovery rejoins the room; offline, the local copy opens and the provider syncs when it can.
    // A sharer's recovery may remake the room (the idle rule may have deleted it); a joiner's never makes one.
    const room = recover?.room ? connectRoom(store, recover.room, !recover.joined) : undefined;
    const next: Open = { id, name: file.name, session, styles, fonts, lastModified: file.lastModified, persistence, release, bytes: file.bytes, ...(file.handle ? { handle: file.handle } : {}), ...(file.source ? { source: file.source } : {}), ...(room ? { room } : {}), ...(recover?.joined ? { joined: true } : {}), ...(onDisk ? { onDisk: true as const } : {}) };
    setCurrentSessionId(id);
    firstEditRef.current = recover?.dirty ? recover.firstEdit ?? null : null;
    await putSession(recordOf(next, recover?.dirty ?? false, firstEditRef.current));
    setDirty(recover?.dirty ?? false);
    setOpen(next);
    setDefaultTabStop(String(session.defaultTabStop / 20));
    // A file read through Word's repair is said at the open; the load's warnings are in the normalisation report (ED-008 section 8.8 install item 18 (h)).
    setStatus([recover ? t('status.recovered') : '', session.pkg.repaired === 'none' ? '' : t('status.open.repaired')].filter((s) => s !== '').join(' '));
    setDetails(false);
    setRecovery(null);
    setSuggested(null);
    setTemplateOffer(templateKind(session) ? session : null);
    // The document's embedded fonts as faces for the session (ED-003 section 8.4); the page reflows as they load.
    void loadEmbeddedFonts(session).then((fonts) => { const text = describeEmbeddedFonts(fonts); if (text) setStatus((s) => `${s} ${text}`.trim()); });
    // Which perspectives the document suggests: the parts are parsed after the open, the XPath engine not until the perspective asks.
    setDocMacros(null);
    setStale(null);
    setDiskChange(null);
    void session.template.load({ engine: false }).then(async () => {
      setSuggested({ session, ids: perspectives().filter((p) => p.suggest?.(session)).map((p) => p.id) });
      setDocMacros({ session, list: await documentMacros(session) });
      // Bound controls a refresh would change: offered, never done unasked, so an unedited file saves as it was.
      const counts = await staleBoundControls(session, stateFor(session));
      if (counts.total > 0) setStale({ session, counts });
    }).catch(() => { /* a document whose custom XML parts do not parse suggests nothing */ });
  }, [closeCurrent, writeApp, connectRoom]);

  /** A Markdown file as a new document made from it, in the kept page size, its tables fit once drawn; returns what its reading dropped, described. */
  const openMarkdown = useCallback(async (file: OpenedFile, onDisk = false): Promise<string> => {
    const made = await documentFromMarkdown(file, pageSize, margins);
    fitAfterOpen.current = true;
    await openBytes(made.file, undefined, onDisk);
    return describeDropped(made.dropped);
  }, [openBytes, pageSize, margins]);

  /** After a data node was dropped on a control (ED-003 section 4.5's bind by drag): the control shows its node, as a Data tab bind's does (ED-005 section 12.124). */
  const onDrop = useCallback<NonNullable<DocxEditorViewOptions['onDrop']>>(({ bound, drag, target }) => {
    if (!target) setStatus(t('status.drop.noControl'));
    else if (bound) {
      setStatus(t('status.drop.bound', { xpath: drag.xpath }));
      const view = editorRef.current;
      const session = openRef.current?.session;
      if (view && session) void refreshBoundAtCaret(view, session).then((notes) => { for (const note of notes) setStatus(note); }, () => undefined);
    } else setStatus(t('status.drop.notBound', { xpath: drag.xpath }));
  }, []);

  /** A surface took focus, or went away while active: the body is the fallback. */
  const onActivate = useCallback((surface: EditorSurface) => {
    setActive(surface);
    // A header or footer being edited says so in the status line (a note's panel names the note itself).
    const key = surface.story?.key ?? '';
    if (key.startsWith('header:') || key.startsWith('footer:')) setStatus(t('status.editingStory', { story: surface.view.dom.getAttribute('aria-label') ?? t('status.editingStory.header') }));
  }, []);
  const onSurfaceDestroy = useCallback((surface: EditorSurface) => setActive((a) => (a === surface ? surfaces.get('body') ?? null : a)), []);
  /** A paste from Word or elsewhere (ED-003 section 8.3): what was kept and what was dropped, said. */
  const onPaste = useCallback((_report: unknown, text: string) => setStatus(text), []);
  const onUntracked = useCallback((what: string) => setStatus(untrackedStatus(what)), []);
  const onRefused = useCallback((reason: LockRefusal) => setStatus(lockRefusedStatus(reason)), []);
  const onRevealed = useCallback(() => setStatus(t('status.review.revealed')), []);
  const onMarkdownOffer = useCallback((view: EditorView | null, from: EditorView | undefined) => {
    if (view) { setMarkdownOffer(view); setStatus(t('paste.markdown.offer')); }
    else setMarkdownOffer((cur) => (cur === from ? null : cur));
  }, []);
  /** Ctrl+F or Ctrl+H: the bar opens (or takes the focus again) with the selection's text when it lies inside one paragraph. */
  const openFind = useCallback((replace: boolean) => {
    const surface = activeRef.current;
    let text: string | undefined;
    // From the bar's own fields (Ctrl+F to search, Ctrl+H for the replacement row), the query typed there stays: the
    // document's selection is the last hit, and it would put the old term back over the new (Jason, 2026-10-19).
    const inBar = document.activeElement instanceof Element && document.activeElement.closest('.find-bar') !== null;
    if (surface && !inBar) {
      const { from, to, $from, $to } = surface.state.selection;
      if (from !== to && $from.parent === $to.parent && $from.parent.isTextblock) text = surface.state.doc.textBetween(from, to, '\t').slice(0, 255);
    }
    setFindOpen(true);
    if (replace) setFindReplace(true);
    setFindRequest((r) => ({ seq: r.seq + 1, ...(text !== undefined ? { text } : {}) }));
  }, []);
  const findNext = useCallback((back: boolean) => {
    if (!findNav.current) { openFind(false); return; }
    if (back) findNav.current.previous(); else findNav.current.next();
  }, [openFind]);
  const bindFind = useCallback((nav: FindNavigation | null) => { findNav.current = nav; }, []);
  // The next and the previous tracked change (ED-005 section 6.5), through the body view the keys' closure does not see.
  const stepRef = useRef<(dir: 1 | -1) => void>(() => undefined);
  // Track changes on and off (ED-005 section 6.6), Word's Ctrl+Shift+E, through the switch's path: a name is asked for first when there is none.
  const toggleRef = useRef<() => void>(() => undefined);
  // New Comment (ED-005 section 7.5), Word's Ctrl+Alt+M, through the same path.
  const commentRef = useRef<() => void>(() => undefined);
  const shortcuts = useMemo(() => ({
    'Mod-Shift-p': () => setPalette(true),
    'Mod-f': () => openFind(false),
    'Mod-h': () => openFind(true),
    'F3': () => findNext(false),
    'Shift-F3': () => findNext(true),
    'Alt-PageDown': () => stepRef.current(1),
    'Alt-PageUp': () => stepRef.current(-1),
    'Mod-Shift-e': () => toggleRef.current(),
    'Mod-Alt-m': () => commentRef.current(),
  }), [openFind, findNext]);

  // The view lives as long as the session and its mount point: not remade when only the record around it changes (Share adds the room).
  const openSession = open?.session;
  useEffect(() => {
    if (!openSession || !mountEl) return;
    const view = new DocxEditorView(mountEl, openSession, { shortcuts, headersAndFooters, onDrop, onActivate, onDestroy: onSurfaceDestroy, onPaste, onMarkdownOffer, onUntracked, onRefused, onRevealed });
    setEditor(view);
    setActive(view);
    (window as unknown as { docx4jEditor?: DocxEditorView; docx4jSurfaces?: Map<string, EditorSurface> }).docx4jEditor = view;
    (window as unknown as { docx4jSurfaces?: Map<string, EditorSurface> }).docx4jSurfaces = surfaces;
    performance.mark('docx4j:ready');
    // The first frame after the view mounted: what the benchmark calls painted (ED-002 section 7.5).
    requestAnimationFrame(() => requestAnimationFrame(() => performance.mark('docx4j:painted')));
    view.focus();
    // In the frame before the document's first paint, so the page is not seen at its start. Claimed there, guarded by the
    // view being current: an effect's replayed mount spends a claim made in the effect itself.
    let stopHolding: (() => void) | undefined;
    requestAnimationFrame(() => {
      const kept = placeRef.current;
      if (!kept || (view.view as { isDestroyed?: boolean }).isDestroyed) return;
      placeRef.current = null;
      const { found, opening, stop } = restorePlace(view, kept);
      stopHolding = stop;
      const note = found === 'gone' ? t('status.place.gone', { text: opening }) : found === 'unrelated' ? t('status.place.unrelated') : '';
      if (note) setStatus((s) => `${s} ${note}`.trim());
    });
    if (fitAfterOpen.current) {
      fitAfterOpen.current = false;
      // After the first paint (the frame after the view's, as the benchmark's painted mark counts it), so the cells
      // are measured under the document's stylesheet as drawn.
      requestAnimationFrame(() => requestAnimationFrame(() => void (async () => {
        if ((view.view as { isDestroyed?: boolean }).isDestroyed) return;
        await fitImportedRef.current(view.view, 0, view.state.doc.content.size, true);
        openSession.undoManager.clear();
      })()));
    }
    return () => { stopHolding?.(); view.destroy(); setEditor(null); setActive(null); };
  }, [openSession, mountEl, headersAndFooters, onDrop, onActivate, onSurfaceDestroy, onPaste, onMarkdownOffer, onUntracked, onRefused, onRevealed, shortcuts]);

  // Dirty state: any store update since the last successful save; the session record follows.
  useEffect(() => {
    if (!open) return;
    let timer: ReturnType<typeof setTimeout> | undefined;
    // Every change counts, a colleague's included: a save writes everyone's edits (ED-004 section 4.3). The import and the room map are not edits.
    const store = open.session.store;
    const agentTypes = new Set<unknown>([store.agents, store.agentLog]);
    const onTransaction = (tr: Y.Transaction): void => {
      if (importFitRef.current) return;
      // Neither the import, the room map, nor an agent's log or Stop is an edit.
      if (tr.origin === IMPORT_ORIGIN || tr.origin === ROOM_ORIGIN || tr.origin === AGENTS_ORIGIN) return;
      // A peer's arrives under the provider's origin, so it is told by what it changed: the companion's
      // log of every call, reads included, and an invitation clearing a Stop marked the document unsaved
      // (Jason's by-hand run, ED-005 section 12.29). A transaction that changed nothing is not an edit either.
      if ([...tr.changed.keys()].every((type) => agentTypes.has(type))) return;
      // The first one since the open or the last save is described and kept (ED-011 section 9.9): a document said to be
      // changed with nothing edited can then be read afterwards, even once its session is closed and deleted.
      if (!firstEditRef.current) {
        const cause = editCauseOf(tr, store.body, inputRef.current?.seen(), deepStack());
        firstEditRef.current = cause;
        logEditCause({ session: open.id, name: open.name, ...cause });
      }
      // "Saved with changes" is no longer true once the document changes again.
      const saved = savedStatusRef.current;
      if (saved !== null) {
        savedStatusRef.current = null;
        setStatus((s) => (s === saved ? '' : s));
      }
      setDirty(true);
      clearTimeout(timer);
      timer = setTimeout(() => { void putSession(recordOf(open, true, firstEditRef.current)); }, 300);
    };
    store.ydoc.on('afterTransaction', onTransaction);
    return () => { store.ydoc.off('afterTransaction', onTransaction); clearTimeout(timer); };
  }, [open]);

  // The companion probed when a document opens and when its address changes; nothing answers, no button.
  useEffect(() => {
    let cancelled = false;
    setCompanionInfo(null);
    if (!open || companionBase === '') return;
    // Found at the page's own /companion (no option set), a companion counts only when it names an agent: the dev and
    // preview servers mount Re-paginate's test double there, which must not show its button unasked.
    const fallback = companion.trim() === '';
    void probeCompanionInfo(companionBase).then((info) => { if (!cancelled) setCompanionInfo(fallback && !info?.agent ? null : info); });
    return () => { cancelled = true; };
  }, [open, companion, companionBase]);

  /** Re-paginate (ED-003 section 6.5): the saved bytes to the companion, its answer's body adopted through the tree run, one undo step. */
  const doRepaginate = useCallback(async () => {
    if (!open || repaginating) return;
    setRepaginating(true);
    try {
      const saved = await open.session.save({ selfCheck: false });
      const answer = await paginateWith(companionBase, saved.bytes);
      const result = await open.session.adoptPaginated(answer);
      if (result.error) throw result.error instanceof Error ? result.error : new Error(String(result.error));
      setStatus(result.changed ? t('status.repaginated.changed') : t('status.repaginated.unchanged'));
      editor?.refreshCss();
    } catch (e) {
      setStatus(t('status.repaginated.failed', { error: (e as Error).message }));
    } finally {
      setRepaginating(false);
    }
  }, [open, companionBase, repaginating, editor]);

  useEffect(() => {
    const guard = (e: BeforeUnloadEvent): void => { if (dirty) e.preventDefault(); };
    window.addEventListener('beforeunload', guard);
    return () => window.removeEventListener('beforeunload', guard);
  }, [dirty]);

  // The page's language follows the interface's, for screen readers and font selection (the application root carries `dir` too).
  useEffect(() => { document.documentElement.lang = lang; }, [lang]);

  // An open menu follows its button (ED-005 section 12.112): a menu stays open while the bar around it changes.
  useEffect(() => {
    placeOpenMenus();
    window.addEventListener('resize', placeOpenMenus);
    return () => window.removeEventListener('resize', placeOpenMenus);
  }, [perspective, lang, open]);

  // Recovery on the first visit after a crash or a closed tab.
  useEffect(() => { void recoverable().then((r) => { if (r) setRecovery(r); }); }, []);

  /** An offer another tab has overtaken, as tabs a browser restores after a crash all make it: said, and gone, with the session left alone. */
  const overtaken = useCallback(async (record: SessionRecord): Promise<boolean> => {
    if (!(await sessionInUse(record.id))) return false;
    setRecovery(null);
    setStatus(t('status.recover.elsewhere', { name: record.name }));
    return true;
  }, []);

  // A file changed on disk while open (ED-011 section 3.2): watched through its handle, or its Markdown
  // source's (section 3.8), outside a room (the room's document is the peers' shared state, P3). The offer
  // closes by its buttons and is made again for a later write; the editor's own Save in place is not a change.
  useEffect(() => {
    const handle = open && watchedHandle(open);
    if (!open || !handle || open.room) return;
    const { session } = open;
    const watcher = watchFile(handle, open.source?.lastModified ?? open.lastModified, (at) => { if (at !== ownWriteRef.current) setDiskChange({ session, lastModified: at }); });
    return () => watcher.stop();
  }, [open]);

  // The comparison for the offer (ED-011 section 3.4): the disk version's bytes read once per report and
  // the paragraph strings of both sides read through the engine, in the background, never per keystroke;
  // a later report, or the offer closed, drops a result that arrives after it.
  useEffect(() => {
    if (!diskChange || diskChange.comparison !== undefined) return;
    const change = diskChange;
    const cur = openRef.current;
    const handle = cur && watchedHandle(cur);
    if (!cur || !handle || cur.session !== change.session) return;
    let current = true;
    void (async () => {
      let comparison: TextComparison | null = null;
      try {
        const file = await readHandle(handle);
        // A Markdown source is compared as it would open, its conversion; the text does not depend on the page size.
        comparison = await compareWithBytes(change.session, cur.source ? (await documentFromMarkdown(file, undefined)).file.bytes : file.bytes);
      } catch { comparison = null; }
      if (current) setDiskChange((d) => (d === change ? { ...d, comparison } : d));
    })();
    return () => { current = false; };
  }, [diskChange]);

  /**
   * With local edits (ED-011 section 3.6): the disk version in a new browser tab, through a hand-over record
   * written here and taken there. The tab is opened within the click, before the reads, so no pop-up rule
   * closes it; the record's address is given to it once written.
   */
  const openDiskInTab = useCallback(async () => {
    const cur = openRef.current;
    const handle = cur && watchedHandle(cur);
    if (!cur || !handle) return;
    const tab = window.open('', '_blank');
    if (!tab) { setStatus(t('banner.disk.tabBlocked')); return; }
    setDiskChange(null);
    try {
      const file = await readHandle(handle);
      const id = newSessionId();
      const kept = editorRef.current && keepPlace(editorRef.current);
      await putHandover({ id, name: file.name, bytes: file.bytes, lastModified: file.lastModified, handle, ...(cur.source ? { markdown: true as const } : {}), ...(kept ? { place: kept } : {}) });
      tab.location.href = `${siteAddress()}#disk=${id}`;
    } catch (e) { tab.close(); setStatus(t('status.open.failedFile', { name: cur.name, error: (e as Error).message })); }
  }, []);

  /**
   * Reload (ED-011 section 3.5): the handle's file as it is now, a Markdown source converted again, the old session's
   * record gone. With local edits only through the confirmation (`discarding`, section 9.9): they go with the record.
   */
  const reloadFromDisk = useCallback(async (discarding = false) => {
    const cur = openRef.current;
    const handle = cur && watchedHandle(cur);
    if (!cur || !handle) return;
    setDiskChange(null);
    try {
      const file = await readHandle(handle);
      const kept = editorRef.current && keepPlace(editorRef.current);
      await closeCurrent(true);
      placeRef.current = kept;
      const reloaded = t(discarding ? 'status.reloaded.discarded' : 'status.reloaded', { name: file.name });
      if (cur.source) setStatus(withDropped(reloaded, await openMarkdown(file)));
      else { await openBytes(file); setStatus((s) => `${reloaded} ${s}`.trim()); }
    } catch (e) { placeRef.current = null; setStatus(t('status.open.failedFile', { name: cur.name, error: (e as Error).message })); }
  }, [closeCurrent, openBytes, openMarkdown]);

  /**
   * The new tab's side (ED-011 section 3.6): the record taken by the `#disk` hash, the handle tried once
   * (Chromium answers a stored handle's `getFile()` with `NotAllowedError` until a click asks its permission,
   * which the first Save in place does, P5), the bytes opened as a new session named "(on disk)" until a save
   * names it. The hash is cleared first: a reload of this tab opens nothing.
   */
  const openHandover = useCallback(async (id: string) => {
    history.replaceState(null, '', siteAddress());
    try {
      const record = await takeHandover(id);
      if (!record) return;
      let reachable = false;
      if (record.handle) { try { await record.handle.getFile(); reachable = true; } catch { /* asked from the first Save click */ } }
      const file = { name: record.name, bytes: record.bytes, lastModified: record.lastModified };
      placeRef.current = record.place ?? null;
      if (record.markdown) {
        // A source is only read, so no Save ever asks its permission: a handle this tab cannot read is left out, and the tab says so.
        const dropped = await openMarkdown({ ...file, ...(reachable && record.handle ? { handle: record.handle } : {}) }, true);
        setStatus(withDropped(t(reachable ? 'status.disk.openedTab' : 'status.disk.openedTabNotFollowed', { name: record.name }), dropped));
        return;
      }
      await openBytes({ ...file, ...(record.handle ? { handle: record.handle } : {}) }, undefined, true);
      const opened = t(record.handle && !reachable ? 'status.disk.openedTabNoAccess' : 'status.disk.openedTab', { name: record.name });
      setStatus((s) => `${opened} ${s}`.trim());
    } catch (e) { placeRef.current = null; setStatus(t('status.open.failed', { error: (e as Error).message })); }
  }, [openBytes, openMarkdown]);

  useEffect(() => {
    const offer = (): void => { const id = parseHandoverHash(location.hash); if (id) void openHandover(id); };
    offer();
    window.addEventListener('hashchange', offer);
    return () => window.removeEventListener('hashchange', offer);
  }, [openHandover]);

  // A room link in the address (ED-004 section 4.4): asked about before anything connects, since
  // opening a stranger's link would otherwise reach their relay before anything is shown. At load, and
  // whenever the address's hash changes: a link pasted into a tab already on the editor changes only
  // the hash, and the page does not reload (Jason's by-hand run of E4.f, ED-005 section 12.29).
  useEffect(() => {
    const offer = (): void => { const link = parseRoomLink(location.hash); if (link) setJoinOffer(link); };
    offer();
    window.addEventListener('hashchange', offer);
    return () => window.removeEventListener('hashchange', offer);
  }, []);

  // Presence (ED-004 section 5): the session's surfaces draw the peers' carets while it is in a room.
  useEffect(() => {
    if (!open?.room) return;
    setPresence(open.session, open.room.awareness);
    return () => setPresence(open.session, null);
  }, [open]);

  // A name changed in Options reaches the others at once.
  useEffect(() => { open?.room?.setUser(localUser(userName)); }, [open, userName]);
  // The author of this user's tracked changes: the "Your name" option, or the guest name (ED-005 section 5.4).
  useEffect(() => { if (open) open.session.trackingAuthor = localUser(userName).name; }, [open, userName]);

  // What arrived from the room and could not be applied: the banner and the status line follow (section 4.4).
  useEffect(() => {
    if (!open) return;
    let count = open.session.roomProblems.length;
    const heard = (): void => {
      const now = open.session.roomProblems.length;
      if (now > count) setStatus(t('status.room.problems', { count: now }));
      count = now;
      setProblemsSeen((n) => n + 1);
    };
    const offSources = open.session.onSourcesReceived(heard);
    const offParts = open.session.onPartsReceived(heard);
    return () => { offSources(); offParts(); };
  }, [open]);

  /** Share into a relay at `address`: a new room claimed by this connection, the room map written; the link, or undefined when it failed (said). */
  const shareInto = useCallback(async (address: string): Promise<RoomLink | undefined> => {
    const cur = openRef.current;
    if (!cur || cur.room) return undefined;
    const link = newRoomLink(address);
    setStatus(t('status.share.connecting', { host: relayHost(address) }));
    const room = connectRoom(cur.session.store, link, true);
    try {
      await within(room.synced, 20_000);
    } catch (e) {
      room.close();
      if (!(e instanceof RoomRefusedError)) setStatus(t('status.share.failed', { error: (e as Error).message }));
      return undefined;
    }
    await cur.session.share(cur.bytes, cur.name);
    const next: Open = { ...cur, room };
    setOpen(next);
    await putSession(recordOf(next, dirtyRef.current, firstEditRef.current));
    return link;
  }, [connectRoom]);

  /** Share (section 4.4): a new room on the relay, the link shown. */
  const doShare = useCallback(async () => {
    let address: string;
    try { address = relayAddress(relay); } catch { setStatus(t('status.share.badRelay', { relay })); return; }
    if (address === '') return;
    const link = await shareInto(address);
    if (!link) return;
    setShareUrl(roomLinkUrl(siteAddress(), link));
    setStatus(t('status.share.shared'));
  }, [relay, shareInto]);

  /** The companion's relay, where its agent can be invited (ED-005 section 8.2, the browser first). */
  const agentRelay = companionInfo?.agent ? companionRelay(companionBase, companionInfo) : undefined;

  /**
   * Invite an agent (ED-005 section 8.2): a document not yet shared is shared into the companion's
   * relay, and the companion is told the room (`POST <companion>/rooms`); the agent joins as a peer
   * and its edits arrive as its tracked changes. A room on another relay cannot take it.
   */
  const doInvite = useCallback(async () => {
    const cur = openRef.current;
    const agent = companionInfo?.agent;
    if (!cur || !agent || !agentRelay) return;
    let link: RoomLink | undefined;
    if (cur.room) {
      if (cur.room.link.relay !== agentRelay) { setStatus(t('status.invite.otherRelay', { agent })); return; }
      link = cur.room.link;
    } else {
      link = await shareInto(agentRelay);
      if (!link) return;
    }
    setStatus(t('status.invite.inviting', { agent }));
    try {
      await inviteAgent(companionBase, link);
      setStatus(t('status.invite.invited', { agent }));
    } catch (e) {
      setStatus(t('status.invite.failed', { agent, error: (e as Error).message }));
    }
  }, [companionInfo, companionBase, agentRelay, shareInto]);

  /** Join (section 4.4): connect, wait for the room, open its base file over the synced store; a room not ready opens nothing. */
  const joinRoom = useCallback(async (link: RoomLink) => {
    setJoinOffer(null);
    setJoinDialog(false);
    if (location.hash !== '') history.replaceState(null, '', siteAddress());
    await closeCurrent(false);
    const id = `room:${link.room}`;
    const release = holdSession(id);
    const { store, persistence } = await persistedStore(id);
    setStatus(t('status.join.connecting', { host: relayHost(link.relay) }));
    const room = connectRoom(store, link);
    const giveUp = (message?: string): void => { room.close(); persistence.destroy(); release(); if (message) setStatus(message); };
    // A local copy of the room (joined before) opens at once; the provider syncs when it can.
    if (store.room.get('base') === undefined) {
      try { await within(room.synced, 20_000); } catch (e) { giveUp(e instanceof RoomRefusedError ? undefined : t('status.join.failed', { error: (e as Error).message })); return; }
    }
    const base = store.room.get('base') as Uint8Array | undefined;
    if (!base) { giveUp(t('status.join.empty')); return; }
    const name = (store.room.get('name') as string | undefined) ?? 'shared.docx';
    let session: EditorDocument;
    try { session = await EditorDocument.open(base, { store, mode: 'join', writeApplication: writeApp, repair: 'word' }); } catch (e) { giveUp(t('status.join.failed', { error: (e as Error).message })); return; }
    const [styles, fonts] = await Promise.all([documentStyles(session), documentFonts(session)]);
    const next: Open = { id, name, session, styles, fonts, lastModified: 0, persistence, release, bytes: base, room, joined: true };
    setCurrentSessionId(id);
    firstEditRef.current = null;
    await putSession(recordOf(next, false));
    setDirty(false);
    setOpen(next);
    setDefaultTabStop(String(session.defaultTabStop / 20));
    setStatus(t('status.join.joined', { name }));
    setDetails(false);
    setRecovery(null);
    setSuggested(null);
    setTemplateOffer(null);
    setStale(null);
    void loadEmbeddedFonts(session);
    setDocMacros(null);
    void session.template.load({ engine: false }).then(async () => { setDocMacros({ session, list: await documentMacros(session) }); }).catch(() => { /* nothing to suggest */ });
  }, [closeCurrent, connectRoom, writeApp]);

  const switchPerspective = (id: PerspectiveId): void => { setPerspective(id); savePerspective(id); setPanelState(loadPanel(id)); };
  const setPanel = (next: PanelState): void => { setPanelState(next); savePanel(perspective, next); };
  const setTab = (tab: TabId): void => setPanel({ open: true, tab });
  /** Hand a script to the console: the Developer perspective's Reveal codes tab comes to the front, from any perspective. */
  const handToConsole = (text: string, mode: ConsoleScript['mode']): void => {
    if (perspective !== 'developer') { setPerspective('developer'); savePerspective('developer'); }
    const next: PanelState = { open: true, tab: 'codes' };
    setPanelState(next);
    savePanel('developer', next);
    setConsoleScript({ text, mode, seq: ++consoleSeq.current });
  };
  const setWidth = (w: number): void => { setPanelWidth(w); savePanelWidth(w); };

  // ---- Macros (ED-003 section 7.4): a run is a console run without the console, one undo step.
  const currentDocMacros: Macro[] = docMacros && open && docMacros.session === open.session ? docMacros.list : [];
  const perform = useCallback(async (macro: Macro) => {
    const cur = openRef.current;
    const surface = active ?? editor;
    if (!cur || !surface) return;
    setStatus(t('status.macro.running', { name: macro.name }));
    const run = await runMacro(macro, cur.session, surface, { id: perspective, panel });
    setStatus(describeRun(macro, run));
    surface.focus();
  }, [active, editor, perspective, panel]);
  /** A document macro not yet run in this browser is shown first (ED-001 section 11 item 5). */
  const runFrom = useCallback(async (macro: Macro, home: MacroHome) => {
    if (home === 'document' && !(await isTrusted(macro.script))) { setMacroShow(macro); return; }
    await perform(macro);
  }, [perform]);
  const runShown = useCallback(async () => {
    const macro = macroShow;
    if (!macro) return;
    setMacroShow(null);
    await trust(macro.script);
    await perform(macro);
  }, [macroShow, perform]);
  const saveMacro = useCallback(async (name: string, home: MacroHome) => {
    const script = macroSave;
    setMacroSave(null);
    if (script === null) return;
    if (home === 'browser') { setBrowserList(putBrowserMacro({ name, script })); setStatus(t('status.macro.savedBrowser', { name })); return; }
    const cur = openRef.current;
    if (!cur) return;
    try {
      const list = await putDocumentMacro(cur.session, { name, script });
      setDocMacros({ session: cur.session, list });
      setStatus(t('status.macro.savedDocument', { name }));
    } catch (e) { setStatus(t('status.notSaved', { error: reasonOf(e) })); }
  }, [macroSave]);
  const removeMacro = useCallback(async (macro: Macro, home: MacroHome) => {
    if (home === 'browser') { setBrowserList(deleteBrowserMacro(macro.name)); return; }
    const cur = openRef.current;
    if (!cur) return;
    try { setDocMacros({ session: cur.session, list: await deleteDocumentMacro(cur.session, macro.name) }); } catch (e) { setStatus(t('status.notDeleted', { error: (e as Error).message })); }
  }, []);
  useEffect(() => { stepRef.current = (dir) => { if (editor && open) stepChange(editor, open.session, dir, setStatus); }; }, [editor, open]);
  // The tracking switch (ED-005 section 6.6): set while this page changes the mode, so the switch does not report it as a peer's.
  const trackingLocal = useRef(false);
  const [askTrackingName, setAskTrackingName] = useState(false);
  const [askCommentName, setAskCommentName] = useState(false);
  /** Set the document's tracking mode, `toggle` or on, asking for a name first when tracking would turn on with none set. */
  const setTracking = useCallback(async (value: 'toggle' | 'TrackAll', name: string = userName): Promise<void> => {
    if (!editor || !open) return;
    const turningOn = value === 'TrackAll' || open.session.trackingMode === 'Off';
    if (turningOn && open.session.trackingMode === 'TrackAll') return;
    if (turningOn && name.trim() === '') { setAskTrackingName(true); return; }
    open.session.trackingAuthor = localUser(name).name;
    trackingLocal.current = true;
    await editor.runAsync('document.setChangeTrackingMode', { value });
    setStatus(open.session.trackingMode === 'TrackAll' ? t('status.tracking.on', { author: open.session.trackingAuthor ?? '' }) : t('status.tracking.off'));
  }, [editor, open, userName]);
  useEffect(() => { toggleRef.current = () => void setTracking('toggle'); }, [setTracking]);
  /** A comment on the body's selection, or the word at its caret, asking for a name first when none is set (ED-005 section 7.5). */
  const newComment = useCallback(async (name: string = userName): Promise<void> => {
    if (!editor || !open) return;
    if (name.trim() === '') { setAskCommentName(true); return; }
    open.session.trackingAuthor = localUser(name).name;
    if (!(await editor.runAsync('range.insertComment', {}))) setStatus(tr('comments.notDone'));
  }, [editor, open, userName, tr]);
  useEffect(() => { commentRef.current = () => void newComment(); }, [newComment]);
  // The three views of the tracked changes, in the palette beside the macros: a view is not an edit, so not a registry command (ED-005 section 6.4).
  const paletteActions: PaletteAction[] = editor ? [
    ...(['all', 'accepted', 'original'] as const).map((v) => ({ id: `view:${v}`, label: t('palette.view', { name: t(`review.view.${v}`) }), run: () => chooseMarkupView(editor, v, setStatus) })),
    { id: 'print', label: t('palette.print'), run: () => void doPrint() },
    // Save As's picker offers Flat OPC too; without a picker this is the one way to it, as a download.
    { id: 'save-flat-opc', label: t('palette.saveFlatOpc'), run: () => void doSaveAs({ flatOpc: true }) },
    { id: 'zoom-in', label: t('zoom.in'), run: () => setZoom(stepZoom(zoom, 1)) },
    { id: 'zoom-out', label: t('zoom.out'), run: () => setZoom(stepZoom(zoom, -1)) },
    { id: 'zoom-reset', label: t('zoom.reset'), run: () => setZoom(1) },
    ...(editor.view && isInTable(editor.state) ? [{ id: 'fit-table', label: t('palette.fitTable'), run: () => void fitTable() }] : []),
  ] : [];
  const paletteMacros: PaletteMacro[] = [
    ...browserList.map((m) => ({ name: m.name, home: 'browser' as const, run: () => void runFrom(m, 'browser') })),
    ...currentDocMacros.map((m) => ({ name: m.name, home: 'document' as const, run: () => void runFrom(m, 'document') })),
  ];

  /** Print (ED-006 section 3.6): the document alone, its pages laid out by the browser. */
  const doPrint = useCallback(async () => {
    const view = editorRef.current;
    const cur = openRef.current;
    if (!view || !cur) return;
    await printDocument(view, { onStatus: setStatus, prepare: () => prepareNotes(view, cur.session) });
  }, []);
  /** Fit the table's columns to its content (ED-007 section 3.7): the active surface measures, the status line reports. */
  const fitTable = useCallback(async () => {
    const surface = activeRef.current ?? editorRef.current;
    if (!surface) return;
    setStatus(tr('status.tablefit.measuring'));
    const report = await fitTableColumns(surface.view, (name, args) => surface.run(name, args), { onProgress: () => setStatus(tr('status.tablefit.measuring')) });
    const pt = (px: number): number => Math.round((px * 15) / 20);
    let message = '';
    switch (report.outcome) {
      case 'shorter': message = tr('status.tablefit.shorter', { saved: pt(report.before - (report.after ?? report.before)), before: pt(report.before), after: pt(report.after ?? report.before) }) + (report.leftAutoFit ? ` ${tr('status.tablefit.autoFitOff')}` : ''); break;
      case 'unbroken': message = tr('status.tablefit.unbroken', { before: pt(report.before), after: pt(report.after ?? report.before) }); break;
      case 'same': message = tr('status.tablefit.same'); break;
      case 'tooLarge': message = tr('status.tablefit.tooLarge', { max: 2000 }); break;
      case 'tooNarrow': message = tr('status.tablefit.tooNarrow'); break;
      case 'cancelled': message = tr('status.tablefit.cancelled'); break;
    }
    if (report.rowsReleased > 0 && message) message += ` ${tr('status.tablefit.rowsReleased', { count: report.rowsReleased })}`;
    if (report.floorsLetGo && (report.brokenWords ?? 0) > 0 && message) message += ` ${tr('status.tablefit.wordsBreak', { count: report.brokenWords ?? 0 })}`;
    setStatus(message);
    setBreakOffer(message && report.floorsLetGo && report.breakable && report.breakable.length > 0 ? { surface, words: report.breakable } : null);
    // The report is about the table the caret was in: it goes when the caret leaves that table (Jason, 2026-10-14),
    // unless another message has replaced it.
    if (message && report.tablePos !== undefined) {
      let tablePos = report.tablePos;
      const unsubscribe = surface.subscribe((state, transactions) => {
        for (const t of transactions) tablePos = t.mapping.map(tablePos);
        const $from = state.selection.$from;
        let inside = false;
        for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.spec.tableRole === 'table') { inside = $from.before(d) === tablePos; break; }
        if (inside) return;
        unsubscribe();
        setStatus((s) => (s === message ? '' : s));
        setBreakOffer((o) => (o?.surface === surface ? null : o));
      });
    }
    surface.focus();
  }, [tr]);
  /**
   * The tables a Markdown import made, each fit to its content (ED-007 section 8.1 item 14): Markdown states no
   * widths, so the import's proportional grid is a guess the measurement replaces. The fits join the import's undo
   * step where Yjs's capture window allows; the status line says how many were fit. Returns the count.
   */
  const fitImported = useCallback(async (view: EditorView, from: number, to: number, atOpen = false): Promise<number> => {
    const surface = [...surfaces.values()].find((s) => s.view === view);
    if (!surface) return 0;
    // Set around each command alone: a person typing between the fit's measurements is still an edit.
    const run = (name: string, args: unknown): boolean => {
      importFitRef.current = atOpen;
      try { return surface.run(name, args, { joinUndo: true }); } finally { importFitRef.current = false; }
    };
    const reports = await fitTablesIn(view, run, from, to);
    // Said for the tables whose widths changed; one the reader's grid already had at its least height is left as it is, unsaid.
    const count = reports.filter((r) => r.outcome === 'shorter' || r.outcome === 'unbroken').length;
    if (count > 0) setStatus((s) => `${s} ${tr('status.tablefit.imported', { count })}`.trim());
    return count;
  }, [tr]);
  const fitImportedRef = useRef(fitImported);
  fitImportedRef.current = fitImported;
  /** A `.md` file opened: its tables fit once the view has drawn them, then the undo stack cleared, so the fitted widths are the document's. */
  const fitAfterOpen = useRef(false);
  /** The status line's offer taken: the words broken at their periods and capitals, then the table fit again (two undo steps). */
  const breakWords = useCallback(async () => {
    const offer = breakOffer;
    if (!offer) return;
    setBreakOffer(null);
    if (!offer.surface.run('table.breakWords', { words: offer.words })) return;
    await fitTable();
  }, [breakOffer, fitTable]);

  // Ctrl+P (⌘P) prints the document from anywhere in the application, a code editor included, so it is
  // taken in the capture phase; with no document open the key is the browser's (proposal 2).
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (!isPrintKey(e) || !editorRef.current) return;
      e.preventDefault();
      e.stopPropagation();
      void doPrint();
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [doPrint]);

  // Outside the editor (a panel, the toolbar) the same shortcut opens the palette; inside it the
  // view's keymap handles it first, after ProseMirror has read any pending selection change.
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      const inSurface = e.target instanceof Element && e.target.closest('.ProseMirror') !== null;
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        if (!inSurface) setPalette((p) => !p);   // inside a surface its own keymap handled it
        return;
      }
      // Find: inside a surface its keymap handled it; inside Monaco (the console, the XML panes) Monaco's own find opens (ED-005 section 3.8).
      if (inSurface || !openRef.current || (e.target instanceof Element && e.target.closest('.monaco-editor'))) return;
      const key = e.key.toLowerCase();
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && (key === 'f' || key === 'h')) {
        e.preventDefault();
        openFind(key === 'h');
      } else if (e.key === 'F3' && !(e.target instanceof Element && e.target.closest('.find-bar'))) {
        e.preventDefault();
        findNext(e.shiftKey);
      } else if ((e.ctrlKey || e.metaKey) && e.shiftKey && !e.altKey && key === 'e') {
        // Track changes on and off from anywhere (ED-005 section 6.6), as in Word; inside a surface its keymap handled it.
        e.preventDefault();
        toggleRef.current();
      } else if (e.altKey && !e.ctrlKey && !e.metaKey && (e.key === 'PageDown' || e.key === 'PageUp')) {
        // The next and the previous tracked change from the panels too (ED-005 section 6.5).
        e.preventDefault();
        stepRef.current(e.key === 'PageDown' ? 1 : -1);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [openFind, findNext]);

  /** A file opened: a Markdown file as a new document made from it, in the kept page size, what its reading dropped said. */
  const openFile = useCallback(async (file: OpenedFile) => {
    if (!isMarkdownName(file.name)) { await openBytes(file); return; }
    const dropped = await openMarkdown(file);
    setStatus(dropped ? t('status.open.markdownDropped', { name: file.name, dropped }) : t('status.open.markdown', { name: file.name }));
  }, [openBytes, openMarkdown]);

  const onFile = useCallback(async (file: File | undefined, handle?: FileSystemFileHandle) => {
    if (!file) return;
    try { await openFile(await fromFile(file, handle)); } catch (e) { setStatus(t('status.open.failedFile', { name: file.name, error: (e as Error).message })); }
  }, [openFile]);

  /** New: the page size asked the first time, then the document made in the kept size. */
  const startNew = useCallback(async (size?: PageSizePaper) => {
    const chosen = size ?? pageSize;
    if (!chosen) { setAskPageSize(true); return; }
    try { await openBytes(await newDocument(chosen, margins)); } catch (e) { setStatus(t('status.new.failed', { error: (e as Error).message })); }
  }, [openBytes, pageSize, margins]);

  const choosePageSize = (size: PageSizePaper): void => {
    setPageSize(size);
    savePageSize(size);
    setAskPageSize(false);
    void startNew(size);
  };

  const fromTemplate = useCallback(async () => {
    const cur = openRef.current;
    if (!cur) return;
    try { await openBytes(await documentFromTemplate(cur.session)); } catch (e) { setStatus(t('status.fromTemplate.failed', { error: (e as Error).message })); }
  }, [openBytes]);

  const onPick = useCallback(async () => {
    try {
      const picked = await pickFile();
      if (picked) await openFile(picked);
    } catch (e) { setStatus(t('status.open.failed', { error: (e as Error).message })); }
  }, [openFile]);

  /** The status line's Format as Markdown: the offer's paste replaced, one undo step. */
  const formatMarkdown = useCallback(async () => {
    const view = markdownOffer;
    const cur = openRef.current;
    if (!view || !cur) return;
    try {
      const report = await formatAsMarkdown(view, cur.session, markdownOfferOf);
      setStatus(report ? describePaste(report) : t('paste.markdown.changed'));
      if (report?.inserted && report.tables > 0) await fitImported(view, report.inserted.from, report.inserted.to);
    } catch (e) { setStatus(t('paste.failed', { error: (e as Error).message })); }
    setMarkdownOffer(null);
    // Back to the text, so that Undo (or typing on) acts there and not on the status bar.
    if (!(view as { isDestroyed?: boolean }).isDestroyed) view.focus();
  }, [markdownOffer, fitImported]);

  /** How a save went, as the status line says it: `download`, `saveAs`, `inPlace`. */
  const afterSave = useCallback(async (cur: Open, changed: boolean, how: 'download' | 'saveAs' | 'inPlace') => {
    setDirty(false);
    firstEditRef.current = null;
    await putSession(recordOf(cur, false));
    // An edit made while the record was written has already made the document unsaved again: no "Saved" to show.
    if (firstEditRef.current !== null) return;
    const message = t(changed ? 'status.saved.changed' : 'status.saved.unchanged', { how: t(`status.saved.how.${how}`) });
    savedStatusRef.current = message;
    setStatus(message);
  }, []);

  const handleOutcome = useCallback(async (cur: Open, outcome: Awaited<ReturnType<typeof saveAs>>, how: 'download' | 'saveAs' | 'inPlace') => {
    if (outcome.kind === 'saved') await afterSave(cur, outcome.result.changed, how);
    else if (outcome.kind === 'self-check-failed') { setSelfCheck({ error: outcome.error, diagnosis: outcome.diagnosis, format: outcome.format }); setStatus(t('status.save.selfCheckFailed')); }
    else if (outcome.kind === 'changed-on-disk') setStatus(t('status.save.changedOnDisk'));
    else setStatus(t('status.save.cancelled'));
  }, [afterSave]);

  const doDownload = useCallback(async () => {
    if (!open) return;
    try {
      const outcome = await saveAsDownload(open.session, open.name);
      if (outcome.kind === 'saved' && open.onDisk) setOpen(saved(open));
      await handleOutcome(open, outcome, 'download');
    } catch (e) { setStatus(t('status.notSaved', { error: reasonOf(e) })); }
  }, [open, handleOutcome]);

  /** Save As, or with `flatOpc` the palette's Save as Flat OPC XML (ED-005 section 12.122): the name the file is given decides its form. */
  const doSaveAs = useCallback(async (options: { flatOpc?: boolean } = {}) => {
    if (!open) return;
    try {
      const outcome = await saveAs(open.session, open.name, options);
      if (outcome.kind === 'saved' && outcome.handle) {
        const file = await outcome.handle.getFile();
        // The document is now the file just written: a Markdown source ends here (ED-011 section 3.8), and any offer about the old file with it.
        const { source: _source, ...next } = saved(open, { handle: outcome.handle, name: outcome.name ?? open.name, lastModified: file.lastModified });
        setOpen(next);
        setDiskChange(null);
      }
      await handleOutcome(open, outcome, 'saveAs');
    } catch (e) { setStatus(t('status.notSaved', { error: reasonOf(e) })); }
  }, [open, handleOutcome]);

  const doSave = useCallback(async () => {
    if (!open) return;
    if (!open.handle || !inPlace) return doDownload();
    try {
      const outcome = await saveInPlace(open.session, open.handle, open.lastModified);
      if (outcome.kind === 'saved' && outcome.lastModified !== undefined) { ownWriteRef.current = outcome.lastModified; setOpen(saved(open, { lastModified: outcome.lastModified })); }
      await handleOutcome(open, outcome, 'inPlace');
    } catch (e) { setStatus(t('status.notSaved', { error: reasonOf(e) })); }
  }, [open, inPlace, doDownload, handleOutcome]);

  // Ctrl+S (⌘S) saves from anywhere in the application, as Ctrl+P prints: taken in the capture phase,
  // since the browser would otherwise offer to save the page itself; with no document open the key
  // is the browser's. The save is the Save button's when the file can be written in place, else the download.
  const saveRef = useRef(doSave);
  saveRef.current = doSave;
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey || e.key.toLowerCase() !== 's' || !openRef.current) return;
      e.preventDefault();
      e.stopPropagation();
      void saveRef.current();
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, []);

  /** The settings part's `w:defaultTabStop` written and every tab measured again (a part edit, no undo). */
  const commitDefaultTabStop = async (): Promise<void> => {
    if (!open) return;
    const pt = Number(defaultTabStop);
    if (defaultTabStop === '' || !Number.isFinite(pt) || pt <= 0) { setDefaultTabStop(String(open.session.defaultTabStop / 20)); return; }
    const tw = Math.round(pt * 20);
    if (tw === open.session.defaultTabStop) return;
    try {
      await open.session.setDefaultTabStop(tw);
      editor?.refreshCss();
      setStatus(t('status.defaultTabStop', { pt: tw / 20 }));
    } catch (e) { setStatus(t('status.defaultTabStop.failed', { error: (e as Error).message })); }
  };

  const setOption = (key: string, value: boolean): void => {
    try { localStorage.setItem(key, String(value)); } catch { /* ignore */ }
    if (key === OPTION_APP) { setWriteApp(value); open?.session.setOptions({ writeApplication: value }); }
    else if (key === OPTION_HEADERS) setHeadersAndFooters(value);
    else if (key === OPTION_MARKS) setFormattingMarks(value);
    else if (key === OPTION_ASK_BEFORE_CLOSE) setAskBeforeClose(value);
    else if (key === OPTION_SPELLCHECK) { setSpellcheckOption(value); setSpellcheck(value); }
    else setInPlace(value);
  };

  const current = perspectiveNamed(perspective) ?? perspectives()[0]!;
  const summary = open ? diagnosticSummary(open.session.diagnostics, listedChanges(open.session)) : '';
  const fsa = hasFileSystemAccess();
  const offer = suggested && open && suggested.session === open.session ? suggested.ids.filter((id) => id !== perspective).map((id) => perspectiveNamed(id)).filter((p) => p !== undefined) : [];
  const panelShown = panel.open && editor !== null && current.tabs.length > 0;

  return (
    <div className="app" lang={lang} dir={directionOf(lang)} onDragOver={(e) => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); const file = e.dataTransfer.files[0]; void handleOfDrop(e.dataTransfer).then((handle) => onFile(file, handle)); }}>
      <div className="app-bar" role="toolbar" aria-label={tr('app.bar')}>
        <span className="title">docx4j editor</span>
        <button type="button" data-testid="new" onClick={() => void startNew()}>{tr('app.new')}</button>
        {fsa && <button type="button" onClick={() => void onPick()}>{tr('app.open')}</button>}
        {/* The plain file input is the way in for a browser without the File System Access API (and for the tests); with Open… available it stays in the DOM, hidden. */}
        <label className="open" hidden={fsa}>
          {tr('app.openInput')} <input type="file" accept={OPEN_ACCEPT} data-testid="open" onChange={(e) => void onFile(e.target.files?.[0])} />
        </label>
        {open?.handle && inPlace && <button type="button" data-testid="save" onClick={() => void doSave()}>{tr('app.save')}</button>}
        {fsa && <button type="button" data-testid="save-as" onClick={() => void doSaveAs()} disabled={!open}>{tr('app.saveAs')}</button>}
        <button type="button" data-testid="download" onClick={() => void doDownload()} disabled={!open}>{tr('app.download')}</button>
        <button type="button" data-testid="print" onClick={() => void doPrint()} disabled={!open || !editor} title={tr('app.print.title')}>{tr('app.print')}</button>
        {open && companionTools?.includes('paginate') && (
          <button type="button" data-testid="repaginate" onClick={() => void doRepaginate()} disabled={repaginating}
            title={tr('app.repaginate.title')}>{tr('app.repaginate')}</button>
        )}
        {open && !open.room && relay.trim() !== '' && (
          <button type="button" data-testid="share" onClick={() => void doShare()} title={tr('app.share.title')}>{tr('app.share')}</button>
        )}
        {open && companionInfo?.agent && agentRelay && (
          <button type="button" data-testid="invite-agent" onClick={() => void doInvite()}
            disabled={open.room !== undefined && open.room.link.relay !== agentRelay}
            title={open.room !== undefined && open.room.link.relay !== agentRelay ? tr('app.invite.otherRelay', { agent: companionInfo.agent }) : tr('app.invite.title', { agent: companionInfo.agent })}>
            {tr('app.invite', { agent: companionInfo.agent })}
          </button>
        )}
        <button type="button" data-testid="join" onClick={() => setJoinDialog(true)} title={tr('app.join.title')}>{tr('app.join')}</button>
        <button type="button" data-testid="close" onClick={() => { if (askBeforeClose && dirty) setCloseAsk(true); else void closeCurrent(true); }} disabled={!open}>{tr('app.close')}</button>
        {open && <span className="name" data-testid="name">{open.onDisk ? onDiskName(open.name, tr) : open.name}{dirty ? ' •' : ''}</span>}
        {open?.room && roomStatus && (
          <span data-testid="room-status" data-status={roomStatus} className={`room-status ${roomStatus}`} title={tr('room.status.title', { host: relayHost(open.room.link.relay) })}>
            {tr(`room.status.${roomStatus}`)}
          </span>
        )}
        {open?.room && <PeopleList awareness={open.room.awareness} store={open.session.store} onStatus={setStatus}
          noteLabel={(key) => (editor ? decorationsKey.getState(editor.state)?.noteLabels.get(key) : undefined)} />}
        {open?.room && (
          <button type="button" className="link" data-testid="share-show" onClick={() => setShareUrl(roomLinkUrl(siteAddress(), open.room!.link))}>{tr('app.share.link')}</button>
        )}
        {editor && open && <TrackingSwitch editor={editor} author={localUser(userName).name} onToggle={() => void setTracking('toggle')} onStatus={setStatus} local={trackingLocal} />}
        {editor && open && <button type="button" className="new-comment" data-testid="new-comment" title={tr('comments.new.title')} onMouseDown={(e) => e.preventDefault()} onClick={() => void newComment()}>{tr('comments.new')}</button>}
        <span className="spacer" />
        <div className="perspectives" role="tablist" aria-label={tr('app.perspectives')}>
          {perspectives().map((p) => (
            <button key={p.id} type="button" role="tab" aria-selected={p.id === perspective} data-testid={`perspective-${p.id}`} className={p.id === perspective ? 'selected' : ''} onClick={() => switchPerspective(p.id)}>{perspectiveLabel(p)}</button>
          ))}
        </div>
        <button type="button" aria-label={tr('app.palette')} title={tr('app.palette.title')} data-testid="palette-button" onClick={() => setPalette(true)} disabled={!editor}>⌘</button>
        <details className="panels-menu" onToggle={(e) => placeMenuPane(e.currentTarget)}>
          <summary aria-label={tr('panels.menu')}>{tr('panels.menu')}</summary>
          <div className="menu-pane">
            <label><input type="checkbox" data-testid="option-panel" checked={panel.open && current.tabs.length > 0} disabled={current.tabs.length === 0} onChange={(e) => setPanel({ ...panel, open: e.target.checked })} /> {current.tabs.length > 0 ? tr('panels.sidePanel', { tabs: current.tabs.map((tab) => tabLabel(tab).toLocaleLowerCase(lang)).join(', ') }) : tr('panels.sidePanel.none')}</label>
            <button type="button" onClick={() => setPanelState(resetPanel(perspective))}>{tr('panels.reset', { perspective: perspectiveLabel(current) })}</button>
          </div>
        </details>
        <details className="panels-menu macros-menu" onToggle={(e) => placeMenuPane(e.currentTarget)}>
          <summary aria-label={tr('macros.menu')} data-testid="macros-menu">{tr('macros.menu')}</summary>
          <div className="menu-pane">
            {browserList.length + currentDocMacros.length === 0 && <p className="meta">{tr('macros.none')}</p>}
            <ul aria-label={tr('macros.list')}>
              {browserList.map((m) => (
                <li key={`b:${m.name}`} data-testid="macro-entry" data-home="browser">
                  <span className="badge macro-browser" title={tr('macros.home.browser.title')}>{tr('macros.home.browser')}</span><span className="name">{m.name}</span>
                  <button type="button" data-testid="macro-run" disabled={!open} onClick={() => void runFrom(m, 'browser')}>{tr('macros.run')}</button>
                  <button type="button" className="link" data-testid="macro-delete" onClick={() => void removeMacro(m, 'browser')}>{tr('macros.delete')}</button>
                </li>
              ))}
              {currentDocMacros.map((m) => (
                <li key={`d:${m.name}`} data-testid="macro-entry" data-home="document">
                  <span className="badge macro-document" title={tr('macros.home.document.title')}>{tr('macros.home.document')}</span><span className="name">{m.name}</span>
                  <button type="button" data-testid="macro-run" onClick={() => void runFrom(m, 'document')}>{tr('macros.run')}</button>
                  <button type="button" className="link" data-testid="macro-delete" onClick={() => void removeMacro(m, 'document')}>{tr('macros.delete')}</button>
                </li>
              ))}
            </ul>
            {!documentMacrosEnabled() && <p className="meta" data-testid="macros-disabled">{tr('macros.disabled')}</p>}
            <p className="meta">{tr('macros.note')}</p>
          </div>
        </details>
        <details className="panels-menu options-menu" onToggle={(e) => placeMenuPane(e.currentTarget)}>
          <summary aria-label={tr('options.menu')}>{tr('options.menu')}</summary>
          <div className="menu-pane">
            <fieldset>
              <legend>{tr('options.group.thisDocument')}</legend>
              <label>{tr('options.defaultTabStop')}{' '}
                <input type="number" step={0.5} min={0.5} data-testid="option-default-tab-stop" disabled={!open} value={open ? defaultTabStop : ''}
                  onFocus={() => { if (open && defaultTabStop === '') setDefaultTabStop(String(open.session.defaultTabStop / 20)); }}
                  onChange={(e) => setDefaultTabStop(e.target.value)}
                  onBlur={() => void commitDefaultTabStop()}
                  onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); void commitDefaultTabStop(); } }} />
              </label>
            </fieldset>
            <fieldset>
              <legend>{tr('options.group.application')}</legend>
              <label>{tr('options.language')}{' '}
                <select data-testid="option-language" value={languageChoice ?? ''} onChange={(e) => { const v = e.target.value === '' ? null : e.target.value; setLanguageChoice(v); setLanguage(v).catch(() => setStatus(t('status.language.failed'))); }}>
                  <option value="">{tr('options.language.auto', { language: nativeName(browserMatch()) })}</option>
                  {availableLanguages().map((tag) => <option key={tag} value={tag} lang={tag}>{nativeName(tag)}</option>)}
                </select>
              </label>
              {isDraft(lang) && <p className="meta" data-testid="language-draft">{tr('options.language.draft')}</p>}
              <label className="check"><input type="checkbox" data-testid="option-headers" checked={headersAndFooters} onChange={(e) => setOption(OPTION_HEADERS, e.target.checked)} /> {tr('options.headers')}</label>
            </fieldset>
            <fieldset>
              <legend>{tr('options.group.allDocuments')}</legend>
              <label>{tr('options.pageSize')}{' '}
                <select data-testid="option-page-size" value={pageSize ?? ''} onChange={(e) => { const v = isPageSize(e.target.value) ? e.target.value : undefined; setPageSize(v); savePageSize(v); }}>
                  <option value="">{tr('options.pageSize.ask')}</option>
                  {PAGE_SIZES.map((s) => <option key={s.id} value={s.id}>{tr(`pageSize.${s.id}`)}</option>)}
                </select>
              </label>
              <label>{tr('options.margins')}{' '}
                <select data-testid="option-margins" value={margins} onChange={(e) => { const v = isMarginPreset(e.target.value) ? e.target.value : 'normal'; setMargins(v); saveMargins(v); }}>
                  <option value="normal">{tr('sectionDialog.margins.normal')}</option>
                  <option value="narrow">{tr('sectionDialog.margins.narrow')}</option>
                </select>
              </label>
              <label className="check"><input type="checkbox" data-testid="option-app" checked={writeApp} onChange={(e) => setOption(OPTION_APP, e.target.checked)} /> {tr('options.writeApp', { name: APPLICATION_NAME })}</label>
              <label className="check"><input type="checkbox" data-testid="option-save-in-place" checked={inPlace} onChange={(e) => setOption(OPTION_SAVE_IN_PLACE, e.target.checked)} /> {tr('options.saveInPlace')}</label>
              <label className="check"><input type="checkbox" data-testid="option-marks" checked={formattingMarks} onChange={(e) => setOption(OPTION_MARKS, e.target.checked)} /> {tr('options.formattingMarks')}</label>
              <label className="check"><input type="checkbox" data-testid="option-ask-before-close" checked={askBeforeClose} onChange={(e) => setOption(OPTION_ASK_BEFORE_CLOSE, e.target.checked)} /> {tr('options.askBeforeClose')}</label>
              <label className="check"><input type="checkbox" data-testid="option-spellcheck" checked={spellcheck} onChange={(e) => setOption(OPTION_SPELLCHECK, e.target.checked)} /> {tr('options.spellcheck')}</label>
            </fieldset>
            <fieldset>
              <legend>{tr('options.group.sharing')}</legend>
              <label>{tr('options.name')}{' '}
                <input type="text" data-testid="option-name" placeholder={localUser('').name} value={userName}
                  onChange={(e) => { setUserName(e.target.value); setNameOption(e.target.value); }} />
              </label>
              <label>{tr('options.relay')}{' '}
                <input type="url" data-testid="option-relay" placeholder="wss://relay.example.org" value={relay}
                  onChange={(e) => { setRelay(e.target.value); setRelayOption(e.target.value); }} />
              </label>
            </fieldset>
            <fieldset>
              <legend>{tr('options.group.services')}</legend>
              <label>{tr('options.companion')}{' '}
                <input type="url" data-testid="option-companion" placeholder="http://localhost:4174" value={companion}
                  onChange={(e) => { setCompanion(e.target.value); try { localStorage.setItem(OPTION_COMPANION, e.target.value); } catch { /* ignore */ } }} />
              </label>
              <label>{tr('options.askKey')}{' '}
                <input type="password" data-testid="option-ask-key" autoComplete="off" value={askKey} onChange={(e) => { setAskKey(e.target.value); setAskOption(OPTION_ASK_KEY, e.target.value); }} />
              </label>
              <label>{tr('options.askBaseUrl')}{' '}
                <input type="url" data-testid="option-ask-base-url" placeholder="https://api.anthropic.com" value={askBaseUrl} onChange={(e) => { setAskBaseUrl(e.target.value); setAskOption(OPTION_ASK_BASE_URL, e.target.value); }} />
              </label>
            </fieldset>
          </div>
        </details>
      </div>
      {editor && open && current.toolbar !== 'none' && <Toolbar editor={active ?? editor} styles={open.styles} fonts={open.fonts} preset={current.toolbar} controls={current.controls === true} onControl={current.tabs.some((t) => t.id === 'control') ? () => setTab('control') : undefined} onFitTable={() => void fitTable()} />}
      {open && editor && stale?.session === open.session && (
        <div className="banner offer" role="status" data-testid="stale-offer">
          {tr('banner.stale', { count: stale.counts.total })}
          <button type="button" data-testid="stale-refresh" onClick={() => {
            const session = open.session;
            setStale(null);
            applyBoundRefresh(session, editor.state);
            void Promise.all([refreshXhtmlControls(editor.view, session), refreshPictureControls(editor.view, session), refreshFragmentControls(editor, session, { shape: 'w15' })])
              .then((results) => { const notes = [...new Set(results.flatMap((r) => r.notes))]; setStatus(notes.length > 0 ? notes.join(' ') : tr('banner.stale.refreshed')); });
          }}>{tr('banner.stale.refresh')}</button>
          <button type="button" className="link" data-testid="stale-keep" onClick={() => setStale(null)}>{tr('banner.stale.keep')}</button>
        </div>
      )}
      {open && diskChange?.session === open.session && (
        <div className="banner offer" role="status" data-testid="disk-offer">
          {tr(dirty ? 'banner.disk.changedDirty' : 'banner.disk.changed', { name: open.source?.name ?? open.name })}
          {' '}
          <span data-testid="disk-counts">{diskCountsText(diskChange.comparison, tr)}</span>
          {!dirty && <button type="button" data-testid="disk-reload" onClick={() => void reloadFromDisk()}>{tr('banner.disk.reload')}</button>}
          {dirty && <button type="button" data-testid="disk-open-tab" onClick={() => void openDiskInTab()}>{tr('banner.disk.openTab')}</button>}
          {dirty && <button type="button" data-testid="disk-discard" onClick={() => setDiscardAsk(true)}>{tr('banner.disk.discard')}</button>}
          <button type="button" className="link" data-testid="disk-keep" onClick={() => setDiskChange(null)}>{tr('banner.disk.keep')}</button>
        </div>
      )}
      {open && templateOffer === open.session && (
        <div className="banner offer" role="status" data-testid="template-offer">
          {tr('banner.template', { kind: templateKind(open.session) ?? '' })}
          <button type="button" data-testid="template-new" onClick={() => void fromTemplate()}>{tr('banner.template.new')}</button>
          <button type="button" className="link" data-testid="template-edit" onClick={() => setTemplateOffer(null)}>{tr('banner.template.edit')}</button>
        </div>
      )}
      {open && offer.length > 0 && (
        <div className="banner offer" role="status" data-testid="perspective-offer">
          {tr('banner.perspective')}
          {offer.map((p) => <button key={p.id} type="button" data-testid={`offer-${p.id}`} onClick={() => switchPerspective(p.id)}>{tr('banner.perspective.open', { perspective: perspectiveLabel(p) })}</button>)}
          <button type="button" className="link" data-testid="offer-dismiss" onClick={() => setSuggested(null)}>{tr('banner.perspective.dismiss')}</button>
        </div>
      )}
      {joinOffer && (
        <div className="banner offer" role="status" data-testid="join-offer">
          {tr('banner.join', { host: relayHost(joinOffer.relay) })}
          <button type="button" data-testid="join-accept" onClick={() => void joinRoom(joinOffer)}>{tr('banner.join.accept')}</button>
          <button type="button" className="link" data-testid="join-decline" onClick={() => { setJoinOffer(null); history.replaceState(null, '', siteAddress()); }}>{tr('banner.join.decline')}</button>
        </div>
      )}
      {open && open.session.roomProblems.length > 0 && (
        <div className="banner" role="status" data-testid="room-problems">
          {tr('banner.roomProblems', { count: open.session.roomProblems.length })}
          <ul className="meta">
            {open.session.roomProblems.map((p) => <li key={`${p.kind}:${p.key}`}><code>{p.key}</code>: {p.reason}</li>)}
          </ul>
        </div>
      )}
      {recovery && !open && (
        <div className="banner recovery" role="status" data-testid="recovery">
          {tNodes('banner.recovery', { name: <b>{recovery.name}</b> })}
          <button type="button" data-testid="recover" onClick={() => {
            void overtaken(recovery).then((gone) => {
              if (gone) return;
              // A session kept by an earlier version of the editor, or a room of another schema, is refused (ED-005 section 12.3): said, and the banner stays to discard it.
              openBytes({ name: recovery.name, bytes: recovery.bytes, lastModified: 0 }, recovery).catch((e: unknown) => setStatus(t('status.recover.failed', { error: (e as Error).message })));
            });
          }}>{tr('banner.recovery.recover')}</button>
          <button type="button" data-testid="discard" onClick={() => {
            void overtaken(recovery).then((gone) => { if (!gone) { void deleteSession(recovery.id); setRecovery(null); } });
          }}>{tr('banner.recovery.discard')}</button>
        </div>
      )}
      {open && summary !== '' && (
        <div className="banner" role="status" data-testid="banner">
          {summary} <button type="button" className="link" data-testid="banner-details" onClick={() => setDetails((d) => !d)}>{details ? tr('banner.details.hide') : tr('banner.details.show')}</button>
          {/* The one offer of the Review perspective (ED-005 section 6.7): Review registers no `suggest`. */}
          {perspective !== 'review' && perspectiveNamed('review') && (open.session.diagnostics.comments > 0 || listedChanges(open.session) + preservedCount(open.session.diagnostics) > 0) && (
            <>{' '}<button type="button" className="link" data-testid="review-offer" onClick={() => switchPerspective('review')}>{tr('banner.perspective.open', { perspective: perspectiveLabel(perspectiveNamed('review')!) })}</button></>
          )}
          {details && editor && <DiagnosticList editor={editor} entries={open.session.diagnostics.opaque} testId="banner-list" />}
        </div>
      )}
      {findOpen && editor && open && (
        <FindBar editor={editor} session={open.session} replace={findReplace} onReplace={setFindReplace} request={findRequest}
          onClose={(surface) => { setFindOpen(false); surface.focus(); }} onActivate={onActivate} onStatus={setStatus} bind={bindFind} />
      )}
      <div ref={setWorkspaceEl} className={`workspace${panelShown ? ' with-panels' : ''}`} style={panelShown ? { gridTemplateColumns: `minmax(0, 1fr) auto ${panelWidth}px` } : undefined}>
        <div className="page-area" ref={setPageAreaEl}>
          {open
            ? (
              <div className="page-and-margin" style={{ zoom }}>
                <div className={`page-column${current.controlTags ? ' show-control-tags' : ' word-controls'}${current.formattingMarks || formattingMarks ? ' show-marks' : ''}${active?.story && active.story.key.startsWith('footnote') === false && active.story.key.startsWith('endnote') === false ? ' editing-header' : ''}`}>
                  <div ref={setMountEl} data-testid="editor" />
                  {editor && <FootnotesPanel editor={editor} session={open.session} onStatus={setStatus} onActivate={onActivate} onDestroy={onSurfaceDestroy} onPaste={onPaste} onMarkdownOffer={onMarkdownOffer} shortcuts={shortcuts} />}
                </div>
                {editor && <CommentsMargin editor={editor} session={open.session} me={localUser(userName).name} onStatus={setStatus} onActivate={onActivate} onDestroy={onSurfaceDestroy} shortcuts={shortcuts} />}
              </div>
            )
            : <div className={`drop-zone${over ? ' over' : ''}`}>{tr('app.dropZone')}</div>}
        </div>
        {editor && open && panelShown && (
          <>
            <Splitter width={panelWidth} onWidth={setWidth} container={workspaceEl} />
            <DeveloperPanel tabs={current.tabs} panel={panel} onTab={setTab}
              context={{ editor, active: active ?? editor, session: open.session, perspective: { id: perspective, panel }, consoleScript, onConsoleScript: (script) => handToConsole(script, 'replace'), onSaveMacro: (script) => setMacroSave(script), onStatus: setStatus, onTab: setTab, onTrackChanges: () => void setTracking('TrackAll') }} />
          </>
        )}
      </div>
      {/* The status line: a bar of its own, one line tall whatever it says, so that the app bar's buttons
          never squeeze it to nothing (they did at 1440 pixels: Jason's by-hand run, ED-005 section 12.29) and
          a message never moves the page under a click. */}
      <div className="status-bar">
        <span data-testid="status" aria-live="polite" className="status">{status}</span>
        {markdownOffer && open && <button type="button" className="status-action" data-testid="format-markdown" onClick={() => void formatMarkdown()}>{t('paste.markdown.format')}</button>}
        {breakOffer && open && <button type="button" className="status-action" data-testid="break-words" onClick={() => void breakWords()}>{t('status.tablefit.breakWords', { count: breakOffer.words.length })}</button>}
        {editor && open && <TrackingIndicator editor={editor} author={localUser(userName).name} onToggle={() => void setTracking('toggle')} />}
        {open && <ZoomControl zoom={zoom} onZoom={setZoom} />}
      </div>
      {palette && editor && <Palette editor={active ?? editor} onClose={() => setPalette(false)} onInsert={(script) => handToConsole(script, 'insert')} macros={paletteMacros} actions={paletteActions} />}
      {macroSave !== null && <SaveMacroDialog script={macroSave} documentAllowed={documentMacrosEnabled() && open !== undefined} onSave={(name, home) => void saveMacro(name, home)} onClose={() => setMacroSave(null)} />}
      {macroShow && <MacroSourceDialog macro={macroShow} onRun={() => void runShown()} onClose={() => setMacroShow(null)} />}
      {shareUrl && open?.room && <ShareDialog url={shareUrl} relay={open.room.link.relay} onClose={() => setShareUrl(null)} />}
      {joinDialog && <JoinDialog onJoin={(link) => void joinRoom(link)} onClose={() => setJoinDialog(false)} />}
      {askPageSize && <PageSizeDialog onChoose={choosePageSize} onClose={() => setAskPageSize(false)} />}
      {closeAsk && open && (
        <DiscardCloseDialog name={open.name} onClose={() => setCloseAsk(false)}
          onConfirm={() => { setCloseAsk(false); void closeCurrent(true); }} />
      )}
      {discardAsk && open && (
        <DiscardReloadDialog name={open.source?.name ?? open.name} onClose={() => setDiscardAsk(false)}
          onConfirm={() => { setDiscardAsk(false); void reloadFromDisk(true); }} />
      )}
      {askCommentName && (
        <TrackingNameDialog purpose="comments" initial={localUser('').name} onClose={() => setAskCommentName(false)}
          onChoose={(name) => { setAskCommentName(false); setUserName(name); setNameOption(name); void newComment(name); }} />
      )}
      {askTrackingName && (
        <TrackingNameDialog initial={localUser('').name} onClose={() => setAskTrackingName(false)}
          onChoose={(name) => { setAskTrackingName(false); setUserName(name); setNameOption(name); void setTracking('TrackAll', name); }} />
      )}
      {selfCheck && open && (
        <SelfCheckDialog error={selfCheck.error} session={open.session} diagnosis={selfCheck.diagnosis} name={diagnosisName(open.name, selfCheck.format)}
          onDiagnosis={() => download(selfCheck.diagnosis, diagnosisName(open.name, selfCheck.format))} onClose={() => setSelfCheck(null)} />
      )}
    </div>
  );
}

/** The offer's second sentence (ED-011 section 3.4): the counts, the same text, the disk version unread, or the comparison still running. */
function diskCountsText(comparison: TextComparison | null | undefined, tr: (key: string, args?: Record<string, unknown>) => string): string {
  if (comparison === undefined) return tr('banner.disk.comparing');
  if (comparison === null) return tr('banner.disk.compareFailed');
  if (comparison.changed + comparison.added + comparison.removed === 0) return tr('banner.disk.sameText');
  return tr('banner.disk.counts', { changed: comparison.changed, added: comparison.added, removed: comparison.removed });
}

/** A file's name shown with "(on disk)" before its extension (ED-011 section 3.6): "hello (on disk).docx"; the name itself is the file's. */
function onDiskName(name: string, tr: (key: string, args?: Record<string, unknown>) => string): string {
  const m = /^(.*?)(\.[^.]+)?$/.exec(name);
  return tr('file.onDiskName', { base: m?.[1] ?? name, ext: m?.[2] ?? '' });
}
