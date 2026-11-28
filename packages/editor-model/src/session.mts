// SPDX-License-Identifier: GPL-3.0-or-later
//
// An open document: the package (document layer), the projection context, and the Yjs store
// (edit layer), with open, export and save. Headless: no view, no DOM. The view (E1.c) binds
// the store's body fragment with y-prosemirror; tests and agents use `doc` / `replace` directly.
import { addImage, writableWidthEmu, WordprocessingMLPackage, Body, type TextPackage, type RepairMode, type Address, ChangeTracker, highestAnnotationId, type XPathEngine, FootnotesPart, EndnotesPart, FooterPart, HeaderPart, type MainDocumentPart, type CustomXmlPart, linkParents, deepCopy, runOf, p, marshalString, unmarshalString, type Element } from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { initProseMirrorDoc } from 'y-prosemirror';
import type { Node as PMNode, Schema } from 'prosemirror-model';
import * as Y from 'yjs';
import { schema as defaultSchema } from './schema.mjs';
import { projectBody, projectBlocks, newContext, type ProjectionContext } from './import.mjs';
import { isBaseId, randomId, randomRelationshipId, SourceIds } from './ids.mjs';
import { clipboardToken } from './clipboard.mjs';
import { imageBase64Of } from './pictures.mjs';
import { pathIds, assignPathIds, treeEqual, type SourceEntry } from './sources.mjs';
import { PartSync } from './partsync.mjs';
import { exportBody, type ExportResult, type ControlLookup, type RevisionIds } from './export.mjs';
import { writeApplication, appVersionOf } from './appxml.mjs';
import { newEditStore, loadDocument, readDocument, replaceDocument, readControlEntry, readControlEntries, writeControl, type EditStore, type StoreWrite, LOCAL_ORIGIN, IMPORT_ORIGIN, ROOM_ORIGIN, ROOM_SCHEMA, PROJECTION, PROJECTION_KEY, RoomNotReadyError, RoomVersionError, ProjectionVersionError } from './yjs.mjs';
import { compareProjections, SelfCheckError, type SelfCheckResult } from './selfcheck.mjs';
import { mergeRun } from './runmerge.mjs';
import { moveMarkers } from './paginate.mjs';
import { repairedForSave } from './repair.mjs';
import type { Diagnostics } from './diagnostics.mjs';
import { dateUtcWrites } from './revisions.mjs';
import { w15Writes, readControl } from './sdtpr.mjs';
import { PLACEHOLDER_STYLE, namedDefaultPlaceholder, type DefaultPlaceholder } from './bound.mjs';
import { ensureStylesIn } from './knownstyles.mjs';
import { Stories, type Story, type NoteKind, type NoteIds, type CommentIds, type HeaderFooter } from './stories.mjs';
import { textPackageOver } from './textmodel.mjs';
import { COMMENT_STYLES } from './comments.mjs';
import { EditorState } from 'prosemirror-state';
import { TemplateModel } from './template.mjs';
import { DEFAULT_TAB_STOP } from './tabs.mjs';
import { selectIn, type XPathHit, type SelectContainer } from './xpath.mjs';
import { trackingDate, type Tracking } from './track.mjs';
import { changeKey, liveMapping, type TrackedChangeInfo } from './review.mjs';
import { RoomRepair, type ResolutionWrite } from './resolutions.mjs';
import { resolvePreservedOnTree, type PreservedTarget } from './preserved.mjs';

/** The relationship type of a hyperlink's target (ECMA-376 part 1, 17.16.22). */
export const HYPERLINK_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

/** The package-map key a part edit bumps, and the origin it is written under. */
export const PARTS_REVISION = 'partsRevision';
export const PARTS_ORIGIN = 'parts';

/** The editor's version, written to `app.xml` as `AppVersion` (kept equal to package.json by scripts/set-version.mjs). */
export const EDITOR_VERSION = '0.3.0';

export interface OpenOptions {
  schema?: Schema;
  /** The version written to `app.xml`; defaults to `EDITOR_VERSION`. */
  version?: string;
  /** The options pane's switch for the `app.xml` write; default true. */
  writeApplication?: boolean;
  /**
   * Build the package's `PropertyResolver` at open (the styles part read without marking it),
   * for the effective-formatting reads and `w:next` on Enter; default true. Off for a test that
   * wants the projection alone.
   */
  resolver?: boolean;
  /** `Y.UndoManager` capture timeout in ms (typing coalesces within it); default Yjs's 500. */
  captureTimeout?: number;
  /**
   * Use this store instead of a fresh one. When its body already holds content (a persisted
   * session being recovered, ED-002 section 5.4), it is adopted as is and the projection is not
   * loaded into it: the persisted state refers to this document's `sourceId`s, which are
   * deterministic for the same bytes. When it is empty, the projection is loaded as usual.
   */
  store?: EditStore;
  /**
   * `create` (the default) loads the projection into an empty store, as a new session or a
   * recovery does; `join` is a peer entering a room (ED-004 section 3.1): the store must already
   * hold the room's state, and an empty body or a room map this editor's schema does not know
   * refuses, since loading the projection into an unsynced store would duplicate the document
   * once the room's state arrived.
   */
  mode?: 'create' | 'join';
  /**
   * How a file Word opens only through a repair is read (core-ts 0.7.0's `LoadOptions.repair`, CR-031): `refuse`, the
   * engine's default, or `word`, Word's repair with its warning on `pkg.warnings` (a name the container holds twice read
   * from its first entry, several main relationships the last one's, a relationships part in a wrong namespace by local
   * name). The application and the companion open a person's file with `word` (Jason, 2026-11-09 and 2026-11-10; ED-008
   * section 8.8 install item 18 (h)); a save then writes the repaired package, so it is not byte for byte where the
   * repair acted. Ignored for a package passed in.
   */
  repair?: RepairMode;
}

/** What the `room` map records (ED-004 sections 3.1, 3.2 and 4.4). */
export interface RoomRecord {
  schema: number;
  /** The base file, byte for byte: what every peer projects. */
  base: Uint8Array;
  /** Its SHA-256, hex: tells a recovery copy and Save in place that the local base is still the room's. */
  baseHash: string;
  name: string;
}

/** Re-paginate in a room: the companion's answer has a different number of paragraphs from the document (ED-004 section 3.8). */
export class PaginationMismatchError extends Error {
  override name = 'PaginationMismatchError';
  constructor() {
    super('the laid-out document has a different number of paragraphs; the page markers were not moved');
  }
}

/** SHA-256 of bytes, hex (WebCrypto, in the browser and in Node 20). */
export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes as BufferSource);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/** The room a store holds, or undefined when it holds none. */
export function roomOf(store: EditStore): RoomRecord | undefined {
  const schema = store.room.get('schema');
  if (typeof schema !== 'number') return undefined;
  return { schema, base: store.room.get('base') as Uint8Array, baseHash: String(store.room.get('baseHash') ?? ''), name: String(store.room.get('name') ?? '') };
}

export interface SaveOptions {
  /** The document to export; defaults to the store's current body. */
  doc?: PMNode;
  /** Run the round-trip self-check and throw `SelfCheckError` on a mismatch; default true. */
  selfCheck?: boolean;
  /** The file's form (ED-005 section 12.122): a zip, the default, or Flat OPC, the package as one XML document (Word's "Word XML Document"). */
  format?: SaveFormat;
}

/** The forms a save writes: `docx`, the zip, or `flatOpc`, the `pkg:package` XML document. */
export type SaveFormat = 'docx' | 'flatOpc';

export interface SaveResult {
  bytes: Uint8Array;
  /** Whether the main document part was rebuilt (and `app.xml` written). */
  changed: boolean;
  selfCheck?: SelfCheckResult;
}

/** What a tree run receives (ED-002 section 6.1 step 2; the console adds its own on top). */
export interface TreeBindings {
  pkg: WordprocessingMLPackage;
  /** The content API over the tree the projection came from: the script edits the same objects. */
  body: Body;
  main: MainDocumentPart;
  /** The document the tree was exported from, as the edit layer had it when the run began. */
  doc: PMNode;
  /**
   * The tree element a block node of `doc` is (a paragraph, a table, a row, a cell, a block
   * control), for an address (`body.addressOf`) where the edit layer speaks positions (ED-005
   * section 12.27: the agent's tracked changes, and the four commands its resolutions go through).
   * Undefined over a copy (`readTree`'s `copy`), whose elements are not the tree's.
   */
  elementOf(node: PMNode): Element | undefined;
  /** The tree's root, `w:document`, for an XPath over the tree as it is (`bindSnapshot`). */
  document: wml.Document;
  /**
   * The package as the engine's text model (`Texts.of`) is to read it: its main part reading `document`,
   * whether or not the part holds it (ED-008 section 3.1, `textPackageOver`), and over a copy the copy's body.
   */
  textPackage: TextPackage;
}

/**
 * How a tree run is tracked, when not as the session is: the agent's calls are always tracked, as
 * the agent, whatever the document's mode (ED-005 section 8.5, proposal 12), through the engine's
 * `withTracking`, which leaves the settings part and the document's mode as it found them.
 */
export interface TreeRunTracking {
  author: string;
  mode: 'TrackAll' | 'Off';
  date?: Date;
}

/** How a tree run meets a change that reaches the session while it runs (ED-005 proposal 42). */
export interface TreeRunOptions {
  /**
   * An agent's call: the changes that arrived while the callback ran are applied before the merge
   * (a drain), and if one of them reached a block the run replaced or removed, the run is refused:
   * nothing is written, and the session is spent (`TreeRunResult.refused`). Without it, a script's
   * run: the merge puts the run's blocks in place of such a block (ED-004 section 10.2).
   */
  guard?: boolean;
  /**
   * `false`: the run's change is written out of the undo history, as `replaceOutOfHistory` writes (a bound control's
   * refresh from its data, ED-005 proposal 46: the content is the data's, not an edit Undo should take back).
   */
  history?: false;
}

export interface TreeRunResult<T> {
  /** What `fn` returned, when it did not throw. */
  value?: T;
  /** What `fn` threw; the tree is re-imported as it is, there is no rollback. */
  error?: unknown;
  /** Whether the re-import changed the document on screen. */
  changed: boolean;
  /**
   * With `guard`, the blocks the run replaced or removed that another change reached while it ran:
   * each block's index before the run, and its text as the other change left it (none when that
   * change removed it). Nothing was written: no transaction, no part, no comment. The session's tree
   * still holds the refused run, which cannot be taken back in place (the engine edits the tree's
   * elements, and an unchanged block is exported by identity), so the session is spent: its runs,
   * reads and saves throw, and its owner opens a new session on the room (ED-005 proposal 42).
   */
  refused?: { blocks: { before: number; text: string | undefined }[] };
}

/** A turn of the event loop that lets input in: `setImmediate`, after Node's poll for input; a zero timeout elsewhere. */
function yieldToInput(): Promise<void> {
  return new Promise((resolve) => {
    const immediate = (globalThis as { setImmediate?: (f: () => void) => unknown }).setImmediate;
    if (immediate) immediate(resolve); else setTimeout(resolve, 0);
  });
}

/** Copy one projection's side tables into another: every entry (`overwrite`), or only the ids the target lacks. */
function mergeContext(into: ProjectionContext, from: ProjectionContext, overwrite = false): void {
  const copy = <V,>(target: Map<string, V>, source: Map<string, V>): void => {
    for (const [k, v] of source) if (overwrite || !target.has(k)) target.set(k, v);
  };
  copy(into.sources, from.sources);
  copy(into.paragraphs, from.paragraphs);
  copy(into.tables, from.tables);
  copy(into.groups, from.groups);
  copy(into.imported, from.imported);
}

/** Write `ids` into the note references of a typed tree for a save; what it returns puts the room's ids back. */
function applyReferenceIds(root: unknown, ids: NoteIds): () => void {
  const undo: (() => void)[] = [];
  const seen = new Set<object>();
  const walk = (value: unknown): void => {
    if (value === null || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    if (Array.isArray(value)) { value.forEach(walk); return; }
    const node = value as { name?: { localPart?: string }; value?: unknown };
    const local = node.name?.localPart;
    if ((local === 'footnoteReference' || local === 'endnoteReference') && node.value && typeof node.value === 'object') {
      const ref = node.value as { id: number };
      const to = ids.get(local === 'footnoteReference' ? 'footnote' : 'endnote')?.get(Number(ref.id));
      if (to !== undefined) { const from = ref.id; ref.id = to; undo.push(() => { ref.id = from; }); }
    }
    for (const [key, child] of Object.entries(value)) if (key !== 'PARENT') walk(child);
  };
  walk(root);
  return () => { for (const f of undo) f(); };
}

/** The document with the comments made in a room carrying the ids the file was written with, for the self-check (ED-005 section 7.4). */
function withCommentIds(doc: PMNode, ids: CommentIds): PMNode {
  if (ids.size === 0) return doc;
  const tr = EditorState.create({ doc }).tr;
  const comment = doc.type.schema.marks['comment']!;
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    for (const m of node.marks) {
      const to = m.type === comment ? ids.get(Number(m.attrs['id'])) : undefined;
      if (to === undefined) continue;
      tr.removeMark(pos, pos + node.nodeSize, m);
      tr.addMark(pos, pos + node.nodeSize, comment.create({ id: String(to) }));
    }
    return false;
  });
  return tr.doc;
}

/** Whether a run in the document has a character style. */
function usesRunStyle(doc: PMNode, id: string): boolean {
  let found = false;
  doc.descendants((n) => {
    if (found) return false;
    if (n.marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === id)) found = true;
    return !found;
  });
  return found;
}

/** The main part's notes of a read core-ts 0.7.0 makes otherwise: nothing selected (CR-029), a break moved into a run (CR-021), a root Word shows nothing of. */
const REREAD_CODES: ReadonlySet<string> = new Set(['mce.nothing_selected', 'br.outside_run', 'part.unexpected_root']);
/** A highlight objects 0.8.0 reads as another hex, Word's own (CR-014): a mark kept from before holds the old one. */
const MOVED_HIGHLIGHT = /<w:highlight\b[^>]*\bw:val="(?:green|darkBlue|darkCyan|darkGreen|darkMagenta|darkRed|darkYellow|darkGray|lightGray)"/;

/**
 * Whether core-ts 0.7.0 and objects 0.8.0 project this file otherwise than the projection before them (`PROJECTION` 3), as
 * measured over every fixture at the install (ED-008 section 8.8 install item 18 (c)): the main part's notes above, or a
 * highlight the palette moved in any WordprocessingML part, the stories' included (their fragments are kept too).
 * Content read through or skipped (`part.skipped`) projected as before in every fixture that has it.
 */
async function rereadsOtherwise(pkg: WordprocessingMLPackage): Promise<boolean> {
  if (pkg.getMainDocumentPart().warnings.some((w) => REREAD_CODES.has(w.code))) return true;
  const source = pkg.sourcePartStore;
  if (!source) return false;
  const decoder = new TextDecoder();
  for (const name of source.partNames()) {
    if (!/^\/?word\/[^/]+\.xml$/i.test(name)) continue;
    if (MOVED_HIGHLIGHT.test(decoder.decode(await source.load(name)))) return true;
  }
  return false;
}

/** The document with its note references carrying the ids the file was written with, for the self-check. */
function withNoteIds(doc: PMNode, ids: NoteIds): PMNode {
  const tr = EditorState.create({ doc }).tr;
  doc.descendants((node, pos) => {
    if (node.type.name !== 'footnoteRef') return true;
    const to = ids.get(node.attrs['kind'] as 'footnote' | 'endnote')?.get(Number(node.attrs['footnoteId']));
    if (to !== undefined) tr.setNodeMarkup(pos, undefined, { ...node.attrs, footnoteId: to });
    return false;
  });
  return tr.doc;
}

const WORD_COMPAT_URI = 'http://schemas.microsoft.com/office/word';

/** The settings' `compatibilityMode` as Word reads it: the number stated, 12 where none is stated or it is not a number (the engine's reading, core-ts PropertyResolver). */
export function compatibilityModeOf(settings: wml.CTSettings | undefined): number {
  const val = settings?.compat?.compatSetting?.find((cs) => cs.uri === WORD_COMPAT_URI && cs.name === 'compatibilityMode')?.val;
  const mode = typeof val === 'string' ? Number.parseInt(val, 10) : Number.NaN;
  return Number.isFinite(mode) ? mode : 12;
}

export class EditorDocument {
  /** The token the editor's clipboard carries (ED-010 section 3.1): a paste rebuilds a slice only from its own session's. */
  readonly clipboardToken: string = clipboardToken();
  private constructor(
    readonly pkg: WordprocessingMLPackage,
    readonly main: MainDocumentPart,
    /** The main part's tree, from `readContents()`: private until a save with changes hands it to the part. */
    private readonly document: wml.Document,
    readonly schema: Schema,
    private context: ProjectionContext,
    /** The document as imported; the identity fast path on export compares against its nodes. */
    private baseline: PMNode,
    readonly store: EditStore,
    /**
     * Undo over the body and the side maps, scoped to the local origin (ED-001 section 11 item
     * 3; ED-002 section 4.2): a peer's or the agent's edits are never undone by the user.
     */
    readonly undoManager: Y.UndoManager,
    private options: OpenOptions,
  ) {}

  /** Open a docx (zip bytes) or an already loaded package. */
  static async open(source: Uint8Array | WordprocessingMLPackage, options: OpenOptions = {}): Promise<EditorDocument> {
    const pkg = source instanceof Uint8Array ? await WordprocessingMLPackage.load(source, options.repair ? { repair: options.repair } : {}) : source;
    // A save writes run properties as read, never in their normal form (core-ts 0.6.0, CR-015; Jason,
    // 2026-10-26, ED-008 section 8.8): the editor emits an edited part's unchanged paragraphs by identity,
    // so a normalising marshal would rewrite colours through the whole part on one keystroke. Word reads
    // a value outside the normal form as it reads the normal one.
    pkg.normaliseOnSave = false;
    mark('docx4j:zip');
    const main = pkg.getMainDocumentPart();
    const document = await main.readContents();
    mark('docx4j:unmarshalled');
    document.body ??= { TYPE_NAME: 'org_docx4j_wml.Body' };
    // The floor every revision id written afresh goes above (ED-005 section 5.7): the main part's
    // highest from the tree read here, so the engine does not parse the part a second time (1.06 s
    // on the 255-page document), then the parts nobody has read (core-ts 0.2.2, CR-002 section 34).
    await pkg.noteAnnotationIds(main, highestAnnotationId(document));
    await pkg.seedAnnotationIds();
    mark('docx4j:ids');
    if (options.resolver !== false) {
      // The resolver, the run-font selector and the numbering emulator all read their parts
      // privately (`readContents`), so building them costs no part its round trip; the
      // untouched round-trip tests hold that line. The settings part's default tab stop is
      // read the same way (ED-003 section 5.6).
      await pkg.getPropertyResolver();
      await main.getRunFontSelector();
      await pkg.getNumberingEmulator();
    }
    mark('docx4j:resolved');
    const schema = options.schema ?? defaultSchema;
    const { doc, ctx } = projectBody(document.body, schema);
    mark('docx4j:projected');
    const store = options.store ?? newEditStore();
    if (options.mode === 'join') {
      if (store.body.length === 0) throw new RoomNotReadyError('the room\'s state has not arrived');
      if (!roomOf(store)) throw new RoomNotReadyError('the store holds no room');
    }
    // A room of another schema is refused both ways, by a joiner and by its creator's recovery
    // (ED-005 section 5.9): its fragments may hold what this editor would misread.
    const existing = roomOf(store);
    if (existing && existing.schema !== ROOM_SCHEMA) {
      throw new RoomVersionError(`this shared document was made by another version of the editor (room schema ${existing.schema}, this editor ${ROOM_SCHEMA})`);
    }
    // A persisted store (a recovery) made by an earlier projection holds a document with revisions
    // in a shape and with ids this import no longer gives (ED-005 section 12.3).
    if (store.body.length > 0 && store.package.get(PROJECTION_KEY) !== PROJECTION) {
      const kept = Number(store.package.get(PROJECTION_KEY) ?? 1);
      if ((kept < 3 && ctx.diagnostics.tracked.shown > 0) || (kept < 4 && await rereadsOtherwise(pkg))) throw new ProjectionVersionError(`this session was kept by an earlier version of the editor (projection ${String(store.package.get(PROJECTION_KEY) ?? 1)}, this editor ${PROJECTION}); open the file again`);
      store.ydoc.transact(() => store.package.set(PROJECTION_KEY, PROJECTION), IMPORT_ORIGIN);
    }
    // The base projection is over (ED-004 section 3.7): every id minted from here carries this
    // session's Yjs client, so no peer and no recovered session mints an id another has.
    ctx.ids.seal(`c${store.ydoc.clientID}:`);
    if (store.body.length === 0) {
      loadDocument(store, doc);
      // The row- and cell-level groups' properties (ED-003 section 4.2): one map entry per control.
      store.ydoc.transact(() => syncControls(store, ctx), IMPORT_ORIGIN);
    }
    mark('docx4j:loaded');
    // The side maps are in scope: Yjs tracks a fragment nested in a scoped map, so typing in a
    // note or a header is undone by the same manager as typing in the body (ED-003 section 6.2).
    // A resolution's record is in scope too: undoing an accept takes its record back, so the room repair never re-applies an undone resolution (ED-005 section 12.15).
    const undoManager = new Y.UndoManager([store.body, store.package, store.footnotes, store.endnotes, store.comments, store.commentThreads, store.headers, store.controls, store.resolutions], {
      trackedOrigins: new Set([LOCAL_ORIGIN]),
      // A local change marked out of the history (a bound control refreshed from its data, ED-005 section 12.55) is not captured.
      captureTransaction: (t) => t.meta.get('addToHistory') !== false,
      ...(options.captureTimeout !== undefined ? { captureTimeout: options.captureTimeout } : {}),
    });
    const session = new EditorDocument(pkg, main, document, schema, ctx, doc, store, undoManager, options);
    session.partSync = new PartSync(session, source instanceof Uint8Array ? source : undefined);
    // The room repair (ED-005 section 12.15): it acts on remote updates only, so outside a room it does nothing.
    session.roomRepair = new RoomRepair(store);
    if (options.resolver !== false) await session.readSettings();
    // The pristine form of every block of the base file, taken before anything can change one in
    // place, for a tree run's "changed in place" test (ED-004 section 3.7). A deep copy per block,
    // only when the body is first needed by a run: see `pristineOf`.
    // The elements a store names that are not in the base file (a recovery, a join), registered
    // before anything renders, then every entry that arrives later.
    await session.replaySources();
    session.observeSources();
    // The part edits the store holds (a join, a recovery), applied before anything renders (ED-004 section 3.3).
    await session.partSync.replay();
    return session;
  }

  /** Part replication (ED-004 section 3.3): what this session edits in parts goes into the store, what arrives is applied. */
  private partSync!: PartSync;

  /** Resolves once every part edit is written to the store and every arrived one applied. */
  partsReady(): Promise<void> {
    return this.partSync.ready();
  }

  /** Called after a part edit that arrived is applied to this session's tree: the stylesheet, the labels and the panels follow. */
  onPartsReceived(listener: () => void): () => void {
    return this.partSync.onChange(listener);
  }

  /** Part edits that arrived and could not be applied, by map and key. */
  get partProblems(): ReadonlyMap<string, string> {
    return this.partSync.problems;
  }

  /** A part changed by an edit that arrived: written at save like one edited here, not published again. */
  partReceived(partName: string): void {
    this.partsTouched.add(partName);
  }

  /** The settings part's readings (the default tab stop, the decimal symbol, the compatibility mode, the note properties), at open and after the part arrives. */
  async readSettings(): Promise<void> {
    try {
      const settings = await this.main.documentSettingsPart?.readContents();
      const val = settings?.defaultTabStop?.val;
      this.defaultTabStop = typeof val === 'number' && val > 0 ? val : DEFAULT_TAB_STOP;
      const symbol = settings?.decimalSymbol?.val;
      this.decimalSymbol = typeof symbol === 'string' && symbol.length > 0 ? symbol : '.';
      this.compatibilityMode = compatibilityModeOf(settings);
      this.noteSettings = settings ? { ...(settings.footnotePr ? { footnotePr: settings.footnotePr } : {}), ...(settings.endnotePr ? { endnotePr: settings.endnotePr } : {}) } : undefined;
      const track = settings?.trackRevisions;
      this.trackingMode = track !== undefined && track !== null && (track as { val?: boolean }).val !== false ? 'TrackAll' : 'Off';
    } catch { /* an unreadable settings part: Word's default */ }
  }

  /**
   * The document's tracking mode, its settings part's `w:trackRevisions` (ED-005 section 11
   * proposal 4): `TrackAll` or `Off`, a document setting whatever the perspective, read at open and
   * when a peer's settings part arrives. `TrackMineOnly` is not offered: the file cannot store it.
   */
  trackingMode: 'TrackAll' | 'Off' = 'Off';

  /** Whose tracked changes this session's edits are: the application's "Your name" option; the engine's author when it sets none. */
  trackingAuthor: string | undefined;

  /** A fixed date for this session's tracked changes (tests, the agreement suite: ED-005 section 5.8); now, to the minute, otherwise. */
  trackedChangeDate: Date | undefined;

  /**
   * A tracked change's key (ED-005 section 5.5): its start as a Yjs relative position, which
   * survives remote edits around it, and what it is. Given the body view's state, through its live
   * map (section 6.5); else through a map rebuilt from the store, which costs the whole document.
   */
  changeKey(change: TrackedChangeInfo, state?: EditorState): string {
    return changeKey(this.store.body, this.schema, change, liveMapping(state, this.store.body));
  }

  /** The keys of several changes of the document as the store has it now, over one map rebuilt from the store (50 ms on 255 pages) rather than one per change. */
  changeKeys(changes: readonly TrackedChangeInfo[]): string[] {
    if (changes.length === 0) return [];
    const mapping = initProseMirrorDoc(this.store.body, this.schema).mapping;
    return changes.map((c) => changeKey(this.store.body, this.schema, c, mapping));
  }

  /** The author and date a change made now is tracked with, or undefined while tracking is off (the tracker's input: `track.mts`). */
  tracking(): Tracking | undefined {
    if (this.trackingMode === 'Off') return undefined;
    return { author: this.trackingAuthor ?? this.pkg.author.name, date: trackingDate(this.trackedChangeDate) };
  }

  /**
   * Tracking turned on or off: `w:trackRevisions` written through the engine (which makes the
   * settings part when there is none), the part marked changed so a room has it, with no undo, as
   * every settings write (ED-005 section 5.9).
   */
  async setChangeTrackingMode(mode: 'TrackAll' | 'Off'): Promise<void> {
    await this.pkg.setChangeTrackingMode(mode);
    const part = this.main.documentSettingsPart;
    if (part) this.markPartChanged(part.partName.toString());
    this.trackingMode = mode;
  }

  // ---------------------------------------------------------------- the sources map (ED-004 section 3.7)

  /** A block's form as this session last knew it: the base file's, or its `sources` entry's; for the "changed in place" test. */
  private readonly pristine = new Map<string, Element>();
  /** Registrations of arrived entries, in order; `sourcesReady` awaits them. */
  private sourcesPending: Promise<void> = Promise.resolve();
  /** Entries whose registration failed (an XML that does not unmarshal, a path the tree lacks): the finding, by id. */
  readonly sourceProblems = new Map<string, string>();

  private readonly sourceListeners = new Set<() => void>();

  /** Called after an entry that arrived is registered, or fails to be (`sourceProblems`). */
  onSourcesReceived(listener: () => void): () => void {
    this.sourceListeners.add(listener);
    return () => this.sourceListeners.delete(listener);
  }

  /**
   * What arrived from the room and could not be applied (ED-004 section 4.4, "Problems shown"):
   * `sources` entries and part entries, each with its key and the reason. Shown in the
   * diagnostics and counted in the status line, so a bad entry is never silent.
   */
  get roomProblems(): { kind: 'source' | 'part'; key: string; reason: string }[] {
    return [
      ...[...this.sourceProblems].map(([key, reason]) => ({ kind: 'source' as const, key, reason })),
      ...[...this.partSync.problems].map(([key, reason]) => ({ kind: 'part' as const, key, reason })),
    ];
  }

  /** Resolves once every `sources` entry that has arrived is registered: before an export that must see them. */
  sourcesReady(): Promise<void> {
    return this.sourcesPending;
  }

  private async replaySources(): Promise<void> {
    for (const [id, value] of this.store.sources.entries()) await this.registerSource(id, value as SourceEntry);
  }

  private observeSources(): void {
    this.store.sources.observe((event) => {
      for (const [id, change] of event.changes.keys) {
        if (change.action === 'delete') continue;
        const entry = this.store.sources.get(id) as SourceEntry | undefined;
        if (!entry) continue;
        this.sourcesPending = this.sourcesPending.then(() => this.registerSource(id, entry)).then(() => {
          for (const l of [...this.sourceListeners]) l();
        });
      }
    });
  }

  /**
   * An entry made the element of its id on this peer: its XML unmarshalled, the sender's ids
   * assigned by path, the block projected into the context's side tables with those ids (a
   * scratch context, so the diagnostics are not counted twice). Nothing is minted: a projection
   * that had to mint an id means the entry's paths did not fit the tree, which is recorded.
   */
  private async registerSource(id: string, entry: SourceEntry): Promise<void> {
    // A note's element (`note:<kind>:<id>`) and a comment's (`comment:<id>`, a `w:comments` wrapper
    // holding it) are the stories' to adopt (`adoptNotes`, `adoptComments`), not body blocks: projected
    // here, a comment's wrapper came out as an opaque block with a minted id, and every comment made
    // in the editor was reported as a shared change that could not be applied (ED-005 section 12.96).
    if (id.startsWith('note:') || id.startsWith('comment:')) return;
    const ids = this.context.ids;
    const known = ids.get(id);
    if (known && (this.pristine.has(id) || entry.inline)) return;
    // A block written by another peer's tree run carries its revisions' `w16du:dateUtc` as it wrote
    // them: this session's saves declare the prefix for it too (`exportTree`).
    if (!this.adoptedDateUtc && /:dateUtc="/.test(entry.xml)) this.adoptedDateUtc = true;
    try {
      const element = await unmarshalString(entry.xml);
      // Its revisions keep the ids the sender's engine gave them: the ids this session writes
      // afresh go above them, or a save here would give a second revision the same id (ED-005
      // section 12.36).
      this.pkg.annotationIdFloor = highestAnnotationId(element, this.pkg.annotationIdFloor);
      const missing = assignPathIds(element, entry.ids, ids);
      if (missing.length > 0) this.sourceProblems.set(id, `paths not in the entry's tree: ${missing.join(', ')}`);
      if (entry.inline) {
        // A run made by a command (a note's reference): a run source with no holders, nothing to project.
        this.context.sources.set(id, { holders: [] });
        return;
      }
      const minted = ids.mintedCount;
      const scratch = newContext(this.schema, ids);
      projectBlocks([element], scratch, 'sources');
      if (ids.mintedCount !== minted) this.sourceProblems.set(id, `the projection minted ${ids.mintedCount - minted} ids the entry did not name`);
      mergeContext(this.context, scratch, true);
      this.pristine.set(id, deepCopy(element));
    } catch (e) {
      this.sourceProblems.set(id, (e as Error).message);
    }
  }

  /**
   * A run a command made outside a tree run (a note's reference, ED-004 section 3.7), written to
   * `sources` as an inline entry: its XML and the ids of it and its children by path.
   */
  async writeInlineSource(run: Element): Promise<void> {
    const ids = this.context.ids;
    const id = ids.existingIdOf(run);
    if (id === undefined) return;
    const entry: SourceEntry = { xml: await marshalString(run), ids: pathIds(run, ids), inline: true };
    this.store.ydoc.transact(() => this.store.sources.set(id, entry), IMPORT_ORIGIN);
  }

  /** The form a block had before a tree run could touch it: taken the first time a run sees it. */
  private snapshotPristine(content: readonly Element[]): void {
    const ids = this.context.ids;
    for (const el of content) {
      const id = ids.existingIdOf(el);
      if (id !== undefined && !this.pristine.has(id)) this.pristine.set(id, deepCopy(el));
    }
  }

  /** The settings part's `w:defaultTabStop` in twips (720 when it states none), for the tab widths (ED-003 section 5.6). */
  defaultTabStop: number = DEFAULT_TAB_STOP;
  /** The settings part's `w:decimalSymbol` (`.` when it states none): where a decimal tab stop aligns the text after the tab (ED-003 section 11.20). */
  decimalSymbol: string = '.';
  /**
   * Word's compatibility mode (`w:compat/w:compatSetting` `compatibilityMode`): 15 for Word 2013 and later, 14 for Word 2010,
   * and 12 where none is stated, as Word reads it and writes it at save. It decides where a left-aligned table's edge goes
   * (ED-005 section 12.132 item 5, measured in Word 365).
   */
  compatibilityMode: number = 12;

  /** The settings part's `w:footnotePr` and `w:endnotePr`, read at open: the document's note numbering under each section's (ED-003 section 6.3). */
  noteSettings: NoteSettings | undefined;

  /** Whether the document has a notes part: what the note numbering walk costs a document without one. */
  get hasNotes(): boolean {
    return this.main.footnotesPart !== undefined || this.main.endnotesPart !== undefined;
  }

  /**
   * The notes part made when the document has none (ED-003 section 6.3, `prepare`'s part
   * write): `w:footnotes` with the separator and continuation-separator notes (ids -1 and 0)
   * as Word and docx4j write them, live from the start and marked; not undone (section 9 item 7).
   */
  async ensureNotesPart(kind: NoteKind, publish = true): Promise<void> {
    const main = this.main;
    if (kind === 'footnote' ? main.footnotesPart : main.endnotesPart) return;
    const separator = (id: number, type: 'separator' | 'continuationSeparator'): wml.CTFtnEdn => {
      const run = runOf([(type === 'separator' ? el.separator({}) : el.continuationSeparator({})) as Element]);
      const paragraph = p([run]);
      (paragraph.value as wml.P).pPr = { spacing: { after: 0, line: 240, lineRule: 'auto' } };
      return { TYPE_NAME: 'org_docx4j_wml.CTFtnEdn', id, type, content: [paragraph] as wml.CTFtnEdn['content'] };
    };
    const notes = [separator(-1, 'separator'), separator(0, 'continuationSeparator')];
    if (kind === 'footnote') {
      const part = new FootnotesPart();
      main.addTargetPart(part);
      const tree: wml.CTFootnotes = { footnote: notes };
      linkParents(tree, undefined);
      part.setContents(tree);
      if (publish) this.markPartChanged(part.partName.toString()); else this.partReceived(part.partName.toString());
    } else {
      const part = new EndnotesPart();
      main.addTargetPart(part);
      const tree: wml.CTEndnotes = { endnote: notes };
      linkParents(tree, undefined);
      part.setContents(tree);
      if (publish) this.markPartChanged(part.partName.toString()); else this.partReceived(part.partName.toString());
    }
    this.stories.partAdded(kind);
  }

  /**
   * A footer or header part made in the session (ED-005 section 12.95): named free of the parts
   * the package has, added as a target of the main part with a relationship (a random id in a
   * room, as a hyperlink's, so two peers' never collide), its content set, the part published so
   * a peer makes the same (`partsync.mts`, a `part:` entry carrying the kind and the id), and
   * listed with the stories. The caller references it from a section.
   */
  addHeaderFooter(kind: 'header' | 'footer', tree: wml.Hdr | wml.Ftr, options: { rId?: string; partName?: string; publish?: boolean } = {}): HeaderFooter {
    const name = options.partName ?? this.freePartName(kind);
    const part = kind === 'footer' ? new FooterPart(name) : new HeaderPart(name);
    const rel = this.main.addTargetPart(part, undefined, options.rId ?? (this.inRoom ? randomRelationshipId() : undefined));
    linkParents(tree, undefined);
    (part as { setContents(value: unknown): void }).setContents(tree);
    if (options.publish ?? true) this.markPartChanged(name); else this.partReceived(name);
    return this.stories.headerAdded(kind, part, rel.id, tree);
  }

  private freePartName(kind: 'header' | 'footer'): string {
    for (let n = 1; ; n++) {
      const name = `/word/${kind}${n}.xml`;
      if (!this.pkg.parts.get(name)) return name;
    }
  }

  /**
   * The two comment styles (ED-005 section 7.4), spliced when the styles part lacks them by the
   * engine's `styles.ensure` (core-ts 0.3.1, which retired the copy UPSTREAM-11 kept), the part
   * marked and the resolver refreshed; whether anything was added.
   */
  async ensureCommentStyles(): Promise<boolean> {
    return this.ensureStyles([COMMENT_STYLES.text, COMMENT_STYLES.reference]);
  }

  /**
   * The styles a note or a paste needs, spliced from docx4j's defaults into the styles part when
   * the document lacks them (with the styles they link to and are based on) through the engine's
   * `pkg.styles.ensure` (core-ts 0.2.0), and from Word's definitions of the built-ins the engine
   * lacks (`ensureStylesIn`, UPSTREAM-20), the part marked and the resolver refreshed; whether
   * anything was added. An id neither carries is skipped, as before.
   */
  async ensureStyles(ids: readonly string[]): Promise<boolean> {
    const part = this.main.styleDefinitionsPart;
    if (!part) return false;
    if ((await ensureStylesIn(this.pkg, ids)).length === 0) return false;
    this.pkg.refreshPropertyResolver();
    this.markPartChanged(part.partName.toString());
    return true;
  }

  /**
   * The definition half of a list started (ED-003 section 5.3, step C2): a `w:num` over
   * docx4j's default decimal or bullet definition through the engine's paragraph-free
   * `pkg.numbering.newList` (the numbering part and its relationship added when the document
   * has none), the part marked as edited in place (no undo: an undone start leaves the `w:num`
   * unreferenced, which Word tolerates). A `w:num` this session made for the same kind and
   * that no paragraph of `doc` names any more is answered again rather than a new one, which
   * is what keeps Bullets on and off from growing the part. The `w:numId`, as a string.
   */
  async newList(bullet: boolean, doc: PMNode = this.doc): Promise<string> {
    const kind = bullet ? 'bullet' : 'decimal';
    const made = this.madeLists.get(kind);
    if (made !== undefined && !namesList(doc, made)) return made;
    // In a room the new definitions take random ids, so two peers' lists never collide (ED-004
    // section 3.4), chosen here and honoured by the engine (core-ts 0.2.0).
    const ids = this.inRoom ? await this.numberingIds() : undefined;
    const numId = await this.pkg.numbering.newList(ids ? { bullet, numId: randomId(ids.nums), abstractNumId: randomId(ids.abstracts) } : { bullet });
    this.markNumberingChanged();
    this.madeLists.set(kind, numId);
    return numId;
  }
  private readonly madeLists = new Map<'bullet' | 'decimal', string>();

  /** The definition half of a restart: a `w:num` numbering as `numId` does with a `w:startOverride`, through `pkg.numbering.restart`; the new `w:numId` (random in a room). */
  async restartList(numId: string): Promise<string> {
    const ids = this.inRoom ? await this.numberingIds() : undefined;
    const made = this.pkg.numbering.restart(numId, ids ? { newNumId: randomId(ids.nums) } : undefined);
    this.markNumberingChanged();
    return made;
  }

  /** The `w:numId`s and `w:abstractNumId`s the numbering part holds now. */
  private async numberingIds(): Promise<{ nums: Set<number>; abstracts: Set<number> }> {
    const tree = await this.main.numberingDefinitionsPart?.readContents();
    return { nums: new Set((tree?.num ?? []).map((n) => Number(n.numId))), abstracts: new Set((tree?.abstractNum ?? []).map((a) => Number(a.abstractNumId))) };
  }

  /**
   * The relationship half of a link (ED-003 section 5.5, step C4): an external relationship of
   * the main part for the address, through the engine's relationships part (the call the
   * engine's own `Range.hyperlink` makes), and its id for the mark. The relationships part is
   * regenerated by design, so nothing is marked; undo takes the mark off and leaves the
   * relationship, which is harmless (section 9 item 18).
   */
  /**
   * A picture made in the editor (ED-005 proposal 45 part 2, section 12.76): the image part and its
   * relationship added to the main part, the `w:drawing` built by the engine (`addImage`: an inline
   * picture at its natural size, scaled down to the text width, or to `width` points, or `width` by `height`) and registered
   * under a source id, so the export emits it as it emits any inline drawing. The size in points, for
   * the node. In a room the part and the element are this session's: a peer's save has neither.
   */
  addInlinePicture(base64: string, options: { name?: string; width?: number; height?: number } = {}): { sourceId: string; width: number | null; height: number | null } {
    const main = this.main;
    const scope: object = main.isUnmarshalled ? (main.contents as object) : {};
    const sectPr = this.doc.attrs['finalSectPr'] as wml.SectPr | null | undefined;
    const maxCx = writableWidthEmu({ sectPr: sectPr ?? undefined });
    // A data URL as bare base64 (ED-005 section 12.125): the engine's lenient decoder would read its prefix as image bytes.
    const { drawing } = addImage(main, imageBase64Of(base64), scope, maxCx, options);
    const sourceId = this.ctx.ids.idOf(drawing as unknown as Element);
    const inline = (drawing.value as wml.Drawing).anchorOrInline?.[0] as { extent?: { cx?: number; cy?: number } } | undefined;
    const cx = inline?.extent?.cx;
    const cy = inline?.extent?.cy;
    this.partSync.publishRelationships();
    return { sourceId, width: typeof cx === 'number' ? cx / 12700 : null, height: typeof cy === 'number' ? cy / 12700 : null };
  }

  addHyperlink(address: string): string {
    // In a room a random id, so two peers' links never share one (ED-004 section 3.4).
    const id = this.main.getRelationshipsPart(true)!.addExternalRelationship(HYPERLINK_RELATIONSHIP, address, this.inRoom ? randomRelationshipId() : undefined).id;
    this.partSync.publishRelationships();
    return id;
  }

  private markNumberingChanged(): void {
    const part = this.main.numberingDefinitionsPart;
    this.markPartChanged(part ? part.partName.toString() : '/word/numbering.xml');
  }

  // ---------------------------------------------------------------- the room (ED-004 section 3)

  /**
   * A new custom XML part (a data part, the OpenDoPE parts, the document's macros part) through the
   * engine's `customXmlParts.add`. The engine gives it a random `itemID`; in a room the part name
   * is random too (`/customXml/item<n>.xml`, core-ts 0.2.0's `partName` option), so two peers
   * making a part at once never take the same name, and the entry published for it carries the
   * name, so a peer that lacks the part makes it the same (`partsync.mts`). The caller marks it.
   */
  addCustomXmlPart(xml: string): CustomXmlPart {
    if (!this.inRoom) return this.pkg.customXmlParts.add(xml);
    const used = new Set<number>();
    for (const part of this.pkg.parts.values()) {
      const m = /^\/customXml\/item(\d+)\.xml$/.exec(part.partName.toString());
      if (m) used.add(Number(m[1]));
    }
    return this.pkg.customXmlParts.add(xml, { partName: `/customXml/item${randomId(used)}.xml` });
  }

  /** Whether the session is in a room: its store records one. */
  get inRoom(): boolean {
    return this.store.room.has('schema');
  }

  /** The room the session is in, or undefined. */
  get room(): RoomRecord | undefined {
    return roomOf(this.store);
  }

  /**
   * Make the session a room's creator (ED-004 sections 3.1, 3.9, 4.4): every note's and header's
   * fragment made, so no peer ever has to, then the `room` map written with the base file, its
   * hash and name, under the room origin (not undone). `base` is the bytes this session was
   * opened from. Whatever the session did before (a tree run, a macro) is already in `sources`,
   * so there is nothing else to sweep. A session already in a room is left as it is.
   */
  async share(base: Uint8Array, name: string): Promise<RoomRecord> {
    const existing = this.room;
    if (existing) return existing;
    await this.stories.fillAll();
    // A session opened from a package learns its base file here (ED-004 section 10.2 item 7).
    this.partSync.adoptBase(base);
    await this.partSync.ready();
    const baseHash = await sha256Hex(base);
    this.store.ydoc.transact(() => {
      this.store.room.set('base', base);
      this.store.room.set('baseHash', baseHash);
      this.store.room.set('name', name);
      this.store.room.set('schema', ROOM_SCHEMA);
    }, ROOM_ORIGIN);
    return this.room!;
  }

  /** The projection context: ids, sources, the imported nodes, diagnostics. Rebuilt by a tree run. */
  get ctx(): ProjectionContext {
    return this.context;
  }

  /** The document as imported, or as re-imported by the last tree run: the "unchanged" baseline. */
  get imported(): PMNode {
    return this.baseline;
  }

  get body(): wml.Body {
    return this.document.body!;
  }

  /** The other stories (footnotes, endnotes, comments, headers and footers), read on first use; notes and headers editable through their fragments (ED-003 section 6.2). */
  get stories(): Stories {
    return (this.storiesCache ??= new Stories(this));
  }
  private storiesCache: Stories | undefined;

  /** Replace a story's document as one local transaction (headless editing), as `replace` is for the body. */
  replaceStory(story: Story, doc: PMNode, origin: unknown = LOCAL_ORIGIN): void {
    this.stories.replace(story, doc, origin);
  }

  /** The template view (ED-003 section 4): the custom XML parts, the OpenDoPE parts and the verdicts; `await template.load()` once. */
  get template(): TemplateModel {
    return (this.templateCache ??= new TemplateModel(this));
  }
  private templateCache: TemplateModel | undefined;

  /** A row- or cell-level group's properties from the `controls` map, for export. */
  readonly controlEntry: ControlLookup = (groupId) => readControlEntry(this.store, groupId);

  /** True once a tree run handed the tree to the main part: it is live and re-marshalled on save. */
  get treeLive(): boolean {
    return this.live;
  }
  private live = false;

  /**
   * A part other than the main one was edited in place (a style through the Styles tab): the
   * package map takes a revision under its own origin, so the application's dirty state and
   * autosave see a change while the undo manager, which tracks the local origin only, does
   * not (there is no undo for a part edit in E1); `save` then writes `app.xml` as for an edit.
   */
  markPartChanged(partName: string): void {
    this.partsTouched.add(partName);
    this.store.ydoc.transact(() => { this.store.package.set(PARTS_REVISION, (Number(this.store.package.get(PARTS_REVISION)) || 0) + 1); }, PARTS_ORIGIN);
    // What changed in the part, into the store (ED-004 section 3.3), in and out of a room.
    this.partSync.published(partName);
  }
  /** The parts edited in place since the open, by name. */
  get touchedParts(): ReadonlySet<string> {
    return this.partsTouched;
  }
  private readonly partsTouched = new Set<string>();

  /**
   * The settings part's `w:defaultTabStop` written, in twips (ED-003 section 8.5, step F4): a
   * part edit under `markPartChanged` (live, marshalled at save, no undo), the session's own
   * reading updated so the next measurement uses it; the view is told to measure again by the
   * caller. Refused for a value that is not a positive number or a document without a settings
   * part.
   */
  async setDefaultTabStop(twips: number): Promise<void> {
    if (typeof twips !== 'number' || !Number.isFinite(twips) || twips <= 0) throw new Error(`not a tab stop interval: ${twips}`);
    const part = this.main.documentSettingsPart;
    if (!part) throw new Error('the document has no settings part');
    const settings = await part.readContents();
    settings.defaultTabStop = { val: Math.round(twips) };
    linkParents(settings, undefined);
    part.setContents(settings);
    this.defaultTabStop = Math.round(twips);
    this.markPartChanged(part.partName.toString());
  }

  /** Change an option after open (the options pane's `app.xml` switch). */
  setOptions(patch: Pick<OpenOptions, 'writeApplication' | 'version'>): void {
    Object.assign(this.options, patch);
  }

  get diagnostics(): Diagnostics {
    return this.ctx.diagnostics;
  }

  /** The current document, read from the Yjs store. */
  get doc(): PMNode {
    return readDocument(this.store, this.schema);
  }

  /** Replace the document as one local transaction (headless editing), with a command's control-map writes in the same step; `before`, the document the store held, narrows the formatting check (`syncTextFormatting`) to what differs. */
  replace(doc: PMNode, origin: unknown = LOCAL_ORIGIN, writes: readonly StoreWrite[] = [], resolution?: ResolutionWrite, before?: PMNode): void {
    replaceDocument(this.store, doc, origin, writes, resolution, before);
  }

  /**
   * Replace the document as a local change kept out of the undo history: a bound control refreshed
   * from its data (ED-005 section 12.55), which a part edit made and an undo must not take back.
   * `before`, the document the store held, narrows the write as `replace` does.
   */
  replaceOutOfHistory(doc: PMNode, before?: PMNode): void {
    this.store.ydoc.transact((t) => {
      t.meta.set('addToHistory', false);
      replaceDocument(this.store, doc, LOCAL_ORIGIN, [], undefined, before);
    }, LOCAL_ORIGIN);
  }

  /** The room repair over this session's store (ED-005 section 12.15). */
  roomRepair!: RoomRepair;

  /**
   * The element one block node stands for: the original when the block is unchanged, else the
   * rebuilt one (reveal codes, ED-001 section 3.5). Pure: nothing in the tree is touched.
   */
  exportBlock(node: PMNode): Element {
    const doc = this.schema.topNodeType.create({ finalSectPr: this.doc.attrs['finalSectPr'] }, [node]);
    const result = exportBody(doc, this.ctx, { content: [] }, this.controlEntry);
    return result.content[0]!;
  }

  /**
   * Export without saving: the body's new content array and whether anything changed. `ids`
   * numbers the revisions it writes afresh; without it they count on from the body's highest id,
   * which is right for an export nobody saves (reveal codes, the preview, Select by XPath).
   */
  exportTree(doc: PMNode = this.doc, ids?: RevisionIds): ExportResult {
    const before = dateUtcWrites();
    const beforeW15 = w15Writes();
    const result = exportBody(doc, this.ctx, this.body, this.controlEntry, ids);
    // A revision written with `w16du:dateUtc` needs the prefix declared ignorable on the part's
    // root, as the engine's tracker declares it (core-ts 0.3.0): the objects package then declares
    // its namespace, and a consumer that does not know `w16du` skips it rather than failing. That
    // holds for one this export wrote, and for one a peer's tree run wrote into a block this
    // session adopted (ED-005 section 12.36: without it, a person's save of the agent's changes
    // left the prefix undeclared).
    if (dateUtcWrites() !== before || this.adoptedDateUtc) this.declareIgnorable('w16du');
    // A control's properties holding a `w15` element likewise: Word 2010 knows no `w15` and opens the
    // part only with the prefix ignorable (ED-005 section 12.58).
    if (w15Writes() !== beforeW15) this.declareIgnorable('w15');
    return result;
  }

  /** A prefix listed in `mc:Ignorable` on the main part's root; the objects package declares its namespace. */
  private declareIgnorable(prefix: string): void {
    const root = this.document as { ignorable?: string };
    const prefixes = (root.ignorable ?? '').split(/\s+/).filter((p) => p.length > 0);
    if (!prefixes.includes(prefix)) root.ignorable = [...prefixes, prefix].join(' ');
  }

  /** The tree element of a node of the document an export went over: the one it became, else its `sourceId`'s. */
  private elementLookup(exported: ExportResult): (node: PMNode) => Element | undefined {
    return (node) => {
      const made = exported.nodes.get(node);
      if (made) return made;
      const id = node.attrs['sourceId'];
      return typeof id === 'string' ? this.context.ids.get(id) : undefined;
    };
  }

  /**
   * The ids a save's or a tree run's exports give the revisions they write afresh (ED-005 section
   * 5.7): the engine's annotation counter, above every annotation id in every WordprocessingML
   * part (seeded at load), and the one its tracker issues from, so an id the editor writes is
   * never one the engine gives out in a tree run, and the other way round.
   */
  private revisionIds(): RevisionIds {
    const tracker = new ChangeTracker(this.pkg, 'TrackAll');
    return () => tracker.nextId();
  }

  /**
   * Save: export, hand the tree to the part only when a block changed (so an unedited document
   * keeps the main part's source bytes; ED-002 section 8 item 2), write `app.xml` on a change,
   * zip, and run the self-check.
   */
  /**
   * The one queue a save, a tree run and Re-paginate go through, each after the last has finished
   * (ED-004 section 10.9 item 1). A save writes the file's note ids into the live trees for its
   * zip and puts the room's back after it, and a tree run exports those same trees, so the two
   * must never overlap. A script's callback must not itself save or run: it would wait on itself.
   */
  private exclusive: Promise<unknown> = Promise.resolve();

  private serially<T>(work: () => Promise<T>): Promise<T> {
    const next = this.exclusive.then(work, work);
    this.exclusive = next.catch(() => undefined);
    return next;
  }

  save(options: SaveOptions = {}): Promise<SaveResult> {
    return this.serially(() => this.saveNow(options));
  }

  private async saveNow(options: SaveOptions): Promise<SaveResult> {
    this.notSpent();
    await this.sourcesReady();
    await this.partsReady();
    await this.partSync.normalizeOrder();
    // Tables merged broken by concurrent edits are written repaired (ED-004 section 6.1); the store is not changed.
    const doc = repairedForSave(options.doc ?? this.doc, (id) => this.context.imported.get(id));
    const ids = this.revisionIds();
    const result = this.exportTree(doc, ids);
    if (result.changed) this.applyExport(result);
    // The stories (ED-003 section 6.2): a note or header part is written where a fragment
    // differs, a note the body no longer references is removed, one made here is appended.
    for (const name of await this.stories.exportAll(doc, ids)) this.partsTouched.add(name);
    // The comments (ED-005 section 7.4), over the body's tree as the export left it: a dropped
    // comment's markers go with it, so the body can change here too.
    const comments = await this.stories.exportComments(doc, this.body, ids);
    for (const name of comments.touched) this.partsTouched.add(name);
    const bodyChanged = result.changed || comments.bodyChanged;
    if (bodyChanged || this.live || this.partsTouched.size > 0) {
      // A tree run made the part live already; an edit-layer change hands the tree over now.
      if (bodyChanged && !this.live) this.main.setContents(this.document);
      if (bodyChanged) await this.ensurePlaceholders(doc);
      const version = this.options.version ?? EDITOR_VERSION;
      await writeApplication(this.pkg, { version, enabled: this.options.writeApplication ?? true });
      // An OpenDoPE part touched: the XPaths part names this editor as the authoring tool (REQ-011).
      if (this.templateCache?.touched) this.templateCache.stampAuthoringTool(appVersionOf(version) ?? version);
    }
    // In a room, the notes made there are written with small ids Word 2010 accepts (ED-004
    // section 10.8 item 6): set in the notes parts and the body's references for the zip, then
    // put back, so the session and the room keep the random ids that keep peers' notes apart.
    const noteIds: NoteIds = this.inRoom ? this.stories.noteIdsForFile(doc) : new Map();
    const restore = noteIds.size > 0 ? [this.stories.applyNoteIds(noteIds), applyReferenceIds(this.document, noteIds)] : [];
    // The comments made in a room the same way (section 7.4).
    const commentIds: CommentIds = this.inRoom ? this.stories.commentIdsForFile(doc) : new Map();
    if (commentIds.size > 0) restore.push(this.stories.applyCommentIds(commentIds, this.body));
    let bytes: Uint8Array;
    try {
      // Flat OPC (ED-005 section 12.122): the same parts as the zip's, each XML part's text as the zip would hold it, in one
      // `pkg:package` document, UTF-8, which Word opens as a Word XML document.
      bytes = options.format === 'flatOpc' ? new TextEncoder().encode(await this.pkg.saveFlatOpc()) : await this.pkg.save();
    } finally {
      for (const undo of restore) undo();
    }
    const changed = bodyChanged || this.partsTouched.size > 0;
    if (options.selfCheck === false) return { bytes, changed };
    const checked = withCommentIds(noteIds.size > 0 ? withNoteIds(doc, noteIds) : doc, commentIds);
    const selfCheck = await this.selfCheck(bytes, checked, noteIds, commentIds);
    if (!selfCheck.equal) throw new SelfCheckError(selfCheck);
    return { bytes, changed, selfCheck };
  }

  /**
   * The `PlaceholderText` style written into the styles part when a body the save writes has a run in
   * it and the part lacks it, as Word 15 writes it when it saves. Word 2010 drops a `w:rStyle` naming a
   * style the part lacks (core-ts check 35, `35e-word2010.docx`; ED-005 section 12.59). The engine's
   * `styles.ensure` splices Word's definition since core-ts 0.3.2 (UPSTREAM-19, closed).
   */
  private async ensurePlaceholders(doc: PMNode): Promise<void> {
    if (usesRunStyle(doc, PLACEHOLDER_STYLE)) await this.ensureStyles([PLACEHOLDER_STYLE]);
    // The glossary entries the editor's own placeholders name (Word's form, ED-008 section 8.8 item 14): required
    // of the engine, which writes them into the glossary at this save, making the part where the file has none
    // (core-ts CR-002 section 63, D9); an entry the glossary holds is left as it is.
    const named: DefaultPlaceholder[] = [];
    doc.descendants((n) => {
      if (n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') {
        const entry = readControl(n.attrs['sdtPr']).showingPlaceholder ? namedDefaultPlaceholder(n.attrs['sdtPr']) : undefined;
        if (entry && !named.includes(entry)) named.push(entry);
      }
      return !n.isAtom;
    });
    if (named.length === 0) return;
    const glossary = this.pkg.glossaryEntries;
    await glossary.load(this.pkg.getMainDocumentPart().glossaryDocumentPart);
    for (const entry of named) glossary.require(entry.name, entry.text);
  }

  /** Put an export's content into the tree (the private one, or the live one after a tree run). */
  private applyExport(result: ExportResult): void {
    const body = this.body;
    body.content = result.content as wml.Body['content'];
    if (result.sectPr === undefined) delete body.sectPr; else body.sectPr = result.sectPr;
    linkParents(body, this.document);
  }

  /**
   * Decision 20, the protocol between the edit layer and the tree (ED-002 section 6.1): export
   * the edit layer into the tree, run `fn` against the tree through the content API, hand the
   * tree to the main part (live from now on, since a tree run is an edit whether or not the
   * projection can see it), then project the body again with the same ids and replace the
   * document as one transaction, so untouched blocks keep their nodes and the run is one undo
   * step. No rollback: a throwing script has done what it did, and undo takes the run back.
   */
  runOnTree<T>(fn: (bindings: TreeBindings) => T | Promise<T>, origin: unknown = LOCAL_ORIGIN, tracking?: TreeRunTracking, options: TreeRunOptions = {}): Promise<TreeRunResult<T>> {
    return this.serially(() => this.runOnTreeNow(fn, origin, tracking, options));
  }

  /** Set when a guarded run was refused: the tree holds it, so nothing may read or write the tree again. */
  private spent = false;

  /** Set when a block adopted from `sources` carries a `w16du:dateUtc`, which the saves must then declare. */
  private adoptedDateUtc = false;

  private notSpent(): void {
    if (this.spent) throw new Error('this session is spent: a guarded tree run was refused, and its tree still holds the refused run; open a new session on the room');
  }

  private async runOnTreeNow<T>(fn: (bindings: TreeBindings) => T | Promise<T>, origin: unknown, tracking?: TreeRunTracking, options: TreeRunOptions = {}): Promise<TreeRunResult<T>> {
    this.notSpent();
    await this.sourcesReady();
    const before = this.doc;
    const beforeItems = this.store.body.toArray();
    // A table merged irregular by concurrent edits is exported repaired, and the run's change carries the repair into the store (ED-004 section 6.1).
    const revisionIds = this.revisionIds();
    const repairedDoc = repairedForSave(before, (id) => this.context.imported.get(id));
    const exported = this.exportTree(repairedDoc, revisionIds);
    if (exported.changed) this.applyExport(exported);
    // The tree goes to the part before the callback (ED-003 section 9 item 27): `pkg.body`,
    // `main.body` and the `Word` shim's document all reach it through `getContents()`, which
    // would otherwise unmarshal a second tree from the source bytes and lose the run's edits.
    if (!this.live) { this.main.setContents(this.document); this.live = true; }
    // The comments on screen into their part, so a script sees them (ED-005 section 7.4).
    await this.stories.beforeTreeRun(before, this.body, revisionIds);
    // Every block the run can see, as it is now (ED-004 section 3.7's "changed in place" test).
    this.snapshotPristine(this.body.content ?? []);
    const body = new Body(this.main, this.body, 'body', this.pkg);
    const exportedDoc = repairedDoc;
    const elementOf = this.elementLookup(exported);
    const out: TreeRunResult<T> = { changed: false };
    try {
      // The engine tracks a tree run's edits when the document's tracking is on, as this session's
      // author, without writing the setting; with it off, a mode the engine read earlier cannot
      // track them (ED-005 section 5.4: two trackers, one document setting).
      const author = { name: tracking?.author ?? this.trackingAuthor ?? this.pkg.author.name };
      const mode = tracking?.mode ?? this.trackingMode;
      const date = tracking?.date ?? this.trackedChangeDate;
      out.value = await this.pkg.withTracking({ author, mode, ...(date ? { date } : {}) }, () => fn({ pkg: this.pkg, body, main: this.main, doc: exportedDoc, elementOf, document: this.document, textPackage: this.pkg }));
    } catch (e) {
      out.error = e;
    }
    // A guarded run (an agent's call) lets the changes that arrived while it ran reach the store
    // first: two turns of the event loop, so that one polls for input between them (ED-005 proposal 42).
    if (options.guard) { await yieldToInput(); await yieldToInput(); }
    // Every block not known as it is (new, rebuilt by the export, or changed in place by the
    // script) goes through its own XML, so this session holds the objects every other peer and
    // any recovered session will unmarshal, and gets fresh ids: the `sources` entries below.
    const entries = await this.enterNewBlocks();
    const { doc: fresh, ctx } = projectBody(this.body, this.schema, this.context.ids);
    // Only the run's own change (ED-004 section 10.2 item 1): the store may have moved on while
    // the callback awaited, and replacing it with `fresh` whole would revert that.
    const current = this.doc;
    const merged = mergeRun(this.schema, before, beforeItems, fresh, current, this.store.body.toArray());
    if (options.guard) {
      // A block the run replaced or removed that another change reached meanwhile: refused, nothing written.
      // Without a merge (the store's blocks did not line up) any change meanwhile would be reverted: refused too.
      const conflicts = merged ? merged.conflicts : current.eq(before) ? [] : [{ before: -1, now: undefined }];
      if (conflicts.length > 0) {
        this.spent = true;
        return { ...out, changed: false, refused: { blocks: conflicts.map((c) => ({ before: c.before, text: c.now?.textContent })) } };
      }
    }
    const previous = this.context;
    this.context = ctx;
    this.baseline = fresh;
    const written = new Map<string, SourceEntry>();
    for (const { element, xml } of entries) {
      const id = ctx.ids.existingIdOf(element);
      if (id === undefined) continue;
      written.set(id, { xml, ids: pathIds(element, ctx.ids) });
      this.pristine.set(id, deepCopy(element));
    }
    // Ids churn for every rebuilt element; the change that matters is the one a reload would show.
    out.changed = !compareProjections(before, fresh).equal;
    this.undoManager.stopCapturing();
    this.store.ydoc.transact((t) => {
      if (options.history === false) t.meta.set('addToHistory', false);
      // The entries first, in the same transaction: a receiver never sees a node before its source.
      for (const [id, entry] of written) this.store.sources.set(id, entry);
      if (merged && merged.kept.some((b) => !ctx.imported.has(b.attrs['sourceId'] as string))) {
        // A block kept as the store has it whose id the fresh projection did not give (the run's
        // export rebuilt it): the previous projection's entries still describe it.
        mergeContext(ctx, previous);
      }
      replaceDocument(this.store, merged?.doc ?? fresh, origin, [], undefined, current);
      syncControls(this.store, ctx);
    }, origin);
    this.templateCache?.invalidate();
    // A run of the content API can change styles, numbering and the main part's hyperlinks too, and comments.
    this.partSync.publishAll();
    await this.stories.afterTreeRun(origin);
    return out;
  }

  /**
   * The body's blocks that are not known as they are (ED-004 section 3.7): a block without an
   * id (the script made it, or the run's export rebuilt it), or whose id's known form differs
   * (the script changed it in place), replaced in the tree by its own XML's unmarshalling, whose
   * objects have no ids, so the projection that follows mints new ones. What each became, with its XML.
   */
  private async enterNewBlocks(): Promise<{ element: Element; xml: string }[]> {
    const content = this.body.content ?? [];
    const ids = this.context.ids;
    const out: { element: Element; xml: string }[] = [];
    for (let i = 0; i < content.length; i++) {
      const el = content[i] as Element;
      const id = ids.existingIdOf(el);
      const known = id !== undefined && (this.pristine.has(id) ? treeEqual(el, this.pristine.get(id)) : isBaseId(id));
      if (known) continue;
      // The copy is all new objects, so it has no ids and the projection mints them; the old
      // element keeps whatever ids it had (an export-rebuilt block may share run objects with a
      // base block, whose ids an undo can still need).
      const xml = await marshalString(el);
      const copy = await unmarshalString(xml);
      (content as Element[])[i] = copy;
      out.push({ element: copy, xml });
    }
    if (out.length > 0) linkParents(this.body, this.document);
    return out;
  }

  /**
   * Select by XPath (ED-005 section 3.5): the expression over the main document part, or over a
   * notes, header or footer part named by `part`, as the edit layer has it, the hits placed where
   * they project. The edit layer is exported into the tree first, as a save does, and never
   * handed to the part (so an unedited document keeps its bytes and nothing is re-imported); the
   * call waits its turn with the saves and tree runs (`serially`), since a save writes ids into
   * the same trees. Throws the engine's error for an invalid expression.
   */
  selectXPath(xpath: string, options: SelectXPathOptions = {}): Promise<XPathHit[]> {
    return this.serially(() => this.selectXPathNow(xpath, options));
  }

  private async selectXPathNow(xpath: string, options: SelectXPathOptions): Promise<XPathHit[]> {
    await this.sourcesReady();
    let root: Element;
    let containers: SelectContainer[];
    if (options.part === undefined || options.part === this.main.partName.toString()) {
      const doc = this.doc;
      const exported = this.exportTree(repairedForSave(doc, (id) => this.context.imported.get(id)));
      if (exported.changed) this.applyExport(exported);
      root = el.document(this.document) as unknown as Element;
      containers = [{ content: (this.body.content ?? []) as Element[], doc }];
    } else {
      const tree = await this.stories.selectTree(options.part);
      if (!tree) throw new Error(`${options.part} is not a part Select by XPath runs over`);
      root = tree.root;
      containers = tree.containers;
    }
    return selectIn(root, containers, xpath, options.engine ?? this.pkg.xpathEngine, this.context.ids, options.namespaces);
  }

  /**
   * The revisions shown in Word only, resolved through the engine in a tree run (`preserved.mts`,
   * core-ts 0.3.0): every revision left (`body.acceptAll()` or `rejectAll()`, once the command's
   * transaction has resolved the edit layer's, which also takes what the engine never lists: a
   * cell merge, a list's numbering), an author's, or one entry's. Untracked, as every resolution;
   * the number the engine resolved.
   */
  async resolvePreserved(action: 'accept' | 'reject', which: { all: true } | { author: string } | { target: PreservedTarget }): Promise<number> {
    const run = await this.runOnTree(({ body }) => {
      if ('all' in which) return action === 'accept' ? body.acceptAll() : body.rejectAll();
      return resolvePreservedOnTree(body, this.body, action, which);
    }, LOCAL_ORIGIN, { author: this.trackingAuthor ?? this.pkg.author.name, mode: 'Off' });
    if (run.error) throw run.error;
    return run.value ?? 0;
  }

  /**
   * Where a block an engine address names is in the edit layer (ED-005 section 12.28: a click on the
   * agent's log goes there): the position before its node, or undefined when nothing is at the
   * address now. The address in any of the engine's forms (`body/3`, `{ paraId }`, `{ contains }`).
   */
  positionOf(address: Address): Promise<number | undefined> {
    return this.readTree(({ body, doc, elementOf }) => {
      const block = body.elementAt(address);
      if (!block) return undefined;
      const target = block.element.value;
      let found: number | undefined;
      doc.descendants((node, pos) => {
        if (found !== undefined) return false;
        if (elementOf(node)?.value === target) { found = pos; return false; }
        return true;
      });
      return found;
    });
  }

  /**
   * The tree as the document is now, for reading (ED-005 section 8.9, the agent's reading tools
   * and its dry runs): the edit layer exported into the session's private tree, as Select by XPath
   * does, and `fn` run with a `Body` over it; the tree is never handed to the part and nothing is
   * projected again or written to the store. With `copy`, `fn` gets a deep copy of the body, which
   * it may change freely (a dry run makes its change there and reports it); `tracking` then tracks
   * what it does, as a tree run's is tracked. With `stories`, the session's notes, headers and comments
   * are written into their parts first, as a save writes them, for a reader of the parts themselves (the
   * engine's text model, `textPackage`; ED-008 section 3.1): without it the parts hold what they held at the
   * last save or tree run. Waits its turn with the saves and tree runs.
   */
  readTree<T>(fn: (bindings: TreeBindings) => T | Promise<T>, options: { copy?: boolean; tracking?: TreeRunTracking; stories?: boolean } = {}): Promise<T> {
    return this.serially(async () => {
      this.notSpent();
      await this.sourcesReady();
      const doc = repairedForSave(this.doc, (id) => this.context.imported.get(id));
      const exported = this.exportTree(doc);
      if (exported.changed) this.applyExport(exported);
      if (options.stories) await this.storiesIntoParts(doc);
      const container = options.copy ? deepCopy(this.body) : this.body;
      if (options.copy) linkParents(container, undefined);
      const body = new Body(options.copy ? undefined : this.main, container, 'body', this.pkg);
      const elementOf = options.copy ? () => undefined : this.elementLookup(exported);
      const document = options.copy ? { ...this.document, body: container } : this.document;
      const textPackage = textPackageOver(this.pkg, document);
      const run = (): T | Promise<T> => fn({ pkg: this.pkg, body, main: this.main, doc, elementOf, document: this.document, textPackage });
      if (!options.tracking) return run();
      const t = options.tracking;
      return this.pkg.withTracking({ author: { name: t.author }, mode: t.mode, ...(t.date ? { date: t.date } : {}) }, run);
    });
  }

  /**
   * The session's notes, headers and comments into their parts, as a save writes them (`stories.exportAll`,
   * `stories.exportComments`) and as a tree run writes the comments (`beforeTreeRun`): a part is written only
   * where the session's story differs from it, and is then one the next save writes, as it would have been.
   */
  private async storiesIntoParts(doc: PMNode): Promise<void> {
    for (const name of await this.stories.exportAll(doc)) this.partsTouched.add(name);
    const comments = await this.stories.exportComments(doc, this.body);
    for (const name of comments.touched) this.partsTouched.add(name);
  }

  /**
   * Re-paginate's reload (ED-003 section 6.5; ED-001 Appendix D.1): the companion answered with
   * the document laid out and its `w:lastRenderedPageBreak` markers rewritten (and a `w14:paraId`
   * on every paragraph). The answer's body replaces the tree's through decision 20's protocol:
   * one tree run, the re-projection with the same ids giving every block a node whose text and
   * attributes equal the old one's and whose markers are the new ones, the diff keeping what is
   * equal, one undo step taking the run back. Call it with the bytes of a save of this session,
   * so the answer carries every pending edit. The boundaries are labelled computed from then on.
   */
  adoptPaginated(paginated: Uint8Array): Promise<TreeRunResult<void>> {
    return this.serially(() => this.adoptPaginatedNow(paginated));
  }

  private async adoptPaginatedNow(paginated: Uint8Array): Promise<TreeRunResult<void>> {
    const pkg = await WordprocessingMLPackage.load(paginated);
    const theirs = await pkg.getMainDocumentPart().readContents();
    if (this.inRoom) return this.moveMarkersFrom(theirs.body ?? { TYPE_NAME: 'org_docx4j_wml.Body' });
    const before = this.paginationSource;
    const result = await this.runOnTreeNow(() => {
      const body = this.body;
      body.content = theirs.body?.content ?? [];
      if (theirs.body?.sectPr) body.sectPr = theirs.body.sectPr;
      linkParents(body, this.document);
      // Set inside the run: the re-import's transaction redraws the boundaries, and they are the companion's from that draw on.
      this.paginationSource = 'companion';
    }, LOCAL_ORIGIN);
    if (result.error) this.paginationSource = before;
    return result;
  }

  /**
   * Re-paginate in a room (ED-004 section 3.8): the answer's markers moved into the store, one
   * local transaction and one undo step, nothing else touched and no `sources` entry written. The
   * answer is projected in a registry of its own, so its ids never enter this session's. A
   * different number of paragraphs refuses (`PaginationMismatchError`) and changes nothing.
   */
  private moveMarkersFrom(body: wml.Body): TreeRunResult<void> {
    const { doc: answer } = projectBody(body, this.schema, new SourceIds());
    const move = moveMarkers(this.doc, answer);
    if (!move) return { changed: false, error: new PaginationMismatchError() };
    this.paginationSource = 'companion';
    if (move.moved === 0) return { changed: false };
    this.undoManager.stopCapturing();
    this.replace(move.doc);
    this.undoManager.stopCapturing();
    return { changed: true };
  }

  /** Where the page-break markers came from: Word's file, or the companion's layout (ED-001 Appendix D.1 labels the latter computed). */
  paginationSource: 'file' | 'companion' = 'file';

  /**
   * Reload saved bytes, re-project, and compare with `doc`; then the bound stories with their re-projections (ED-003 section 6.2).
   * A zip is read as a zip, anything else as the Flat OPC a save in that format writes (section 12.122), so the check reads the file
   * as written.
   */
  async selfCheck(bytes: Uint8Array, doc: PMNode = this.doc, noteIds: NoteIds = new Map(), commentIds: CommentIds = new Map()): Promise<SelfCheckResult> {
    const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
    const source = zip ? bytes : await WordprocessingMLPackage.load(new TextDecoder().decode(bytes));
    const reloaded = await EditorDocument.open(source, { schema: this.schema });
    const body = compareProjections(doc, reloaded.imported);
    if (!body.equal || !this.storiesCache) return body;
    return this.storiesCache.selfCheck(reloaded, doc, noteIds, commentIds);
  }
}

/** Select by XPath's options (ED-005 section 3.5). */
export interface SelectXPathOptions {
  /** A notes, header or footer part by name; the main document part when absent. */
  part?: string;
  /** Prefixes beyond the object model's. */
  namespaces?: Record<string, string>;
  /** The engine; the package's (`pkg.xpathEngine`) when absent. */
  engine?: XPathEngine;
}

/** A performance mark of the open path (ED-003 section 8.2's breakdown), where the platform has marks. */
function mark(name: string): void {
  if (typeof performance !== 'undefined' && typeof performance.mark === 'function') performance.mark(name);
}

/** The settings part's note properties as the numbering reads them. */
export interface NoteSettings {
  footnotePr?: unknown;
  endnotePr?: unknown;
}

/** Whether any paragraph of the document names the list in its own `numPr`. */
function namesList(doc: PMNode, numId: string): boolean {
  let found = false;
  doc.descendants((node) => {
    if (found) return false;
    if (node.type.name !== 'paragraph') return true;
    if ((node.attrs['numPr'] as { numId: string | null } | null)?.numId === numId) found = true;
    return false;
  });
  return found;
}

/**
 * The `controls` map after a projection: an entry per `w:sdt` row or cell group of the context,
 * entries of groups the projection no longer has removed (after a tree run, an editor-made
 * `new:` group has become an element with an id of its own). Call inside a transaction.
 */
function syncControls(store: EditStore, ctx: ProjectionContext): void {
  const wanted = new Set<string>();
  for (const [id, group] of ctx.groups) {
    if (group.kind !== 'sdt') continue;
    wanted.add(id);
    writeControl(store, id, { form: group.form, sdtPr: group.sdtPr });
  }
  for (const id of readControlEntries(store).keys()) if (!wanted.has(id)) writeControl(store, id, null);
}
