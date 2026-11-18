// SPDX-License-Identifier: GPL-3.0-or-later
//
// Part edits replicated (ED-004 section 3.3): what a session edits in place in a part other than
// the main one goes into the store, and what arrives there is applied to the session's own tree.
//
// - Styles and numbering by element, since two people plausibly edit them at once: `styles`
//   holds `style:<styleId>`, `docDefaults` and `latentStyles`, `numbering` holds `num:<numId>`
//   and `abstract:<abstractNumId>`, each the element's XML inside a one-child `w:styles` or
//   `w:numbering` (the object model has no standalone element for most of them).
// - Other parts whole in `parts`: `part:<name>` (settings), `customXml:<itemID>` (a custom XML
//   part's XML and its part name, so a peer that lacks the part makes it with the same itemID and
//   name through core-ts 0.2.0's `customXmlParts.add` options), and
//   `notes:<kind>` for a notes part created in the session.
// - External relationships of the main part (hyperlinks) in `relationships`, by `<part>#<id>`.
//
// What changed is found by comparing each element with what was last known (the base file's,
// read from the base bytes, or the last written or received), so a run of the content API that
// touched a style is caught like an edit through the Styles tab. Conflicts are last writer per
// key; there is no undo (ED-003 section 9 item 7). As for `sources`, this holds in and out of a
// room: a recovered session replays its own earlier part edits.
import {
  WordprocessingMLPackage, NumberingDefinitionsPart, linkParents, deepCopy, marshalString, unmarshalString,
  type Element, type XmlPart, FooterPart, HeaderPart,
} from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as Y from 'yjs';
import { treeEqual } from './sources.mjs';
import type { EditorDocument } from './session.mjs';
import type { NoteKind } from './stories.mjs';

type LatentStyles = NonNullable<wml.Styles['latentStyles']>;
type Num = NonNullable<wml.Numbering['num']>[number];
type AbstractNum = NonNullable<wml.Numbering['abstractNum']>[number];

/** The origin part replication writes under: not undone, like every part edit. */
export const PARTSYNC_ORIGIN = 'parts';

/** The relationship type of a hyperlink's target (ECMA-376 part 1, 17.16.22). */
const HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

interface KeyedEntries {
  /** The element objects of the part now, by key. */
  now: Map<string, unknown>;
  /** The element's XML for the store, by key. */
  xml: (key: string, value: unknown) => Promise<string>;
}

function styleEntries(styles: wml.Styles | undefined): Map<string, unknown> {
  const out = new Map<string, unknown>();
  if (!styles) return out;
  for (const s of styles.style ?? []) if (s.styleId) out.set(`style:${s.styleId}`, s);
  if (styles.docDefaults) out.set('docDefaults', styles.docDefaults);
  if (styles.latentStyles) out.set('latentStyles', styles.latentStyles);
  return out;
}

function numberingEntries(numbering: wml.Numbering | undefined): Map<string, unknown> {
  const out = new Map<string, unknown>();
  if (!numbering) return out;
  for (const n of numbering.num ?? []) if (n.numId !== undefined) out.set(`num:${n.numId}`, n);
  for (const a of numbering.abstractNum ?? []) if (a.abstractNumId !== undefined) out.set(`abstract:${a.abstractNumId}`, a);
  return out;
}

async function styleXml(key: string, value: unknown): Promise<string> {
  const wrapper: wml.Styles = key === 'docDefaults' ? { docDefaults: value as wml.DocDefaults } : key === 'latentStyles' ? { latentStyles: value as LatentStyles } : { style: [value as wml.Style] };
  return marshalString(el.styles(wrapper) as unknown as Element);
}

async function numberingXml(key: string, value: unknown): Promise<string> {
  const wrapper: wml.Numbering = key.startsWith('num:') ? { num: [value as Num] } : { abstractNum: [value as AbstractNum] };
  return marshalString(el.numbering(wrapper) as unknown as Element);
}

/** The base file's elements first in their order, then the rest by key. */
function ordered<T, K extends string | number>(items: T[], key: (item: T) => K, baseOrder: readonly K[]): T[] {
  const rank = new Map(baseOrder.map((k, i) => [k, i] as const));
  return [...items].sort((a, b) => {
    const ka = key(a);
    const kb = key(b);
    const ra = rank.get(ka);
    const rb = rank.get(kb);
    if (ra !== undefined && rb !== undefined) return ra - rb;
    if (ra !== undefined) return -1;
    if (rb !== undefined) return 1;
    return ka < kb ? -1 : ka > kb ? 1 : 0;
  });
}

/** Whether a sort left every element where it was (an untouched part is then not rewritten). */
function sameOrder<T>(before: readonly T[] | undefined, after: readonly T[] | undefined): boolean {
  return before === after || (!!before && !!after && before.every((item, i) => item === after[i]));
}

export class PartSync {
  /** What each key was last known as (the base file's, or written or received since), by map and key. */
  private readonly known = new Map<string, unknown>();
  private base: Promise<WordprocessingMLPackage | undefined> | undefined;
  private pending: Promise<void> = Promise.resolve();
  private readonly listeners = new Set<() => void>();
  /** Received entries that could not be applied, by map and key: the finding. */
  readonly problems = new Map<string, string>();

  /**
   * `baseBytes` is the file the session opened, which the styles and numbering entries are
   * compared against. A session opened from a package (tests, a package made in memory) has none:
   * its keyed publications wait until `adoptBase` gives one (Share does), since comparing with
   * nothing would publish every style and definition as changed.
   */
  constructor(private readonly session: EditorDocument, private baseBytes: Uint8Array | undefined) {}

  /** The base file, for a session opened without its bytes; then what changed since it is published. */
  adoptBase(bytes: Uint8Array): void {
    if (this.baseBytes) return;
    this.baseBytes = bytes;
    this.base = undefined;
    this.publishAll();
  }

  /** Resolves once every queued publication and application is done. */
  ready(): Promise<void> {
    return this.pending;
  }

  /** Called after a received part edit is applied (the application refreshes the stylesheet, the labels, the panels). */
  onChange(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private queue(task: () => Promise<void>): void {
    this.pending = this.pending.then(task).catch((e: unknown) => { this.problems.set('queue', (e as Error).message); });
  }

  // ---------------------------------------------------------------- publishing

  /** A part was edited in place (`markPartChanged`): what changed in it, written to the store. */
  published(partName: string): void {
    this.queue(() => this.publishPart(partName));
  }

  /** After a tree run: the parts a run of the content API can change (styles, numbering, the main part's hyperlinks). */
  publishAll(): void {
    this.queue(async () => {
      await this.publishStyles();
      await this.publishNumbering();
      this.publishRelationships();
    });
  }

  private async publishPart(partName: string): Promise<void> {
    const main = this.session.main;
    if (main.styleDefinitionsPart?.partName.toString() === partName) return this.publishStyles();
    if (main.numberingDefinitionsPart?.partName.toString() === partName || partName === '/word/numbering.xml') return this.publishNumbering();
    if (main.footnotesPart?.partName.toString() === partName) return this.publishNotesPart('footnote');
    if (main.endnotesPart?.partName.toString() === partName) return this.publishNotesPart('endnote');
    const custom = this.session.pkg.customXmlParts.items.find((p) => p.part.partName.toString() === partName);
    if (custom) return this.publishCustomXml(custom.id, custom.getXml(), partName);
    if (partName.startsWith('/customXml/')) return this.publishCustomXmlRemovals();
    const part = this.session.pkg.parts.get(partName) as XmlPart<unknown> | undefined;
    if (part && typeof (part as { getXml?: unknown }).getXml === 'function') {
      const xml = await part.getXml();
      // A header or footer part made here (ED-005 section 12.95) carries what a peer that lacks it needs to make the same: its kind and its relationship id.
      const kind = part instanceof FooterPart ? 'footer' : part instanceof HeaderPart ? 'header' : undefined;
      const relId = kind ? part.sourceRelationship?.id : undefined;
      this.write(this.session.store.parts, `part:${partName}`, kind && relId ? { contentType: part.contentType, xml, kind, relId } : { contentType: part.contentType, xml });
    }
  }

  private async basePart<T>(pick: (pkg: WordprocessingMLPackage) => XmlPart<T> | undefined): Promise<T | undefined> {
    // The base file read as the session read it, through Word's repair where it took one.
    this.base ??= this.baseBytes ? WordprocessingMLPackage.load(this.baseBytes, { repair: this.session.pkg.loadOptions.repair ?? 'refuse' }).catch(() => undefined) : Promise.resolve(undefined);
    const pkg = await this.base;
    const part = pkg ? pick(pkg) : undefined;
    return part ? part.readContents() : undefined;
  }

  /** Seed what the base file's part holds under a map's prefix, once, so the first comparison is against the base. */
  private async seed(prefix: string, base: () => Promise<Map<string, unknown>>): Promise<void> {
    if (this.known.has(`${prefix}#seeded`)) return;
    this.known.set(`${prefix}#seeded`, true);
    for (const [k, v] of await base()) if (!this.known.has(`${prefix}|${k}`)) this.known.set(`${prefix}|${k}`, deepCopy(v as object));
  }

  private async publishKeyed(prefix: 'styles' | 'numbering', map: Y.Map<unknown>, entries: KeyedEntries): Promise<void> {
    const written: [string, string | null][] = [];
    for (const [key, value] of entries.now) {
      const known = this.known.get(`${prefix}|${key}`);
      if (known !== undefined && treeEqual(value, known)) continue;
      written.push([key, await entries.xml(key, value)]);
      this.known.set(`${prefix}|${key}`, deepCopy(value as object));
    }
    for (const k of [...this.known.keys()]) {
      if (!k.startsWith(`${prefix}|`)) continue;
      const key = k.slice(prefix.length + 1);
      if (!entries.now.has(key)) { written.push([key, null]); this.known.delete(k); }
    }
    if (written.length === 0) return;
    this.session.store.ydoc.transact(() => {
      for (const [key, xml] of written) { if (xml === null) map.delete(key); else map.set(key, xml); }
    }, PARTSYNC_ORIGIN);
  }

  private async publishStyles(): Promise<void> {
    const part = this.session.main.styleDefinitionsPart;
    if (!part || !this.baseBytes) return;
    await this.seed('styles', async () => styleEntries(await this.basePart((p) => p.getMainDocumentPart().styleDefinitionsPart)));
    await this.publishKeyed('styles', this.session.store.styles, { now: styleEntries(await part.readContents()), xml: styleXml });
  }

  private async publishNumbering(): Promise<void> {
    const part = this.session.main.numberingDefinitionsPart;
    if (!part || !this.baseBytes) return;
    await this.seed('numbering', async () => numberingEntries(await this.basePart((p) => p.getMainDocumentPart().numberingDefinitionsPart)));
    await this.publishKeyed('numbering', this.session.store.numbering, { now: numberingEntries(await part.readContents()), xml: numberingXml });
  }

  private async publishNotesPart(kind: NoteKind): Promise<void> {
    // The notes' content travels in their fragments; the part's creation is what a peer needs.
    const had = await this.basePart((p) => (kind === 'footnote' ? p.getMainDocumentPart().footnotesPart : p.getMainDocumentPart().endnotesPart) as XmlPart<unknown> | undefined);
    if (had === undefined) this.write(this.session.store.parts, `notes:${kind}`, { create: kind });
  }

  private async publishCustomXml(itemID: string, xml: string, partName: string): Promise<void> {
    this.write(this.session.store.parts, `customXml:${itemID.toUpperCase()}`, { xml, partName });
  }

  /** A custom XML part removed (the last macro deleted): its key goes. */
  private async publishCustomXmlRemovals(): Promise<void> {
    const present = new Set(this.session.pkg.customXmlParts.items.map((p) => `customXml:${p.id.toUpperCase()}`));
    const gone = [...this.session.store.parts.keys()].filter((k) => k.startsWith('customXml:') && !present.has(k));
    if (gone.length === 0) return;
    this.session.store.ydoc.transact(() => { for (const k of gone) this.session.store.parts.delete(k); }, PARTSYNC_ORIGIN);
  }

  /** The main part's external relationships (hyperlinks) not yet in the store. */
  publishRelationships(): void {
    const rels = this.session.main.relationshipsPart;
    if (!rels) return;
    const source = this.session.main.partName.toString();
    const add: [string, unknown][] = [];
    for (const rel of rels.list) {
      if (rel.type !== HYPERLINK || rel.targetMode !== 'External') continue;
      const key = `${source}#${rel.id}`;
      if (this.session.store.relationships.has(key)) continue;
      add.push([key, { type: rel.type, target: rel.target, mode: 'External' }]);
    }
    if (add.length === 0) return;
    this.session.store.ydoc.transact(() => { for (const [k, v] of add) this.session.store.relationships.set(k, v); }, PARTSYNC_ORIGIN);
  }

  private write(map: Y.Map<unknown>, key: string, value: unknown): void {
    this.session.store.ydoc.transact(() => map.set(key, value), PARTSYNC_ORIGIN);
  }

  // ---------------------------------------------------------------- order (in a room)

  /**
   * In a room, every peer writes a part's elements in the same order (ED-004 section 3.3): the
   * base file's in the base file's order, then the ones made since in order of key (the random
   * ids of section 3.4), since each peer appended its own first and the others' as they arrived.
   * Outside a room the order is the session's own, as before. Called before a save.
   */
  async normalizeOrder(): Promise<void> {
    if (!this.session.inRoom) return;
    const numbering = this.session.main.numberingDefinitionsPart;
    if (numbering?.isUnmarshalled) {
      const tree = await numbering.getContents();
      const base = await this.basePart((p) => p.getMainDocumentPart().numberingDefinitionsPart);
      const baseNums = (base?.num ?? []).map((n) => Number(n.numId));
      const baseAbstracts = (base?.abstractNum ?? []).map((a) => Number(a.abstractNumId));
      const nums = tree.num && ordered(tree.num, (n) => Number(n.numId), baseNums);
      const abstracts = tree.abstractNum && ordered(tree.abstractNum, (a) => Number(a.abstractNumId), baseAbstracts);
      if (!sameOrder(tree.num, nums) || !sameOrder(tree.abstractNum, abstracts)) {
        tree.num = nums;
        tree.abstractNum = abstracts;
        linkParents(tree, undefined);
        numbering.setContents(tree);
      }
    }
    const styles = this.session.main.styleDefinitionsPart;
    if (styles?.isUnmarshalled) {
      const tree = await styles.getContents();
      const base = await this.basePart((p) => p.getMainDocumentPart().styleDefinitionsPart);
      const sorted = tree.style && ordered(tree.style, (s) => s.styleId ?? '', (base?.style ?? []).map((s) => s.styleId ?? ''));
      if (!sameOrder(tree.style, sorted)) {
        tree.style = sorted;
        linkParents(tree, undefined);
        styles.setContents(tree);
      }
    }
  }

  // ---------------------------------------------------------------- applying

  /** Every entry the store holds, applied at open (a join, a recovery), then each one that arrives. */
  async replay(): Promise<void> {
    const store = this.session.store;
    const any = store.styles.size + store.numbering.size + store.parts.size + store.relationships.size > 0;
    if (!any) { this.observe(); return; }
    for (const key of store.parts.keys()) await this.applyPart(key, store.parts.get(key));
    for (const [key, xml] of store.styles.entries()) await this.applyStyle(key, xml as string);
    for (const [key, xml] of store.numbering.entries()) await this.applyNumbering(key, xml as string);
    for (const key of store.relationships.keys()) this.applyRelationship(key, store.relationships.get(key));
    await this.refreshed(true, true);
    this.observe();
  }

  private observe(): void {
    const store = this.session.store;
    const on = (map: Y.Map<unknown>, apply: (key: string, value: unknown) => Promise<void> | void, styles: boolean, numbering: boolean): void => {
      map.observe((event) => {
        if (event.transaction.local) return;
        const keys = [...event.changes.keys.keys()];
        this.queue(async () => {
          // The room's value when the apply runs, not when the update arrived: a publish of this
          // peer's own that was queued first may have written the same key since, and Yjs's winner
          // is what the map holds (ED-005 section 12.37: two pages applying one style each kept the
          // other's).
          for (const key of keys) await apply(key, map.has(key) ? map.get(key) : null);
          await this.refreshed(styles, numbering);
        });
      });
    };
    on(store.styles, (k, v) => this.applyStyle(k, v as string | null), true, false);
    on(store.numbering, (k, v) => this.applyNumbering(k, v as string | null), false, true);
    on(store.parts, (k, v) => this.applyPart(k, v), false, false);
    on(store.relationships, (k, v) => this.applyRelationship(k, v), false, false);
  }

  private async refreshed(styles: boolean, numbering: boolean): Promise<void> {
    if (numbering) await this.session.pkg.refresh();
    else if (styles) this.session.pkg.refreshPropertyResolver();
    for (const l of [...this.listeners]) l();
  }

  private async applyStyle(key: string, xml: string | null): Promise<void> {
    const part = this.session.main.styleDefinitionsPart;
    if (!part) { this.problems.set(`styles|${key}`, 'the document has no styles part'); return; }
    const tree = await part.getContents();
    let value: unknown;
    if (xml !== null) {
      const parsed = (await unmarshalString(xml)).value as wml.Styles;
      value = key === 'docDefaults' ? parsed.docDefaults : key === 'latentStyles' ? parsed.latentStyles : parsed.style?.[0];
      if (value === undefined) { this.problems.set(`styles|${key}`, 'the entry holds no element'); return; }
    }
    if (key === 'docDefaults') { if (value) tree.docDefaults = value as wml.DocDefaults; else delete tree.docDefaults; }
    else if (key === 'latentStyles') { if (value) tree.latentStyles = value as LatentStyles; else delete tree.latentStyles; }
    else {
      const id = key.slice('style:'.length);
      const list = (tree.style ??= []);
      const at = list.findIndex((s) => s.styleId === id);
      if (value === undefined) { if (at >= 0) list.splice(at, 1); }
      else if (at >= 0) list[at] = value as wml.Style;
      else list.push(value as wml.Style);
    }
    linkParents(tree, undefined);
    part.setContents(tree);
    this.touched(part.partName.toString(), `styles|${key}`, value);
  }

  private async applyNumbering(key: string, xml: string | null): Promise<void> {
    let part = this.session.main.numberingDefinitionsPart;
    if (!part) {
      // The first list a peer made in a document that had none: the part is made here as the engine makes it.
      part = new NumberingDefinitionsPart();
      this.session.main.addTargetPart(part);
      part.setContents({});
    }
    const tree = await part.getContents();
    let value: unknown;
    if (xml !== null) {
      const parsed = (await unmarshalString(xml)).value as wml.Numbering;
      value = key.startsWith('num:') ? parsed.num?.[0] : parsed.abstractNum?.[0];
      if (value === undefined) { this.problems.set(`numbering|${key}`, 'the entry holds no element'); return; }
    }
    if (key.startsWith('num:')) {
      const id = Number(key.slice(4));
      const list = (tree.num ??= []);
      const at = list.findIndex((n) => Number(n.numId) === id);
      if (value === undefined) { if (at >= 0) list.splice(at, 1); } else if (at >= 0) list[at] = value as Num; else list.push(value as Num);
    } else {
      const id = Number(key.slice(9));
      const list = (tree.abstractNum ??= []);
      const at = list.findIndex((a) => Number(a.abstractNumId) === id);
      // An abstract definition goes before every w:num, as the schema orders them; appending keeps that.
      if (value === undefined) { if (at >= 0) list.splice(at, 1); } else if (at >= 0) list[at] = value as AbstractNum; else list.push(value as AbstractNum);
    }
    linkParents(tree, undefined);
    part.setContents(tree);
    this.touched(part.partName.toString(), `numbering|${key}`, value);
  }

  private async applyPart(key: string, value: unknown): Promise<void> {
    const pkg = this.session.pkg;
    if (key.startsWith('notes:')) {
      if (value) await this.session.ensureNotesPart(key.slice(6) as NoteKind, false);
      return;
    }
    if (key.startsWith('customXml:')) {
      await pkg.customXmlParts.load();
      const id = key.slice('customXml:'.length);
      const view = pkg.customXmlParts.items.find((p) => p.id.toUpperCase() === id);
      if (value === null || value === undefined) {
        if (view) { const name = view.part.partName.toString(); pkg.customXmlParts.removePart(view); this.touched(name); }
        return;
      }
      const entry = value as { xml: string; partName?: string };
      if (!view) {
        // A part a peer made in the room: made here with the same itemID and part name, so every
        // peer's package, and so its save, has the same part.
        if (entry.partName === undefined) { this.problems.set(`parts|${key}`, 'a custom XML part this peer lacks, published without its part name'); return; }
        const made = pkg.customXmlParts.add(entry.xml, { itemID: id, partName: entry.partName });
        this.touched(made.part.partName.toString());
        this.session.template.invalidate();
        return;
      }
      view.setXml(entry.xml);
      this.touched(view.part.partName.toString());
      this.session.template.invalidate();
      return;
    }
    if (key.startsWith('part:')) {
      const name = key.slice(5);
      let part = pkg.parts.get(name) as XmlPart<unknown> | undefined;
      const made = value as { xml: string; kind?: 'header' | 'footer'; relId?: string } | null | undefined;
      if (!part && made?.kind && made.relId && !this.session.main.getRelationshipsPart(true)!.getRelationshipById(made.relId)) {
        // A header or footer a peer made (ED-005 section 12.95): made here with the same name and relationship id, so the peer's section reference resolves.
        const hf = made.kind === 'footer' ? new FooterPart(name) : new HeaderPart(name);
        this.session.main.addTargetPart(hf, undefined, made.relId);
        hf.setXml(made.xml);
        const tree = await hf.getContents() as wml.Hdr | wml.Ftr;
        this.session.stories.headerAdded(made.kind, hf, made.relId, tree);
        this.touched(name);
        return;
      }
      if (!part || value === null || value === undefined) { this.problems.set(`parts|${key}`, 'a part this peer lacks'); return; }
      part.setXml((value as { xml: string }).xml);
      // A typed part is unmarshalled now, so a bad entry is a problem here and not at the save; an untyped one (a `DefaultXmlPart`) is kept as XML.
      if (typeof (part as { getContents?: unknown }).getContents === 'function') await part.getContents();
      this.touched(name);
      if (name === this.session.main.documentSettingsPart?.partName.toString()) await this.session.readSettings();
    }
  }

  private applyRelationship(key: string, value: unknown): void {
    if (!value) return;
    const [source, id] = [key.slice(0, key.lastIndexOf('#')), key.slice(key.lastIndexOf('#') + 1)];
    if (source !== this.session.main.partName.toString()) { this.problems.set(`relationships|${key}`, 'a relationship of a part other than the main one'); return; }
    const rels = this.session.main.getRelationshipsPart(true)!;
    if (rels.getRelationshipById(id)) return;
    const v = value as { type: string; target: string };
    rels.addExternalRelationship(v.type, v.target, id);
  }

  private touched(partName: string, knownKey?: string, value?: unknown): void {
    this.session.partReceived(partName);
    if (knownKey !== undefined) {
      if (value === undefined) this.known.delete(knownKey); else this.known.set(knownKey, deepCopy(value as object));
    }
  }
}
