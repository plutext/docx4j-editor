// SPDX-License-Identifier: GPL-3.0-or-later
//
// A content control's properties as the plain JSON of its `w:sdtPr` (ED-003 section 4.2): the
// shape `toPlain` gives an unmarshalled one and the objects package's `sdtPr` builder gives a new
// one, so a node's `sdtPr` attribute, a `controls` map entry and the tree agree. Every function
// here is pure and returns a new object; nothing mutates an attribute in place (a ProseMirror
// attribute is shared by every state that holds the node).
//
// The Office JS vocabulary is `ContentControl`'s in core-ts (tag, title, appearance, color,
// cannotDelete, cannotEdit, removeWhenEdited, placeholderText); the shapes written are the ones
// the engine's builders write, checked against them in the tests.
import { sdtPr as buildSdtPr, sdtKindOf, type SdtKind } from '@docx4j/generated-objects-ts/builders/wml';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { placeholderEntryOf as enginePlaceholderEntryOf } from '@docx4j/core-ts';
import { revive, toPlain } from './json.mjs';

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** One child of `w:sdtPr` as plain JSON. */
export interface SdtPrItem {
  name: { namespaceURI: string; localPart: string };
  value: Record<string, unknown>;
}

export interface SdtPrJson {
  TYPE_NAME?: string;
  rPrOrAliasOrLock?: SdtPrItem[];
}

export type { SdtKind };

export type ContentControlAppearance = 'BoundingBox' | 'Tags' | 'Hidden';

/** Office JS `Word.ContentControl`'s settable properties, as one patch. */
export interface ControlProperties {
  tag?: string;
  title?: string;
  appearance?: ContentControlAppearance;
  /** '#RRGGBB', or '' for none. */
  color?: string;
  cannotDelete?: boolean;
  cannotEdit?: boolean;
  removeWhenEdited?: boolean;
  /** `w:showingPlcHdr`: the control shows its placeholder rather than a value. */
  showingPlaceholder?: boolean;
  /** `w:placeholder/w:docPart`: the glossary entry the placeholder is, in Word's form (ED-008 section 8.8 item 14); null removes it. */
  placeholderEntry?: string | null;
  /** `w:text/@w:multiLine` of a plain-text control. */
  multiLine?: boolean;
}

/** What the panel shows: every property read from the JSON. */
export interface ControlReading {
  kind: SdtKind;
  tag: string;
  title: string;
  id: number | undefined;
  appearance: ContentControlAppearance;
  color: string;
  cannotDelete: boolean;
  cannotEdit: boolean;
  removeWhenEdited: boolean;
  showingPlaceholder: boolean;
  /** The glossary entry `w:placeholder/w:docPart` names; '' for a `w:placeholder` naming none, undefined for the bare form. */
  placeholderEntry: string | undefined;
  multiLine: boolean;
  binding: DataBinding | undefined;
}

export interface DataBinding {
  xpath: string;
  storeItemID: string;
  prefixMappings: string;
  /** `w15:dataBinding` (a repeating section, a container-bound rich text control) rather than `w:dataBinding`. */
  w15: boolean;
}

export function sdtItems(sdtPr: unknown): SdtPrItem[] {
  return ((sdtPr as SdtPrJson | null | undefined)?.rPrOrAliasOrLock ?? []) as SdtPrItem[];
}

/** The item named `localPart` in `ns` (wml by default; an array for several; '*' for any). */
export function sdtItem(sdtPr: unknown, localPart: string, ns: string | readonly string[] = W_NS): SdtPrItem | undefined {
  const wanted = ns === '*' ? undefined : new Set(typeof ns === 'string' ? [ns] : ns);
  return sdtItems(sdtPr).find((i) => i.name.localPart === localPart && (wanted === undefined || wanted.has(i.name.namespaceURI)));
}

/** The JSON with `item` in place of the one of the same name and namespace, or appended. */
export function withSdtItem(sdtPr: unknown, item: SdtPrItem): SdtPrJson {
  const base = (sdtPr as SdtPrJson | null | undefined) ?? { TYPE_NAME: 'org_docx4j_wml.SdtPr' };
  const items = sdtItems(base);
  const at = items.findIndex((i) => i.name.localPart === item.name.localPart && i.name.namespaceURI === item.name.namespaceURI);
  const next = at >= 0 ? items.map((i, k) => (k === at ? item : i)) : [...items, item];
  return { ...base, TYPE_NAME: base.TYPE_NAME ?? 'org_docx4j_wml.SdtPr', rPrOrAliasOrLock: next };
}

/** The JSON without the item named; unchanged (same object) when there is none. */
export function withoutSdtItem(sdtPr: unknown, localPart: string, ns: string | readonly string[] = W_NS): SdtPrJson | null {
  const base = (sdtPr as SdtPrJson | null | undefined) ?? null;
  if (!base) return null;
  const wanted = ns === '*' ? undefined : new Set(typeof ns === 'string' ? [ns] : ns);
  const items = sdtItems(base);
  const next = items.filter((i) => !(i.name.localPart === localPart && (wanted === undefined || wanted.has(i.name.namespaceURI))));
  if (next.length === items.length) return base;
  return { ...base, rPrOrAliasOrLock: next };
}

function item(localPart: string, value: Record<string, unknown>, ns = W_NS): SdtPrItem {
  return { name: { namespaceURI: ns, localPart }, value };
}

const stringVal = (sdtPr: unknown, localPart: string): string => {
  const v = sdtItem(sdtPr, localPart)?.value['val'];
  return typeof v === 'string' ? v : '';
};

const onOff = (sdtPr: unknown, localPart: string): boolean => {
  const i = sdtItem(sdtPr, localPart);
  return i !== undefined && i.value['val'] !== false;
};

export function sdtTag(sdtPr: unknown): string {
  return stringVal(sdtPr, 'tag');
}

export function sdtTitle(sdtPr: unknown): string {
  return stringVal(sdtPr, 'alias');
}

export function sdtId(sdtPr: unknown): number | undefined {
  const v = sdtItem(sdtPr, 'id')?.value['val'];
  return typeof v === 'number' ? v : undefined;
}

export function sdtKind(sdtPr: unknown): SdtKind {
  return sdtKindOf((sdtPr as wml.SdtPr | null | undefined) ?? undefined);
}

export function sdtBinding(sdtPr: unknown): DataBinding | undefined {
  const w = sdtItem(sdtPr, 'dataBinding');
  const found = w ?? sdtItem(sdtPr, 'dataBinding', W15_NS);
  if (!found) return undefined;
  const v = found.value as { xpath?: string; storeItemID?: string; prefixMappings?: string };
  return { xpath: v.xpath ?? '', storeItemID: v.storeItemID ?? '', prefixMappings: v.prefixMappings ?? '', w15: w === undefined };
}

/** `w:sdtPr/w:rPr` as plain JSON: what a bound value is written with (REQ-021). */
export function sdtRunProperties(sdtPr: unknown): unknown | undefined {
  return sdtItem(sdtPr, 'rPr')?.value;
}

export function readControl(sdtPr: unknown): ControlReading {
  const lock = stringVal(sdtPr, 'lock');
  const appearance = sdtItem(sdtPr, 'appearance', W15_NS)?.value['val'];
  const color = sdtItem(sdtPr, 'color', W15_NS)?.value['val'];
  const text = sdtItem(sdtPr, 'text')?.value as { multiLine?: boolean } | undefined;
  return {
    kind: sdtKind(sdtPr),
    tag: sdtTag(sdtPr),
    title: sdtTitle(sdtPr),
    id: sdtId(sdtPr),
    appearance: appearance === 'tags' ? 'Tags' : appearance === 'hidden' ? 'Hidden' : 'BoundingBox',
    color: typeof color === 'string' && color !== 'auto' ? `#${color.replace(/^#/, '')}` : '',
    cannotDelete: lock === 'sdtLocked' || lock === 'sdtContentLocked',
    cannotEdit: lock === 'contentLocked' || lock === 'sdtContentLocked',
    removeWhenEdited: onOff(sdtPr, 'temporary'),
    showingPlaceholder: onOff(sdtPr, 'showingPlcHdr'),
    placeholderEntry: placeholderEntryOf(sdtPr),
    multiLine: text?.multiLine === true,
    binding: sdtBinding(sdtPr),
  };
}

/** The JSON with a patch of properties applied, in the shapes the engine's builders write. */
export function applyControlProperties(sdtPr: unknown, patch: ControlProperties): SdtPrJson {
  let out: SdtPrJson | null = (sdtPr as SdtPrJson | null | undefined) ?? { TYPE_NAME: 'org_docx4j_wml.SdtPr' };
  const setString = (localPart: string, typeName: string, value: string | undefined): void => {
    if (value === undefined) return;
    out = value === '' ? withoutSdtItem(out, localPart) : withSdtItem(out, item(localPart, { val: value, TYPE_NAME: typeName }));
  };
  setString('tag', 'org_docx4j_wml.Tag', patch.tag);
  setString('alias', 'org_docx4j_wml.SdtPr.Alias', patch.title);
  if (patch.appearance !== undefined) {
    out = patch.appearance === 'BoundingBox'
      ? withoutSdtItem(out, 'appearance', W15_NS)
      : withSdtItem(out, item('appearance', { val: patch.appearance === 'Tags' ? 'tags' : 'hidden', TYPE_NAME: 'org_docx4j_w15.CTSdtAppearance' }, W15_NS));
  }
  if (patch.color !== undefined) {
    out = patch.color === ''
      ? withoutSdtItem(out, 'color', W15_NS)
      : withSdtItem(out, item('color', { val: patch.color.replace(/^#/, '').toUpperCase(), TYPE_NAME: 'org_docx4j_wml.CTColor' }, W15_NS));
  }
  if (patch.cannotDelete !== undefined || patch.cannotEdit !== undefined) {
    const current = readControl(out);
    const cannotDelete = patch.cannotDelete ?? current.cannotDelete;
    const cannotEdit = patch.cannotEdit ?? current.cannotEdit;
    const val = cannotDelete && cannotEdit ? 'sdtContentLocked' : cannotDelete ? 'sdtLocked' : cannotEdit ? 'contentLocked' : undefined;
    out = val === undefined ? withoutSdtItem(out, 'lock') : withSdtItem(out, item('lock', { val, TYPE_NAME: 'org_docx4j_wml.CTLock' }));
  }
  // A flag in Word's form, the empty element (`<w:showingPlcHdr/>`, `<w:temporary/>`), as the engine writes it since 0.6.0.
  const setFlag = (localPart: string, value: boolean | undefined): void => {
    if (value === undefined) return;
    out = value ? withSdtItem(out, item(localPart, { TYPE_NAME: 'org_docx4j_wml.BooleanDefaultTrue' })) : withoutSdtItem(out, localPart);
  };
  setFlag('temporary', patch.removeWhenEdited);
  setFlag('showingPlcHdr', patch.showingPlaceholder);
  if (patch.placeholderEntry !== undefined) out = patch.placeholderEntry === null ? withoutSdtItem(out, 'placeholder') : withPlaceholderEntry(out, patch.placeholderEntry);
  if (patch.multiLine !== undefined) {
    const text = sdtItem(out, 'text');
    if (text) {
      const value = { ...text.value };
      if (patch.multiLine) value['multiLine'] = true; else delete value['multiLine'];
      out = withSdtItem(out, item('text', value));
    }
  }
  return out ?? { TYPE_NAME: 'org_docx4j_wml.SdtPr' };
}

/** The glossary entry a control's `w:placeholder` names; undefined where it names none, the bare form (the engine's, since core-ts 0.7.0). */
export function placeholderEntryOf(sdtPr: unknown): string | undefined {
  return enginePlaceholderEntryOf(sdtPr as wml.SdtPr | undefined);
}

/**
 * The JSON with `w:placeholder` naming `entry`, where Word writes it: after `w:id`, before `w:showingPlcHdr`, the
 * binding and `w:temporary` (check 64's controls as Word saved them). The engine's `namePlaceholderEntry` places it
 * the same way but in place; a control's properties here are a node's attribute and a store's value, never mutated.
 */
function withPlaceholderEntry(sdtPr: unknown, entry: string): SdtPrJson {
  const element = item('placeholder', { TYPE_NAME: 'org_docx4j_wml.CTPlaceholder', docPart: { TYPE_NAME: 'org_docx4j_wml.CTPlaceholder.DocPart', val: entry } });
  const base = (sdtPr as SdtPrJson | null | undefined) ?? { TYPE_NAME: 'org_docx4j_wml.SdtPr' };
  const items = sdtItems(base);
  if (items.some((i) => i.name.localPart === 'placeholder')) return withSdtItem(base, element);
  const after = new Set(['showingPlcHdr', 'dataBinding', 'temporary']);
  const before = items.findIndex((i) => after.has(i.name.localPart));
  const next = [...items];
  next.splice(before < 0 ? next.length : before, 0, element);
  return { ...base, TYPE_NAME: base.TYPE_NAME ?? 'org_docx4j_wml.SdtPr', rPrOrAliasOrLock: next };
}

/** The JSON with the binding written (`w:dataBinding`, or `w15:dataBinding` when asked) or, with null, removed. */
export function withBinding(sdtPr: unknown, binding: { xpath: string; storeItemID: string; prefixMappings?: string } | null, w15 = false): SdtPrJson | null {
  if (binding === null) return withoutSdtItem(withoutSdtItem(sdtPr, 'dataBinding') , 'dataBinding', W15_NS);
  const value: Record<string, unknown> = { xpath: binding.xpath, storeItemID: binding.storeItemID, TYPE_NAME: 'org_docx4j_wml.CTDataBinding' };
  if (binding.prefixMappings) value['prefixMappings'] = binding.prefixMappings;
  const cleared = withoutSdtItem(withoutSdtItem(sdtPr, 'dataBinding'), 'dataBinding', W15_NS);
  return withSdtItem(cleared, item('dataBinding', value, w15 ? W15_NS : W_NS));
}

/**
 * How many control properties holding a `w15` element (a repeating section, `w15:dataBinding`,
 * `w15:appearance`, `w15:color`) the export has written: an export that wrote one declares `w15`
 * ignorable on its part's root, since Word 2010 knows no `w15` and opens such a part only when the
 * prefix is ignorable, as Word 15's own files declare it (ED-005 section 12.58).
 */
let w15Written = 0;
export function w15Writes(): number {
  return w15Written;
}

/** A control's properties as the export writes them, counted when they hold a `w15` element (`w15Writes`). */
export function exportedSdtPr(sdtPr: unknown): wml.SdtPr {
  if ((sdtPr as SdtPrJson | null)?.rPrOrAliasOrLock?.some((i) => i.name?.namespaceURI === W15_NS)) w15Written++;
  return revive(sdtPr) as wml.SdtPr;
}

/** The JSON with `w:id` set. */
export function withSdtId(sdtPr: unknown, id: number): SdtPrJson {
  return withSdtItem(sdtPr, item('id', { val: id, TYPE_NAME: 'org_docx4j_wml.Id' }));
}

/** A new control's properties, through the objects package's builder (alias, tag, id, the kind element, in that order). */
export function newSdtPr(kind: SdtKind, options: { tag?: string; title?: string; id: number }): SdtPrJson {
  const built = buildSdtPr({ kind: kind === 'RichText' ? undefined : kind, id: options.id, ...(options.tag !== undefined ? { tag: options.tag } : {}), ...(options.title !== undefined ? { title: options.title } : {}) });
  return toPlain(built) as SdtPrJson;
}

/** A random 31-bit `w:id`, as Word's are, not among `used`. */
export function allocateSdtId(used: ReadonlySet<number>): number {
  let id = 1 + Math.floor(Math.random() * 0x3fffffff);
  while (used.has(id)) id = 1 + Math.floor(Math.random() * 0x3fffffff);
  return id;
}
