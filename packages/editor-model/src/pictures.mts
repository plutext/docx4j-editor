// SPDX-License-Identifier: GPL-3.0-or-later
//
// Picture bindings (ED-005 proposal 45, section 12.76; the specification's section 9.3). Two shapes
// bind a control to a node holding base64 image data: Word's picture content control (`w:picture`
// with a `w:dataBinding`), and OpenDoPE's `od:Handler=picture` on a rich text control, bound through
// the tag's `od:xpath` entry as an XHTML control is, with an optional `width` parameter (`auto`, or
// twentieths of a point). The view shows them (`view/picturebind.ts`); the engine's bind step fills
// them for a processor (core-ts CR-005 section 9).
import type { Node as PMNode } from 'prosemirror-model';
import { imageInfoOf } from '@docx4j/core-ts';
import { readTag, parseTag, formatTag, withTagParam } from './opendope.mjs';
import { sdtTag, sdtKind, sdtBinding, type DataBinding } from './sdtpr.mjs';
import type { TemplateModel } from './template.mjs';

export interface PictureBinding {
  binding: DataBinding;
  /** Word's own picture control, or OpenDoPE's handler on a rich text control. */
  shape: 'picture' | 'handler';
  /** The handler's `width`: `auto`, points, or undefined (REQ-061: the drawing kept, its blip replaced). */
  width?: number | 'auto';
}

/** Whether a control's tag names OpenDoPE's picture handler. */
export function isPictureHandler(sdtPr: unknown): boolean {
  return readTag(sdtTag(sdtPr)).handler === 'picture';
}

/** The picture binding of a control, or undefined: Word's picture control with its `w:dataBinding`, or the handler's `od:xpath` entry (the template must be loaded). */
export function pictureBindingOf(sdtPr: unknown, template: TemplateModel): PictureBinding | undefined {
  const tag = readTag(sdtTag(sdtPr));
  if (tag.handler === 'picture') {
    const id = tag.xpath;
    const entry = id !== undefined ? template.xpaths?.entries.find((e) => e.id === id) : undefined;
    const direct = sdtBinding(sdtPr);
    const binding: DataBinding | undefined = entry ? { xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings, w15: false } : direct;
    if (!binding || binding.xpath === '') return undefined;
    const raw = tag.params.find((p) => p.key === 'width')?.value;
    const width = raw === undefined ? undefined : raw === 'auto' ? 'auto' : Number.isFinite(Number(raw)) && Number(raw) > 0 ? Number(raw) / 20 : undefined;
    return width === undefined && raw !== undefined && raw !== 'auto' ? { binding, shape: 'handler' } : { binding, shape: 'handler', ...(width !== undefined ? { width } : {}) };
  }
  if (sdtKind(sdtPr) !== 'Picture') return undefined;
  const binding = sdtBinding(sdtPr);
  return binding && binding.xpath !== '' ? { binding, shape: 'picture' } : undefined;
}

export interface PictureControl {
  pos: number;
  node: PMNode;
  picture: PictureBinding;
}

/** Every picture-bound control of the body, in document order. */
export function pictureControlsIn(doc: PMNode, template: TemplateModel): PictureControl[] {
  const out: PictureControl[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'sdtBlock' && node.type.name !== 'sdtInline') return true;
    const picture = pictureBindingOf(node.attrs['sdtPr'], template);
    if (picture) out.push({ pos, node, picture });
    return true;
  });
  return out;
}

/** Base64 image data, by its head: the value is image bytes, not text or markup. */
export function looksLikeImageBase64(value: string): boolean {
  const head = imageBase64Of(value).slice(0, 16);
  return /^(iVBORw0KGgo|\/9j\/|R0lGOD|Qk[0-9A-Za-z]|SUkq|TU0A|PHN2Zy|PD94bWw|AQAAAM|1wAAAA|AQAAAG)/.test(head);
}

/**
 * A node's image data as base64: a data URL's prefix (`data:image/png;base64,`) dropped, as the engine's bind step reads the
 * node (core-ts `customxml/pictures.mts`, `base64Of`; ED-005 section 12.125). Word writes bare base64; a node a person or a
 * program filled holds a data URL as often. Everything that reads a node's image goes through it, never the value as it is:
 * the engine's lenient decoder would read the prefix's letters as image bytes.
 */
export function imageBase64Of(value: string): string {
  return value.replace(/^\s*data:[^,]*,/, '').trim();
}

/** The bytes of a base64 value or a data URL, whitespace ignored. Throws on a value that is not base64. */
export function base64Bytes(base64: string): Uint8Array {
  const binary = atob(imageBase64Of(base64).replace(/\s+/g, ''));
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

/** What a node's image data is, read from the image's header by the engine (`imageInfoOf`). */
export interface PictureFacts {
  /** The format's name: PNG, JPEG, GIF, BMP or SVG. */
  format: string;
  contentType: string;
  /** In pixels; 0 for an SVG, which states none. */
  widthPx: number;
  heightPx: number;
  /** The image's size in bytes, decoded. */
  bytes: number;
  /** Whether the editor can place it as a new picture (`session.addInlinePicture`): PNG, JPEG, GIF and BMP; an SVG has no pixel size to place it at. */
  placeable: boolean;
}

const FORMAT_NAMES: Record<string, string> = { png: 'PNG', jpeg: 'JPEG', jpg: 'JPEG', gif: 'GIF', bmp: 'BMP', svg: 'SVG' };

/** The last values read, since a panel asks on every render and a picture's base64 runs to megabytes. */
const factsCache = new Map<string, PictureFacts | undefined>();

/**
 * The facts of a node's value as image data (ED-005 section 12.117, proposal 45 part 3's tabs: the Control tab's picture
 * facts, the Data tab's thumbnail and its bind), or undefined when the value is no image the engine reads: text, markup,
 * or a format it does not (TIFF, EMF, WMF; `looksLikeImageBase64` tells those apart from text).
 */
export function pictureFactsOf(value: string): PictureFacts | undefined {
  if (factsCache.has(value)) return factsCache.get(value);
  let facts: PictureFacts | undefined;
  if (looksLikeImageBase64(value)) {
    try {
      const bytes = base64Bytes(value);
      const info = imageInfoOf(bytes);
      const format = FORMAT_NAMES[info.extension.toLowerCase()] ?? info.extension.toUpperCase();
      facts = { format, contentType: info.contentType, widthPx: info.widthPx, heightPx: info.heightPx, bytes: bytes.length, placeable: info.widthPx > 0 && info.heightPx > 0 };
    } catch {
      facts = undefined;
    }
  }
  if (factsCache.size >= 32) factsCache.delete(factsCache.keys().next().value!);
  factsCache.set(value, facts);
  return facts;
}

const EMU_PER_POINT = 12700;
/** A picture's resolution where its header states none, the engine's. */
const DEFAULT_DPI = 96;

/**
 * The size in points at which Word shows a mapped picture control's node picture over the picture the control holds: fitted
 * inside that picture's box, keeping the node picture's aspect ratio, flag or not (ED-005 section 12.126; measured in Word 365,
 * the editor's round picture-placeholder and core-ts's rounds core-ts-picture-fit; the engine's `fitInsideBox`, core-ts CR-005
 * section 10.2, E1, whose EMU it gives). Undefined where the value is no image with a pixel size, or the box has no size.
 */
export function fittedInBox(value: string, box: { width: number; height: number }): { width: number; height: number } | undefined {
  const cx = Math.round(box.width * EMU_PER_POINT);
  const cy = Math.round(box.height * EMU_PER_POINT);
  if (!(cx > 0 && cy > 0) || !looksLikeImageBase64(value)) return undefined;
  let info: ReturnType<typeof imageInfoOf>;
  try { info = imageInfoOf(base64Bytes(value)); } catch { return undefined; }
  if (!(info.widthPx > 0 && info.heightPx > 0)) return undefined;
  const w = info.widthPx / (info.dpiX || DEFAULT_DPI);
  const h = info.heightPx / (info.dpiY || DEFAULT_DPI);
  const fitted = cx * h <= cy * w ? { cx, cy: Math.max(1, Math.round((cx * h) / w)) } : { cx: Math.max(1, Math.round((cy * w) / h)), cy };
  return { width: fitted.cx / EMU_PER_POINT, height: fitted.cy / EMU_PER_POINT };
}

/** A handler control's tag with its `width` set: points (written in twentieths of a point, as REQ-062 reads it), `auto`, or none. */
export function withPictureWidth(tag: string, width: number | 'auto' | undefined): string {
  return formatTag(withTagParam(parseTag(tag), 'width', width === undefined ? undefined : width === 'auto' ? 'auto' : String(Math.round(width * 20))));
}

