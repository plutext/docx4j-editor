// SPDX-License-Identifier: GPL-3.0-or-later
//
// Escaped XHTML bindings (ED-005 proposal 44, section 12.75; the specification's section 9). A
// control whose tag carries `od:ContentType=application/xhtml+xml` and names an `od:xpath` entry is
// bound to the entry's node, whose string value is XHTML markup; it carries no `w:dataBinding`
// (Word allows none on a block-level control, and docx4j's bind.xslt reads the entry instead).
// The view converts the markup and writes the content (`view/xhtmlbind.ts`, through the editor's
// HTML converter until the engine's release carries it); nothing is written back to the node.
import type { Node as PMNode } from 'prosemirror-model';
import { readTag } from './opendope.mjs';
import { isFlatOpc } from './bound.mjs';
import { sdtTag, type DataBinding } from './sdtpr.mjs';
import type { TemplateModel } from './template.mjs';

export const XHTML_CONTENT_TYPE = 'application/xhtml+xml';

/** Whether a control's tag binds it to XHTML: the content type and an `od:xpath` entry named. */
export function isXhtmlControl(sdtPr: unknown): boolean {
  const tag = readTag(sdtTag(sdtPr));
  return tag.contentType === XHTML_CONTENT_TYPE && tag.xpath !== undefined;
}

/** The binding an XHTML control's `od:xpath` entry gives, as a `w:dataBinding` would; undefined without the entry (the template must be loaded). */
export function xhtmlBindingOf(sdtPr: unknown, template: TemplateModel): DataBinding | undefined {
  if (!isXhtmlControl(sdtPr)) return undefined;
  const id = readTag(sdtTag(sdtPr)).xpath!;
  const entry = template.xpaths?.entries.find((e) => e.id === id);
  if (!entry) return undefined;
  return { xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings, w15: false };
}

export interface XhtmlControl {
  pos: number;
  node: PMNode;
  binding: DataBinding;
}

/** Every XHTML-bound control of the body with an entry, in document order. */
export function xhtmlControlsIn(doc: PMNode, template: TemplateModel): XhtmlControl[] {
  const out: XhtmlControl[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'sdtBlock' && node.type.name !== 'sdtInline') return true;
    const binding = xhtmlBindingOf(node.attrs['sdtPr'], template);
    if (binding) out.push({ pos, node, binding });
    return true;
  });
  return out;
}

/** HTML's block-level elements, as the converter treats them (a paragraph of its own, a list, a table, a heading). */
const BLOCK_TAGS = /<\/?(p|div|h[1-6]|ul|ol|li|table|thead|tbody|tfoot|tr|td|th|blockquote|pre|hr|section|article|header|footer|nav|dl|dt|dd|figure|figcaption|address|main|aside)\b/i;

/** Whether a node's value is XHTML markup, and of what level: `block` when it holds a block element, `inline` when tags only, `none` for plain text. */
export function xhtmlKindOf(value: string): 'none' | 'inline' | 'block' {
  const v = value.trim();
  // A Word document as Flat OPC is markup too, but no HTML: a bare `pkg:package` read as inline XHTML (ED-005 proposal 46).
  if (isFlatOpc(v)) return 'none';
  if (!/^<([a-zA-Z][\w:-]*)[\s>/]/.test(v) || !/<\/[a-zA-Z][\w:-]*\s*>|\/>/.test(v)) return 'none';
  return BLOCK_TAGS.test(v) ? 'block' : 'inline';
}
