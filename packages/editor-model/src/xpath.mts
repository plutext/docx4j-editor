// SPDX-License-Identifier: GPL-3.0-or-later
//
// Select by XPath (ED-005 section 3.5): an expression over a part's XML as the edit layer has it,
// its hits the tree objects the engine's CR-006 binding returns (`bindSnapshot`, `boundsOf`: a DOM
// marshalled for the call, each selected node with the object it was marshalled from), each placed
// in the story's document where its element projects. The session exports the edit layer into the
// tree first, as a save does (the first half of a tree run, never handing the tree to the part, so
// an unedited document keeps its bytes), and calls this over the part's root. An element is placed
// at its own projection when it has one (a paragraph, a table, a row, a cell, a control, a run's
// text, an atom), else at its nearest ancestor's; a block the export rebuilt (an edited one) has
// new objects with no projection, so a hit inside it is placed at the block.
import { bindSnapshot, boundsOf, NAMESPACE_PREFIXES, type Element, type XPathEngine } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import type { SourceIds } from './ids.mjs';
import type { Story } from './stories.mjs';

/** One hit, placed. */
export interface XPathHit {
  /** The story the hit is in; undefined for the body. */
  story?: Story;
  /** The top-level block of that story's document it is in. */
  block: number;
  /** Where it projects: the node of its element, or of the nearest ancestor that has one. */
  pos: number;
  /** Placed at its own element's projection (or an ancestor's inside the block), not only at its block's. */
  exact: boolean;
  /** Its path in the part: `/w:document/w:body[1]/w:p[3]/w:r[2]`. */
  path: string;
  /** What was selected: the element's qualified name, `@w:val` for an attribute, `#text`. */
  name: string;
  /** Its text, cut to a line. */
  snippet: string;
}

/** A block container of the part (the body, a note, a header): its content in the tree and its document as the edit layer has it. */
export interface SelectContainer {
  story?: Story;
  content: readonly Element[];
  doc: PMNode;
}

/** The prefixes the object model declares, prefix to URI, as an expression uses them. */
export const XPATH_NAMESPACES: Readonly<Record<string, string>> = (() => {
  const out: Record<string, string> = {};
  for (const [uri, prefix] of Object.entries(NAMESPACE_PREFIXES)) if (prefix !== '' && !(prefix in out)) out[prefix] = uri;
  return out;
})();

const SNIPPET = 80;

/**
 * Where each tree object with a projection is in a document: blocks, rows, cells, controls and
 * atoms by `sourceId`, runs by their text's `runSource`; a row- or cell-level control (an
 * attribute stack, not a node) at the first row or cell of its group, and a vertical merge's
 * continuation cell (folded into the cell above) at that cell.
 */
function placesOf(doc: PMNode, ids: SourceIds): Map<object, number> {
  const out = new Map<object, number>();
  const put = (id: unknown, pos: number): void => {
    if (typeof id !== 'string') return;
    const value = ids.get(id)?.value;
    if (typeof value === 'object' && value !== null && !out.has(value)) out.set(value, pos);
  };
  doc.descendants((node, pos) => {
    put(node.attrs['sourceId'], pos);
    put(node.attrs['runSource'], pos);
    for (const g of (node.attrs['sdt'] as { groupId: string }[] | null | undefined) ?? []) put(g.groupId, pos);
    for (const id of (node.attrs['merged'] as string[] | null | undefined) ?? []) put(id, pos);
    for (const m of node.marks) if (m.type.name === 'runSource') put(m.attrs['sourceId'], pos);
    return true;
  });
  return out;
}

/** The position of top-level block `i` of a document. */
function blockPos(doc: PMNode, i: number): number {
  let pos = 0;
  for (let k = 0; k < i && k < doc.childCount; k++) pos += doc.child(k).nodeSize;
  return pos;
}

/**
 * The hits of an expression over a part's root, placed in its containers' documents. `engine` is
 * the package's (`pkg.xpathEngine`: FontoXPath once the Template perspective or the XPath tab
 * installed it, XPath 3.1; the default XPath 1.0 engine otherwise); the object model's prefixes
 * are bound, and `namespaces` adds to them. Throws the engine's error for an invalid expression.
 */
export async function selectIn(root: Element, containers: readonly SelectContainer[], xpath: string, engine: XPathEngine, ids: SourceIds, namespaces: Record<string, string> = {}): Promise<XPathHit[]> {
  const snapshot = await bindSnapshot(root);
  await engine.ready(snapshot.root);
  const bounds = boundsOf(engine.select(xpath, snapshot.root, { ...XPATH_NAMESPACES, ...namespaces }), snapshot);
  const places = containers.map((c) => placesOf(c.doc, ids));
  const blocks = new Map<object, { c: number; i: number }>();
  containers.forEach((c, k) => c.content.forEach((el, i) => { if (typeof el.value === 'object' && el.value !== null) blocks.set(el.value, { c: k, i }); }));
  const out: XPathHit[] = [];
  for (const bound of bounds) {
    let placed: { c: number; pos: number; exact: boolean } | undefined;
    let fallback: { c: number; i: number } | undefined;
    // Up the tree from the selected object to the first ancestor that projects; the block it sits in on the way.
    for (let o = bound.object as { PARENT?: unknown } | undefined; o && typeof o === 'object'; o = o.PARENT as { PARENT?: unknown } | undefined) {
      if (!placed) {
        for (let k = 0; k < containers.length; k++) {
          const pos = places[k]!.get(o);
          if (pos !== undefined) { placed = { c: k, pos, exact: true }; break; }
        }
      }
      const b = blocks.get(o);
      if (b) { fallback = b; break; }
    }
    if (!placed && !fallback) continue;   // above the blocks: the root, the body element, the final section
    const c = placed?.c ?? fallback!.c;
    const container = containers[c]!;
    const pos = placed?.pos ?? blockPos(container.doc, fallback!.i);
    const node = bound.node;
    const text = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
    out.push({
      ...(container.story ? { story: container.story } : {}),
      block: placed ? container.doc.resolve(Math.min(pos, container.doc.content.size)).index(0) : fallback!.i,
      pos,
      exact: placed !== undefined,
      path: bound.path,
      name: bound.attribute ? `@${bound.attribute}` : node.nodeType === 3 ? '#text' : node.nodeName,
      snippet: text.length > SNIPPET ? `${text.slice(0, SNIPPET)}…` : text,
    });
  }
  return out;
}

/** The console script that selects the same (the XPath tab's "Open in console"): the content-API views the expression selects. */
export function selectScript(xpath: string): string {
  return `// Select by XPath: the content-API views the expression selects (CR-006): a Paragraph, Table, TableRow,\n// TableCell, ContentControl, a Range over a run's paragraph, or the Bound (an attribute, a text node).\nconst hits = await body.select(${JSON.stringify(xpath)});\nreturn hits.map((h) => ({ kind: h.constructor.name, text: 'text' in h ? h.text : undefined, path: 'path' in h ? h.path : undefined }));\n`;
}
