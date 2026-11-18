// SPDX-License-Identifier: GPL-3.0-or-later
//
// Macros kept in the document (ED-001 section 3.5; ED-003 section 7.4): a custom XML part whose
// root is `macros` in the editor's namespace, one `macro` element per macro with its name and the
// script as text. Word ignores the part; the editor offers the macros when the document opens and
// never runs one on its own (ED-001 section 11 item 5). Read through `pkg.customXmlParts` after
// the template view's load (the parts parsed once, no part marked); written through `add`,
// `setXml` and `removePart` under the `markPartChanged` protocol (a part edit: dirty and
// autosave, no undo, section 9 item 7), so a document without macros stays byte for byte.
import type { CustomXmlPart } from '@docx4j/core-ts';
import type { EditorDocument } from './session.mjs';

/** CLAUDE.md's namespace for session state and macros. */
export const MACROS_NAMESPACE = 'http://docx4java.org/editor/2026';

export interface Macro {
  name: string;
  script: string;
}

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** The part's XML for a list of macros. */
export function macrosXml(macros: readonly Macro[]): string {
  const body = macros.map((m) => `  <macro name="${escapeXml(m.name)}">${escapeXml(m.script)}</macro>`).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<macros xmlns="${MACROS_NAMESPACE}">${body === '' ? '' : `\n${body}\n`}</macros>`;
}

/** The macros of a parsed part; an element without a name, or in another namespace, is skipped. */
export function macrosOf(part: CustomXmlPart): Macro[] {
  const root = part.document.documentElement;
  if (!root || root.namespaceURI !== MACROS_NAMESPACE || root.localName !== 'macros') return [];
  const out: Macro[] = [];
  for (let i = 0; i < root.childNodes.length; i++) {
    const node = root.childNodes[i] as Element;
    if (node.nodeType !== 1 || node.localName !== 'macro' || node.namespaceURI !== MACROS_NAMESPACE) continue;
    const name = node.getAttribute('name');
    if (name === null || name === '') continue;
    out.push({ name, script: node.textContent ?? '' });
  }
  return out;
}

/** The document's macros part, if it has one (the first part in the namespace). */
export async function macrosPart(session: EditorDocument): Promise<CustomXmlPart | undefined> {
  if (!session.template.loaded) await session.template.load({ engine: false });
  return session.pkg.customXmlParts.getByNamespace(MACROS_NAMESPACE).find((p) => p.documentElement.baseName === 'macros');
}

/** The macros the document carries; none without the part. */
export async function readDocumentMacros(session: EditorDocument): Promise<Macro[]> {
  const part = await macrosPart(session);
  return part ? macrosOf(part) : [];
}

/**
 * Write the list into the document: the part made on the first macro, rewritten while there are
 * any, removed with the last. A part edit: marked for the save, no undo step.
 */
export async function writeDocumentMacros(session: EditorDocument, macros: readonly Macro[]): Promise<void> {
  const part = await macrosPart(session);
  if (macros.length === 0) {
    if (!part) return;
    const name = part.part.partName.toString();
    session.pkg.customXmlParts.removePart(part);
    session.markPartChanged(name);
  } else if (part) {
    part.setXml(macrosXml(macros));
    session.markPartChanged(part.part.partName.toString());
  } else {
    const made = session.addCustomXmlPart(macrosXml(macros));
    session.markPartChanged(made.part.partName.toString());
  }
  session.template.invalidate();
}
