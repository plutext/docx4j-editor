// SPDX-License-Identifier: GPL-3.0-or-later
//
// The preview (ED-003 section 4.7): the bind step in place on a scratch copy of the package. The
// package is cloned (`OpcPackage.clone()`, a save into memory and a reload, so nothing is shared
// with the live one), the clone's body is replaced by a deep copy of what the edit layer would
// export now (unsaved edits included), the engine's `applyBindings` pushes the data part's values
// into the clone's bound controls (REQ-021's run properties, REQ-022's lines, REQ-024's
// placeholder are the engine's), and the clone's body is projected into a story of its own with
// its own ids, so nothing of it lands in the session's context. Repeats are not expanded and
// dropped controls stay (section 9 item 10): the verdicts say what a processor would do with them.
// An XHTML-bound control is converted as the processor converts it (ED-005 section 12.117): the
// engine's bind step does it, with a browser's `DOMParser` unless `html.parser` is given (Node has
// none). The conversion can add styles and lists the live document has not, so the preview
// carries its own clone's numbering and list labels, and the view draws its styles from the
// clone's resolver.
import type { WordprocessingMLPackage, MainDocumentPart, BindingResult, XhtmlBindingOptions } from '@docx4j/core-ts';
import { deepCopy, linkParents } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorDocument } from './session.mjs';
import type { ControlLookup } from './export.mjs';
import type { Story } from './stories.mjs';
import { newContext, projectBlocks } from './import.mjs';
import { liveLabels, type LiveLabel } from './labels.mjs';
import { numberingOfPart, type Numbering } from './commands/lists.mjs';

export interface Preview extends Story {
  /** The scratch package the values were pushed into; the story's images and fields resolve through its main part. */
  pkg: WordprocessingMLPackage;
  main: MainDocumentPart;
  /** What the bind step reported. */
  result: BindingResult;
  /** The row and cell groups of the preview, from its own projection. */
  controls: ControlLookup;
  /** The clone's list definitions: the document's, and any list an XHTML conversion made. */
  numbering: Numbering | undefined;
  /** The list labels of the preview's paragraphs, from the clone's numbering. */
  labels: Map<PMNode, LiveLabel>;
}

/** The instance the bind step would make of the document as it is now, as a read-only story. */
export async function previewInstance(session: EditorDocument, options: { html?: XhtmlBindingOptions } = {}): Promise<Preview> {
  const pkg = await session.pkg.clone();
  const main = pkg.getMainDocumentPart();
  const document = await main.readContents();
  const exported = session.exportTree();
  const body: wml.Body = { ...(document.body ?? { TYPE_NAME: 'org_docx4j_wml.Body' }), content: deepCopy(exported.content) as wml.Body['content'] };
  if (exported.sectPr === undefined) delete body.sectPr; else body.sectPr = deepCopy(exported.sectPr);
  document.body = body;
  linkParents(body, document);
  main.setContents(document);
  await pkg.customXmlParts.load();
  const result = await pkg.customXmlParts.applyBindings(options.html ? { html: options.html } : {});
  // The clone's own resolver, run-font selector and numbering, read after the bind step: a conversion adds styles and lists.
  await pkg.getPropertyResolver();
  await main.getRunFontSelector();
  await pkg.getNumberingEmulator();
  const ctx = newContext(session.schema);
  let blocks = projectBlocks((document.body.content ?? []) as unknown[] as Parameters<typeof projectBlocks>[0], ctx, 'preview');
  if (blocks.length === 0) blocks = [session.schema.nodes['paragraph']!.create()];
  const doc = session.schema.topNodeType.create({ finalSectPr: body.sectPr ?? null }, blocks);
  const controls: ControlLookup = (groupId) => {
    const g = ctx.groups.get(groupId);
    return g && g.kind === 'sdt' ? { form: g.form, sdtPr: g.sdtPr } : undefined;
  };
  const numbering = numberingOfPart(main);
  return { key: 'preview', doc, ctx, pkg, main, result, controls, numbering, labels: liveLabels(doc, session, numbering) };
}
