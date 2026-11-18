// SPDX-License-Identifier: GPL-3.0-or-later
//
// Page numbers in every section's footer (ED-005 section 12.95, Jason's request): Word's plain
// centred number, a `Footer` paragraph holding a simple PAGE field. A footer a section has (its
// own default footer, or its even footer where the document has different odd and even pages)
// gets the paragraph at its end unless it already holds a PAGE field; a section without one
// inherits the previous section's, and where no section before it has one a footer part is made
// and referenced from that section. The first-page footer (`w:titlePg`) is left alone, as a
// title page is not numbered.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import { paragraphOf, runOf, t, type Element } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorDocument } from './session.mjs';
import type { HeaderFooter } from './stories.mjs';

export const PAGE_NUMBER_STYLE = 'Footer';

/** The footer variants a page number goes in: the default always, the even one where the document has odd and even pages. */
export type PageNumberVariant = 'default' | 'even';

/** Word's plain centred page number: a `Footer` paragraph holding a simple PAGE field with "1" as its cached result. */
export function pageNumberParagraph(): Element<wml.P> {
  const field = el.fldSimple({ TYPE_NAME: 'org_docx4j_wml.CTSimpleField', instr: ' PAGE ', content: [runOf([t('1')])] } as unknown as wml.CTSimpleField) as unknown as Element;
  return paragraphOf([field], { TYPE_NAME: 'org_docx4j_wml.PPr', pStyle: { val: PAGE_NUMBER_STYLE }, jc: { val: 'center' } } as unknown as wml.PPr);
}

const isPageInstr = (instr: unknown): boolean => typeof instr === 'string' && /(^|\s)PAGE(?=\s|$)/.test(instr.trim());

/** Whether an object-model tree holds a PAGE field: a simple field, or a complex field's instruction text. */
export function hasPageField(tree: unknown): boolean {
  const seen = new Set<object>();
  const visit = (v: unknown): boolean => {
    if (!v || typeof v !== 'object' || seen.has(v)) return false;
    seen.add(v);
    if (Array.isArray(v)) return v.some(visit);
    const o = v as Record<string, unknown>;
    if (o['TYPE_NAME'] === 'org_docx4j_wml.CTSimpleField' && isPageInstr(o['instr'])) return true;
    const name = o['name'] as { localPart?: string } | undefined;
    if (name?.localPart === 'instrText' && isPageInstr((o['value'] as { value?: unknown } | undefined)?.value)) return true;
    for (const [k, child] of Object.entries(o)) if (k !== 'parent' && visit(child)) return true;
    return false;
  };
  return visit(tree);
}

export interface PageNumberPlan {
  /** Footers the document has that lack a page number: the paragraph goes at their end. */
  extend: HeaderFooter[];
  /** Sections that get a footer made for them, by index and variant: the first section of each run that has none to inherit. */
  create: { section: number; variant: PageNumberVariant }[];
}

/** What a page-number insertion has to do, over the sections of `doc` (a view's state) and the footers read. */
export async function planPageNumbers(session: EditorDocument, doc: PMNode = session.doc): Promise<PageNumberPlan> {
  const sections = session.stories.sections(doc);
  const footers = (await session.stories.headers()).filter((h) => h.kind === 'footer');
  const variants: PageNumberVariant[] = (await session.stories.evenAndOddHeaders()) ? ['default', 'even'] : ['default'];
  const plan: PageNumberPlan = { extend: [], create: [] };
  for (const variant of variants) {
    // The footer the sections so far inherit: one of the file's, or one this plan makes.
    let current: HeaderFooter | 'new' | undefined;
    sections.forEach((sectPr, index) => {
      const refs = (sectPr?.egHdrFtrReferences ?? []) as (wml.HeaderReference | wml.FooterReference)[];
      const own = refs.find((r) => r.TYPE_NAME === 'org_docx4j_wml.FooterReference' && r.type === variant);
      const story = own ? footers.find((f) => f.rId === own.id) : undefined;
      if (story) {
        current = story;
        if (!hasPageField(story.element) && !plan.extend.includes(story)) plan.extend.push(story);
      } else if (current === undefined) {
        plan.create.push({ section: index, variant });
        current = 'new';
      }
    });
  }
  return plan;
}
