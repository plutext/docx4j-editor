// SPDX-License-Identifier: GPL-3.0-or-later
//
// `document.pageNumbers` (ED-005 section 12.95): a plain centred page number in every section's
// footer. The part writes are the `prepare` half (a footer's paragraph through its fragment, a
// footer part made where a section has none to inherit, the `Footer` style spliced); `run` is
// the body's half, the new footers referenced from their sections in one transaction. Editor-only:
// the content API has no form for a footer made and referenced (no `api` line).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { revive } from '../json.mjs';
import { pageNumberParagraph, planPageNumbers, PAGE_NUMBER_STYLE, type PageNumberVariant } from '../pagenumbers.mjs';
import type { ParagraphAttrs } from '../schema.mjs';
import { register, sessionOf, storyOf } from './spec.mjs';

interface PageNumberArgs {
  /** Filled in by `prepare`: the footers made, by the section that references each. */
  refs?: { section: number; variant: PageNumberVariant; rId: string }[];
  /** Filled in by `prepare`: whether a footer the document had was given a page number. */
  extended?: boolean;
}

register<PageNumberArgs | void>({
  name: 'document.pageNumbers',
  label: 'Page numbers',
  tracking: 'untracked',
  prepare: async (state, args) => {
    const session = sessionOf(state);
    if (!session || storyOf(state) !== undefined) return args ?? {};
    const plan = await planPageNumbers(session, state.doc);
    if (plan.extend.length === 0 && plan.create.length === 0) return { ...(args ?? {}), refs: [], extended: false };
    await session.ensureStyles([PAGE_NUMBER_STYLE]);
    for (const story of plan.extend) session.stories.appendBlocks(story, [pageNumberParagraph()]);
    const refs: NonNullable<PageNumberArgs['refs']> = [];
    for (const { section, variant } of plan.create) {
      const footer = session.addHeaderFooter('footer', { TYPE_NAME: 'org_docx4j_wml.Ftr', content: [pageNumberParagraph()] } as unknown as wml.Ftr);
      refs.push({ section, variant, rId: footer.rId });
    }
    return { ...(args ?? {}), refs, extended: plan.extend.length > 0 };
  },
  run: (args) => (state, dispatch) => {
    if (storyOf(state) !== undefined) return false;
    const refs = args?.refs ?? [];
    if (refs.length === 0) return args?.extended === true;
    // The section-break paragraphs in order; an index past them is the final section.
    const positions: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type.name !== 'paragraph') return true;
      if ((node.attrs as ParagraphAttrs).sectPr) positions.push(pos);
      return false;
    });
    const tr = state.tr;
    let final: wml.SectPr | undefined;
    for (const ref of refs) {
      const reference = { TYPE_NAME: 'org_docx4j_wml.FooterReference', type: ref.variant, id: ref.rId };
      const pos = positions[ref.section];
      if (pos !== undefined) {
        const node = tr.doc.nodeAt(pos)!;
        const sectPr = revive((node.attrs as ParagraphAttrs).sectPr) as wml.SectPr;
        ((sectPr.egHdrFtrReferences ??= []) as unknown[]).push(reference);
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, sectPr });
      } else {
        final ??= (state.doc.attrs['finalSectPr'] ? revive(state.doc.attrs['finalSectPr']) : { TYPE_NAME: 'org_docx4j_wml.SectPr' }) as wml.SectPr;
        ((final.egHdrFtrReferences ??= []) as unknown[]).push(reference);
      }
    }
    if (final) tr.setDocAttribute('finalSectPr', final);
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => storyOf(state) === undefined,
});
