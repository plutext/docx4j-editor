// SPDX-License-Identifier: GPL-3.0-or-later
//
// Sections (ED-002 section 4.6; ED-001 section 3.1): a section break is `sectPr` on the paragraph
// that ends the section; the final section is `doc.attrs.finalSectPr`, a projection of the
// store's `package` map. Every property command rewrites a copy of the section's JSON, so header
// references, line numbering and the rest ride through. Points in, twips written.
import type { Command, EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { revive } from '../json.mjs';
import type { ParagraphAttrs } from '../schema.mjs';
import { twips } from './paragraph.mjs';
import { register, storyOf } from './spec.mjs';

export type SectionTarget =
  | { kind: 'paragraph'; pos: number; node: PMNode; sectPr: unknown }
  | { kind: 'final'; sectPr: unknown };

/** The section a position is in: the first paragraph at or after it that ends a section, else the final one. */
export function sectionAt(state: EditorState, pos: number = state.selection.from): SectionTarget {
  let found: SectionTarget | undefined;
  state.doc.nodesBetween(pos, state.doc.content.size, (node, nodePos) => {
    if (found) return false;
    if (node.type.name === 'paragraph') {
      const sectPr = (node.attrs as ParagraphAttrs).sectPr;
      if (sectPr !== null && nodePos + node.nodeSize > pos) found = { kind: 'paragraph', pos: nodePos, node, sectPr };
      return false;
    }
    return true;
  });
  return found ?? { kind: 'final', sectPr: state.doc.attrs['finalSectPr'] };
}

/** Sections are the body's: in a note or a header (ED-003 section 6.4) every section command refuses. */
const inBody = (state: EditorState): boolean => storyOf(state) === undefined;

/** Rewrite the caret's section through `edit` on a copy of its JSON. */
export function updateSection(edit: (sectPr: wml.SectPr) => void): Command {
  return (state, dispatch) => {
    if (!inBody(state)) return false;
    const target = sectionAt(state);
    const sectPr = (target.sectPr === null || target.sectPr === undefined ? { TYPE_NAME: 'org_docx4j_wml.SectPr' } : revive(target.sectPr)) as wml.SectPr;
    edit(sectPr);
    const tr = state.tr;
    if (target.kind === 'paragraph') tr.setNodeMarkup(target.pos, undefined, { ...target.node.attrs, sectPr });
    else tr.setDocAttribute('finalSectPr', sectPr);
    dispatch?.(tr);
    return true;
  };
}

function caretParagraph(state: EditorState): { node: PMNode; pos: number } | undefined {
  const $from = state.selection.$from;
  if ($from.parent.type.name !== 'paragraph') return undefined;
  return { node: $from.parent, pos: $from.before() };
}

type SectionType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';

register<{ type?: SectionType }>({
  name: 'section.insertBreak',
  tracking: 'untracked',
  label: 'Insert section break',
  args: { type: "'nextPage' (default) | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn'" },
  choices: { type: ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'] },
  run: (args) => (state, dispatch) => {
    const p = caretParagraph(state);
    if (!inBody(state) || !p || (p.node.attrs as ParagraphAttrs).sectPr !== null) return false;
    const current = sectionAt(state);
    const copy = (): wml.SectPr => (current.sectPr === null || current.sectPr === undefined ? { TYPE_NAME: 'org_docx4j_wml.SectPr' } : revive(current.sectPr)) as wml.SectPr;
    // The paragraph ends a new section, a copy of the one it is in, which keeps that section's start. The
    // break's kind is the start of the section after it: `w:type` is how a section starts relative to the
    // one before (ECMA-376 17.6.22), so a continuous break makes the rest of the section continuous.
    const tr = state.tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, sectPr: copy() });
    const type = args?.type ?? 'nextPage';
    const following = copy();
    if ((following.type?.val ?? 'nextPage') !== type) {
      if (type === 'nextPage') delete following.type; else following.type = { val: type };
      if (current.kind === 'paragraph') tr.setNodeMarkup(current.pos, undefined, { ...current.node.attrs, sectPr: following });
      else tr.setDocAttribute('finalSectPr', following);
    }
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => { const p = caretParagraph(state); return inBody(state) && p !== undefined && (p.node.attrs as ParagraphAttrs).sectPr === null; },
});

register<void>({
  name: 'section.removeBreak',
  tracking: 'untracked',
  label: 'Remove section break',
  run: () => (state, dispatch) => {
    const p = caretParagraph(state);
    if (!inBody(state) || !p || (p.node.attrs as ParagraphAttrs).sectPr === null) return false;
    dispatch?.(state.tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, sectPr: null }));
    return true;
  },
  isEnabled: (state) => { const p = caretParagraph(state); return inBody(state) && p !== undefined && (p.node.attrs as ParagraphAttrs).sectPr !== null; },
});

register<{ width: number; height: number }>({
  name: 'section.pageSize',
  tracking: 'untracked',
  label: 'Page size',
  args: { width: 'points', height: 'points' },
  run: ({ width, height }) => updateSection((s) => {
    const w = twips(width);
    const h = twips(height);
    s.pgSz = { ...s.pgSz, w, h };
    if (w > h) s.pgSz.orient = 'landscape'; else delete s.pgSz.orient;
  }),
  isEnabled: inBody,
});

register<{ value: 'portrait' | 'landscape' }>({
  name: 'section.orientation',
  tracking: 'untracked',
  label: 'Orientation',
  args: { value: "'portrait' | 'landscape'; swaps width and height" },
  choices: { value: ['portrait', 'landscape'] },
  run: ({ value }) => updateSection((s) => {
    const pgSz = (s.pgSz ??= {});
    const current = pgSz.orient === 'landscape' ? 'landscape' : 'portrait';
    if (current !== value && pgSz.w !== undefined && pgSz.h !== undefined) { const w = pgSz.w; pgSz.w = pgSz.h; pgSz.h = w; }
    if (value === 'landscape') pgSz.orient = 'landscape'; else delete pgSz.orient;
  }),
  isActive: (state, args) => args !== undefined && ((sectionAt(state).sectPr as wml.SectPr | null)?.pgSz?.orient === 'landscape' ? 'landscape' : 'portrait') === args.value,
  isEnabled: inBody,
});

type Margins = Partial<Record<'top' | 'right' | 'bottom' | 'left' | 'header' | 'footer' | 'gutter', number>>;

register<Margins>({
  name: 'section.margins',
  tracking: 'untracked',
  label: 'Margins',
  args: { top: 'points', right: 'points', bottom: 'points', left: 'points', header: 'points', footer: 'points', gutter: 'points' },
  run: (margins) => updateSection((s) => {
    const pgMar = (s.pgMar ??= { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 });
    for (const [k, v] of Object.entries(margins ?? {})) if (typeof v === 'number') (pgMar as unknown as Record<string, number>)[k] = twips(v);
  }),
  isEnabled: inBody,
});

register<{ value: SectionType }>({
  name: 'section.type',
  tracking: 'untracked',
  label: 'Section start',
  args: { value: "'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn'" },
  choices: { value: ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'] },
  run: ({ value }) => updateSection((s) => { if (value === 'nextPage') delete s.type; else s.type = { val: value }; }),
  isEnabled: inBody,
});

register<{ count: number }>({
  name: 'section.columns',
  tracking: 'untracked',
  label: 'Columns',
  args: { count: 'the number of columns' },
  run: ({ count }) => updateSection((s) => {
    const cols = (s.cols ??= {});
    if (count > 1) cols.num = count; else delete cols.num;
  }),
  isEnabled: inBody,
});
