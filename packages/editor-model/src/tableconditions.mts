// SPDX-License-Identifier: GPL-3.0-or-later
//
// A table style's conditional formatting for the page (ED-005 section 12.65): which of a table style's
// conditional formats (`w:tblStylePr`: the header row, the first column, banding) each cell is under, as
// the engine decides it (core-ts CR-007, `PropertyResolver.tableContext`), and what they give the cell.
//
// The engine reads a `w:tbl`. A table node is not one, and exporting it on a keystroke would cost the
// table's whole export, so it is handed a stand-in made of what it reads: the table's `w:tblPr`, each
// row's `w:trPr` (`w:gridBefore`, `w:gridAfter`, `w:cnfStyle`), and a `w:tc` per cell with its `w:tcPr`
// and its span; a placeholder `w:tc` stands where a vertical merge from a row above covers a column, as
// the file has a continuation cell there, so the cells after it count their columns as Word does.
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { CellContext, TableContext } from '@docx4j/core-ts';
import type { EditorDocument } from './session.mjs';
import { revive } from './json.mjs';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** A cell's place in its table's conditional formatting. */
export interface CellConditions {
  /** The class the page draws it by: the table style and the conditional formats that apply to it, `tc-<style>-<mask>`. */
  key: string;
  cell: CellContext;
  table: TableContext;
}

/**
 * A paragraph whose own `w:cnfStyle` puts its text under other conditions than its cell's, as the engine
 * honours it (`TableContext.forParagraph`): the class the page draws its text by, `tp-<style>-<mask>`.
 */
export interface ParagraphConditions {
  key: string;
  context: CellContext;
}

/** The conditional format types, in a fixed order, for a class's mask. */
const TYPES = ['wholeTable', 'firstRow', 'lastRow', 'firstCol', 'lastCol', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'neCell', 'nwCell', 'seCell', 'swCell'];

/** The mask of the conditional formats that apply: those the style has, which these conditions select. */
function maskOf(table: TableContext, cell: CellContext): number {
  let mask = 0;
  for (const s of table.applicable(cell.conditions)) {
    const i = TYPES.indexOf(String(s.type));
    if (i >= 0) mask |= 1 << i;
  }
  return mask;
}

/** A `w:gridBefore` or `w:gridAfter` count in a row's properties. */
function gridOf(trPr: wml.TrPr | undefined, name: 'gridBefore' | 'gridAfter'): number {
  const item = (trPr?.cnfStyleOrDivIdOrGridBefore ?? []).find((e) => (e as { name?: { localPart?: string } }).name?.localPart === name) as { value?: { val?: number } } | undefined;
  const val = item?.value?.val;
  return typeof val === 'number' ? Math.max(0, Math.trunc(val)) : 0;
}

const element = (localPart: string, value: object): { name: { namespaceURI: string; localPart: string }; value: object } => ({ name: { namespaceURI: W_NS, localPart }, value });

/** A class fragment for a table style id. */
function classId(id: string | undefined): string {
  return (id ?? 'none').replace(/[^A-Za-z0-9_-]/g, (c) => `_${c.charCodeAt(0).toString(16)}`);
}

/**
 * The conditions of every cell of a table node that its table style formats (its text, or the cell
 * itself through a conditional format's `w:tcPr`), keyed by the cell node; nested tables are their own.
 * Empty when the document has no styles or the table's style formats nothing conditionally.
 */
export function cellConditionsOf(session: EditorDocument, table: PMNode): Map<PMNode, CellConditions> {
  return tableConditionsOf(session, table).cells;
}

/**
 * A table node's conditional formatting: its cells' (`cellConditionsOf`), and the paragraphs in them (directly,
 * or in a block control directly in the cell) whose own `w:cnfStyle` gives their text other conditions.
 */
export function tableConditionsOf(session: EditorDocument, table: PMNode): { cells: Map<PMNode, CellConditions>; paragraphs: Map<PMNode, ParagraphConditions> } {
  const out = new Map<PMNode, CellConditions>();
  const paragraphs = new Map<PMNode, ParagraphConditions>();
  const resolver = session.pkg.propertyResolverOrUndefined;
  if (!resolver || table.type.name !== 'table') return { cells: out, paragraphs };
  const tblPr = table.attrs['tblPr'] ? (revive(table.attrs['tblPr']) as wml.TblPr) : undefined;
  // The rows as the file has them: a cell's `w:tc`, and a placeholder where a merge from above covers a column.
  const covered: number[] = [];   // per grid column, the rows a merge from above still covers
  const rows: { tr: wml.Tr; cells: { node: PMNode; tc: wml.Tc }[] }[] = [];
  table.forEach((row) => {
    const trPr = row.attrs['trPr'] ? (revive(row.attrs['trPr']) as wml.TrPr) : undefined;
    const content: object[] = [];
    const cells: { node: PMNode; tc: wml.Tc }[] = [];
    let col = gridOf(trPr, 'gridBefore');
    const fill = (): void => {
      while ((covered[col] ?? 0) > 0) {
        let span = 1;
        while ((covered[col + span] ?? 0) > 0 && covered[col + span] === covered[col]) span++;
        content.push(element('tc', { TYPE_NAME: 'org_docx4j_wml.Tc', tcPr: { TYPE_NAME: 'org_docx4j_wml.TcPr', ...(span > 1 ? { gridSpan: { val: span } } : {}) } }));
        for (let k = 0; k < span; k++) covered[col + k]! -= 1;
        col += span;
      }
    };
    row.forEach((cell) => {
      fill();
      const colspan = (cell.attrs['colspan'] as number) || 1;
      const rowspan = (cell.attrs['rowspan'] as number) || 1;
      const own = cell.attrs['tcPr'] ? (revive(cell.attrs['tcPr']) as wml.TcPr) : undefined;
      const tc: wml.Tc = { TYPE_NAME: 'org_docx4j_wml.Tc', tcPr: { ...(own ?? { TYPE_NAME: 'org_docx4j_wml.TcPr' }), ...(colspan > 1 ? { gridSpan: { val: colspan } } : {}) } } as wml.Tc;
      content.push(element('tc', tc));
      cells.push({ node: cell, tc });
      for (let k = 0; k < colspan; k++) covered[col + k] = rowspan - 1;
      col += colspan;
    });
    fill();
    const tr = { TYPE_NAME: 'org_docx4j_wml.Tr', ...(trPr ? { trPr } : {}), content } as unknown as wml.Tr;
    rows.push({ tr, cells });
  });
  const tbl = { TYPE_NAME: 'org_docx4j_wml.Tbl', ...(tblPr ? { tblPr } : {}), content: rows.map((r) => element('tr', r.tr)) } as unknown as wml.Tbl;
  let context: TableContext;
  try {
    context = resolver.tableContext(tbl);
  } catch {
    return { cells: out, paragraphs };
  }
  const style = context.tableStyle;
  if (!style.tblStylePr?.length && !style.pPr && !style.rPr && !style.tcPr) return { cells: out, paragraphs };
  const ids = session.ctx.ids;
  for (const r of rows) {
    for (const c of r.cells) {
      const cell = context.forCell(r.tr, c.tc);
      const formatsCell = style.tcPr !== undefined || context.applicable(cell.conditions).some((s) => s.tcPr !== undefined);
      if (cell.formatsText || formatsCell) out.set(c.node, { key: `tc-${classId(cell.tableStyleId)}-${maskOf(context, cell)}`, cell, table: context });
      // A paragraph's own w:cnfStyle (Word 2007 and 2010 wrote them as caches): its text's conditions, where they differ.
      const visit = (p: PMNode): void => {
        if (p.type.name === 'sdtBlock') { p.forEach(visit); return; }
        if (p.type.name !== 'paragraph') return;
        const source = p.attrs['sourceId'] ? (ids.get(p.attrs['sourceId'] as string)?.value as wml.P | undefined) : undefined;
        if (!source?.pPr?.cnfStyle) return;
        const own = context.forParagraph(r.tr, c.tc, source.pPr);
        if (own.key.equals(cell.key) || !own.formatsText) return;
        paragraphs.set(p, { key: `tp-${classId(own.tableStyleId)}-${own.key.mask}`, context: own });
      };
      c.node.forEach(visit);
    }
  }
  return { cells: out, paragraphs };
}

/**
 * A cell's properties from its table style as it is placed: the style's own `w:tcPr`, then each conditional
 * format's that applies to it, in Word's order (ECMA-376-1 17.7.6), a later one replacing an earlier one's
 * shading, vertical alignment and each border and margin side. The cell's own `w:tcPr` goes over this on the page.
 */
export function conditionalTcPr(conditions: CellConditions): wml.TcPr {
  const out: wml.TcPr = { TYPE_NAME: 'org_docx4j_wml.TcPr' } as wml.TcPr;
  const layers = [conditions.table.tableStyle.tcPr, ...conditions.table.applicable(conditions.cell.conditions).map((s) => s.tcPr)];
  for (const pr of layers) {
    if (!pr) continue;
    if (pr.shd) out.shd = pr.shd;
    if (pr.vAlign) out.vAlign = pr.vAlign;
    if (pr.tcBorders) out.tcBorders = { ...(out.tcBorders ?? {}), ...definedOf(pr.tcBorders) } as wml.TcPrInner.TcBorders;
    if (pr.tcMar) out.tcMar = { ...(out.tcMar ?? {}), ...definedOf(pr.tcMar) } as wml.TcMar;
  }
  return out;
}

/** An object's defined members. */
function definedOf<T extends object>(o: T): Partial<T> {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as Partial<T>;
}
