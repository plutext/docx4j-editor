// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tables on screen (ED-003 section 3.8): the effective table properties from the resolver
// (`getEffectiveTableStyle`, the style chain with the table's own `w:tblPr` on top) become
// CSS variables on the `<table>` that the cell rules read for borders and margins; a cell's own
// `w:tcPr` becomes its inline style. Widths are twips to points as the page is. Conditional
// formatting (`w:tblStylePr`: first row, banding) is not applied in E2.a and is recorded as an
// approximation.
import type { ColorResolver } from './css.js';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { type EditorDocument, revive } from '@docx4j/editor-model';

const BORDER_STYLES: Record<string, string> = {
  single: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed', dotDotDash: 'dotted',
  triple: 'double', wave: 'solid', doubleWave: 'double', inset: 'inset', outset: 'outset', thinThickSmallGap: 'double', thickThinSmallGap: 'double',
};

/**
 * A `w:*Border` as a CSS border shorthand; `none` for `nil` and `none`; undefined when absent. Its colour through
 * the resolver where there is one (a theme reference over the cached value, ED-008 section 8.8 item 13).
 */
export function borderCss(border: wml.CTBorder | undefined, colors?: ColorResolver): string | undefined {
  if (!border) return undefined;
  const val = String(border.val);
  if (val === 'nil' || val === 'none') return 'none';
  const style = BORDER_STYLES[val] ?? 'solid';
  // `w:sz` is in eighths of a point; a border with no size is Word's 0.5 pt.
  const pt = border.sz !== undefined ? Math.max(0.25, border.sz / 8) : 0.5;
  const resolved = colors ? colors.resolveBorderColor(border) : undefined;
  const color = colors ? (resolved ? `#${resolved.toHex()}` : 'currentColor')
    : border.color && border.color !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(border.color) ? `#${border.color}` : 'currentColor';
  return `${pt}pt ${style} ${color}`;
}

/**
 * A `w:shd` as a CSS background; undefined for none. Its fill through the resolver where there is one (item 13).
 * `w:val="clear"` with an automatic fill is an explicit no colour, `transparent`, which hides a table style's shading
 * or a band's beneath it, where `w:val="nil"` is no shading at all and lets it show: Word 365 draws them so (ED-005
 * section 12.128, the row-shading round, cells A2 and B3 against A3 and B5).
 */
export function shadingCss(shd: wml.CTShd | undefined, colors?: ColorResolver): string | undefined {
  if (!shd) return undefined;
  const resolvedFill = colors ? colors.resolveFill(shd) : undefined;
  const fill = colors ? (resolvedFill ? `#${resolvedFill.toHex()}` : undefined)
    : shd.fill && shd.fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(shd.fill) ? `#${shd.fill}` : undefined;
  if (fill) return fill;
  // A pattern with no fill: a light grey stands in for the pattern.
  const val = String(shd.val ?? 'clear');
  if (val === 'nil') return undefined;
  if (val === 'clear') return 'transparent';
  if (val.startsWith('pct')) { const pct = Number(val.slice(3)) || 10; return `rgba(0,0,0,${Math.min(0.9, pct / 100)})`; }
  return '#e0e0e0';
}

const twipsPt = (w: wml.TblWidth | undefined): string | undefined => (w?.type === 'dxa' && typeof w.w === 'number' ? `${w.w / 20}pt` : undefined);

/**
 * The CSS variables a table's effective properties set, read by the `td` rules in editor.css. `firstCellTcPr` is the `tcPr`
 * of the first row's first cell, whose own left margin places a left-aligned table below compatibility mode 15.
 */
export function tableVariables(session: EditorDocument, tblPr: unknown, firstCellTcPr?: unknown): Record<string, string> {
  const resolver = session.pkg.propertyResolverOrUndefined;
  const direct = tblPr ? (revive(tblPr) as wml.CTTblPrBase) : undefined;
  let effective: wml.CTTblPrBase | undefined = direct;
  let cellDefaults: wml.TcPr | undefined;
  if (resolver) {
    try {
      const style = resolver.getEffectiveTableStyle(direct);
      effective = style.tblPr;
      cellDefaults = style.tcPr;
    } catch { effective = direct; }
  }
  const vars: Record<string, string> = {};
  const b = effective?.tblBorders;
  const set = (name: string, value: string | undefined): void => { if (value !== undefined) vars[name] = value; };
  set('--tbl-top', borderCss(b?.top, resolver));
  set('--tbl-bottom', borderCss(b?.bottom, resolver));
  set('--tbl-left', borderCss(b?.left, resolver));
  set('--tbl-right', borderCss(b?.right, resolver));
  set('--tbl-h', borderCss(b?.insideH, resolver));
  set('--tbl-v', borderCss(b?.insideV, resolver));
  const m = effective?.tblCellMar;
  set('--tbl-pad-top', twipsPt(m?.top));
  set('--tbl-pad-bottom', twipsPt(m?.bottom));
  set('--tbl-pad-left', twipsPt(m?.left ?? m?.start));
  set('--tbl-pad-right', twipsPt(m?.right ?? m?.end));
  set('--tbl-shading', shadingCss(cellDefaults?.shd, resolver));
  const jc = effective?.jc?.val as string | undefined;
  if (jc === 'center') vars['--tbl-margin'] = 'auto';
  else if (jc === 'right' || jc === 'end') vars['--tbl-margin-start'] = 'auto';
  else {
    // Where Word puts a left-aligned table's leading edge (measured in Word 365, ED-005 section 12.132 item 5,
    // fixtures/table-indent/). In compatibility mode 15 at `tblInd`, so a cell's text is its margin in from the body
    // text. Below it (14, and 12 where none is stated) at `tblInd` less the first row's first cell's left margin, its
    // own `tcMar` else the table's (108 twips when none is set), so that cell's text lines up with the body text and
    // the table stands out in the page margin; every row moves with the first. The wrapper carries the offset
    // (nodeviews), since a negative margin on the table itself would be clipped by the wrapper's overflow.
    const ind = effective?.tblInd?.type === 'dxa' && typeof effective.tblInd.w === 'number' ? effective.tblInd.w : 0;
    let pad = 0;
    if (session.compatibilityMode < 15) {
      const own = firstCellTcPr ? (revive(firstCellTcPr) as wml.TcPr).tcMar : undefined;
      const lead = own?.left ?? m?.left ?? m?.start;
      pad = lead?.type === 'dxa' && typeof lead.w === 'number' ? lead.w : 108;
    }
    vars['--tbl-offset'] = `${(ind - pad) / 20}pt`;
  }
  return vars;
}

/** A cell's properties as CSS declarations: borders, shading, vertical alignment, margins, text direction. */
export function cellDeclarations(pr: wml.TcPr, colors?: ColorResolver): string[] {
  const decl: string[] = [];
  const b = pr.tcBorders;
  const push = (prop: string, value: string | undefined): void => { if (value !== undefined) decl.push(`${prop}: ${value}`); };
  push('border-top', borderCss(b?.top, colors));
  push('border-bottom', borderCss(b?.bottom, colors));
  push('border-left', borderCss(b?.left ?? b?.start, colors));
  push('border-right', borderCss(b?.right ?? b?.end, colors));
  push('background-color', shadingCss(pr.shd, colors));
  const v = pr.vAlign?.val as string | undefined;
  if (v === 'center') decl.push('vertical-align: middle'); else if (v === 'bottom') decl.push('vertical-align: bottom'); else if (v === 'top') decl.push('vertical-align: top');
  const m = pr.tcMar;
  push('padding-top', twipsPt(m?.top));
  push('padding-bottom', twipsPt(m?.bottom));
  push('padding-left', twipsPt(m?.left));
  push('padding-right', twipsPt(m?.right));
  const dir = pr.textDirection?.val as string | undefined;
  if (dir === 'tbRl' || dir === 'tbRlV') decl.push('writing-mode: vertical-rl'); else if (dir === 'btLr') decl.push('writing-mode: vertical-lr');
  return decl;
}

/** A cell's direct properties as an inline style. */
export function cellInlineStyle(tcPr: unknown, colors?: ColorResolver): string {
  if (!tcPr) return '';
  return cellDeclarations(revive(tcPr) as wml.TcPr, colors).join('; ');
}
