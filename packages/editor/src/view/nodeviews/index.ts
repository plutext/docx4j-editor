// SPDX-License-Identifier: GPL-3.0-or-later
//
// Node and mark views (ED-001 section 2; ED-002 section 5.3): paragraphs carry their style class
// and direct formatting; images, opaque blocks and items, breaks, content controls and footnote
// references are read-only regions; marks carry direct run formatting as inline styles. The
// model's schema stays DOM-agnostic: everything visual is here.
import type { NodeViewConstructor, MarkViewConstructor, NodeView } from 'prosemirror-view';
import type { Node as PMNode, Mark } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { rowsOf, cellsOf, textOf, type Element, type ImagePart } from '@docx4j/core-ts';
import { isEmptyRun, tableAutoFitOf, columnWidthsAside, rowHeightOf, readControl, readTag, levelIndentOf, numberingParts, hyperlinkTarget, type Numbering, blipEmbedOf, type EditorDocument, type SdtKind, type SourceIds, type ControlLookup, type ParagraphAttrs, type Revision } from '@docx4j/editor-model';
import { TableView, updateColumnsOnResize } from 'prosemirror-tables';
import { cssId, paragraphInlineStyle, fontStack, highlightCss } from '../css.js';
import { INVISIBLE_KINDS } from '../plugins/invisibleAtoms.js';
import { t, tOr, formatDate } from '../../app/i18n.js';
import { cssString, type AuthorStyles } from '../authors.js';
import { allOn } from '../markup.js';
import { tableVariables, cellInlineStyle } from '../tables.js';
import { paragraphPagination, rowPagination, leadingHeaderRows, MAX_PRINT_HEADER_ROWS } from '../print.js';

/** Zero-width markers that can stand between blocks: drawn as a thin rule, not a labelled box. */
const BLOCK_MARKERS = new Set(['bookmarkStart', 'bookmarkEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd']);

/** The width `prosemirror-tables` gives a column with no width of its own. */
export const CELL_MIN_WIDTH = 25;

/**
 * `prosemirror-tables`' table view (a wrapper, the `<table>`, its `<colgroup>` kept in step with
 * the cells' `colwidth`), with the document's table properties as CSS variables on the table
 * (ED-003 section 3.8) and the grid's widths as the columns' when no cell states its own.
 */
class DocxTableView extends TableView {
  /** A print-only copy of the leading header rows (ED-006 section 3.4), there from `beforeprint` to `afterprint`. */
  private printHeader: HTMLTableSectionElement | null = null;
  private readonly onBeforePrint = (): void => this.addPrintHeader();
  private readonly onAfterPrint = (): void => this.removePrintHeader();

  constructor(node: PMNode, private readonly session: EditorDocument, private readonly authors?: AuthorStyles) {
    super(node, CELL_MIN_WIDTH);
    this.apply(node);
    const win = this.table.ownerDocument.defaultView;
    win?.addEventListener('beforeprint', this.onBeforePrint);
    win?.addEventListener('afterprint', this.onAfterPrint);
  }

  /**
   * The leading header rows copied into a `thead` beside the content (`prosemirror-tables` keeps every row
   * in one `tbody`, and a `tr` shown as a header group there is not repeated: section 9.1). Print shows the
   * copy on every page and hides the originals by their count (`print.css`).
   */
  private addPrintHeader(): void {
    this.removePrintHeader();
    const count = Math.min(leadingHeaderRows(this.node), MAX_PRINT_HEADER_ROWS);
    if (count === 0) return;
    const rows = Array.from((this.contentDOM as HTMLElement).children).slice(0, count);
    const thead = this.table.ownerDocument.createElement('thead');
    thead.className = 'print-header';
    thead.dataset['rows'] = String(count);
    for (const row of rows) thead.appendChild(row.cloneNode(true));
    this.table.insertBefore(thead, this.contentDOM as HTMLElement);
    this.printHeader = thead;
  }

  private removePrintHeader(): void {
    this.printHeader?.remove();
    this.printHeader = null;
  }

  override ignoreMutation(record: MutationRecord | { type: 'selection'; target: Node }): boolean {
    const header = this.printHeader;
    const nodes = record.type === 'childList' ? [...(record as MutationRecord).addedNodes, ...(record as MutationRecord).removedNodes] : [];
    if (header && (record.target === header || header.contains(record.target))) return true;
    if (nodes.length > 0 && nodes.every((n) => (n as HTMLElement).classList?.contains('print-header'))) return true;
    return super.ignoreMutation(record as MutationRecord);
  }

  destroy(): void {
    const win = this.table.ownerDocument.defaultView;
    win?.removeEventListener('beforeprint', this.onBeforePrint);
    win?.removeEventListener('afterprint', this.onAfterPrint);
    this.removePrintHeader();
  }

  private apply(node: PMNode): void {
    const vars = tableVariables(this.session, node.attrs['tblPr'], node.firstChild?.firstChild?.attrs['tcPr']);
    const style = Object.entries(vars).map(([k, v]) => `${k}: ${v}`).join('; ');
    if (style) this.table.setAttribute('style', style); else this.table.removeAttribute('style');
    // The table's leading edge as Word places it (tables.ts): the wrapper moves, so nothing is clipped.
    (this.dom as HTMLElement).style.marginInlineStart = vars['--tbl-offset'] ?? '';
    this.table.className = 'docx-table';
    // A table's properties or grid changed as a tracked change (shown in Word only): the change bar along it (ED-005 section 6.3).
    const change = (node.attrs['tblPr'] as { tblPrChange?: { author?: string } } | null)?.tblPrChange;
    const source = this.session.ctx.ids.get(node.attrs['sourceId'] as string)?.value as { tblGrid?: { tblGridChange?: unknown } } | undefined;
    (this.dom as HTMLElement).classList.toggle('has-change', !!change || !!source?.tblGrid?.tblGridChange);
    if (change) this.authors?.note(change.author);
    const styleId = (node.attrs['tblPr'] as { tblStyle?: { val?: string } } | null)?.tblStyle?.val;
    if (styleId) this.table.classList.add(`ts-${cssId(styleId)}`);
    updateColumnsOnResize(node, this.colgroup, this.table, CELL_MIN_WIDTH);
    // A grid width for a column whose cells state none.
    const grid = node.attrs['tblGrid'] as number[] | null;
    if (grid) {
      Array.from(this.colgroup.children).forEach((col, i) => {
        const el = col as HTMLElement;
        if (!el.style.width && grid[i]) el.style.width = `${grid[i]! / 20}pt`;
      });
    }
    // Word's AutoFit (ED-005 sections 12.78 and 12.83): a table whose width is a percentage is drawn at it, its
    // columns scaled by the browser from the widths they state; an automatic table takes the browser's automatic
    // layout, its stated widths the columns' minimums (Word's default table, check 42) or, where its cells are
    // automatic, put aside (AutoFit to contents); fixed widths are the sum, past the margin if it must be.
    const mode = tableAutoFitOf(node.attrs['tblPr']);
    const pct = (node.attrs['tblPr'] as { tblW?: { type?: string; w?: number } } | null)?.tblW;
    this.table.classList.toggle('autofit', mode === 'contents');
    this.table.classList.toggle('autofit-contents', mode === 'contents' && columnWidthsAside(node));
    if (mode === 'window') {
      this.table.style.width = `${(pct?.w ?? 5000) / 50}%`;
      this.table.style.minWidth = '';
    } else if (mode === 'contents') {
      this.table.style.width = '';
      this.table.style.minWidth = '';
    }
  }

  override update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    const ok = super.update(node);
    if (ok) this.apply(node);
    return ok;
  }
}

export type RevokeRegistrar = (fn: () => void) => void;

const UNDERLINE_CSS: Record<string, string> = {
  Single: 'underline', Double: 'underline double', Dotted: 'underline dotted', Dashed: 'underline dashed', Wave: 'underline wavy', Thick: 'underline', Words: 'underline',
};

/** A control kind's name in the application's language (`control.kind.<kind>`), the kind itself for one no file names. */
export function kindLabelOf(kind: SdtKind | string): string {
  return tOr(`control.kind.${kind}`, kind);
}

/** How a control is drawn (ED-003 section 4.9): its label (the title, else the tag, else the kind), its role for the colour, and the tooltip. */
export function controlDisplay(sdtPr: unknown): { label: string; role: string; kind: string; title: string; appearance: string; name: string; color: string } {
  const reading = readControl(sdtPr);
  const tag = readTag(reading.tag);
  const kind = kindLabelOf(reading.kind);
  const role = tag.role !== 'none' ? tag.role : reading.binding || tag.xpath !== undefined ? 'bound' : tag.positionCondition !== undefined ? 'positionCondition' : 'none';
  const label = reading.title || reading.tag || kind;
  const title = [reading.title ? t('control.tooltip.title', { title: reading.title }) : '', reading.tag ? t('control.tooltip.tag', { tag: reading.tag }) : '', kind].filter(Boolean).join('  ');
  return { label, role, kind: reading.kind, title, appearance: reading.appearance, name: reading.title, color: reading.color };
}

/**
 * The control's own look as Word draws it (ED-005 section 12.60): its appearance (`w15:appearance`;
 * a bounding box when it states none), its title for the box's tab and the tags, and its colour
 * (`w15:color`) as `--sdt-own`. The Template perspective's stylesheet draws by role instead.
 */
function applyAppearance(dom: HTMLElement, d: ReturnType<typeof controlDisplay>, prefix: '' | 'sdt' = ''): void {
  const key = (k: string): string => (prefix ? prefix + k[0]!.toUpperCase() + k.slice(1) : k);
  dom.dataset[key('appearance')] = d.appearance || 'BoundingBox';
  if (d.name) dom.dataset[key('name')] = d.name; else delete dom.dataset[key('name')];
  if (/^#[0-9a-f]{6}$/i.test(d.color)) dom.style.setProperty('--sdt-own', d.color); else dom.style.removeProperty('--sdt-own');
}

/** A block or run control's element: the label, role and kind the stylesheet draws from, its own look, and the tooltip. */
function applyControl(dom: HTMLElement, node: PMNode): void {
  const d = controlDisplay(node.attrs['sdtPr']);
  dom.dataset['label'] = d.label;
  dom.dataset['role'] = d.role;
  dom.dataset['kind'] = d.kind;
  applyAppearance(dom, d);
  dom.title = d.title;
}

/** A part whose relationships resolve the content on screen: the main part, or a header's own. */
export interface RelationshipsOwner {
  relationshipsPart?: { getPart(id: string): unknown } | undefined;
}

export interface NodeViewOptions {
  /** The part the content came from, for `r:embed` and friends; the main document part by default. */
  part?: RelationshipsOwner;
  /**
   * Field results known for the content on screen, by instruction keyword (`PAGE`, `NUMPAGES`):
   * a header at a page boundary shows the page's number rather than the cached result.
   */
  fields?: Record<string, string>;
  /** The ids the content was projected with, when not the session's (the preview projects a clone into a context of its own). */
  ids?: SourceIds;
  /** The row and cell groups' properties, when not the session's controls map (the preview's own projection). */
  controls?: ControlLookup;
  /** The body view's rule per author (ED-005 section 6.3); a story has no revisions to colour. */
  authors?: AuthorStyles;
  /** The list definitions a list paragraph's indent is read from, when not the session's (the preview's clone). */
  numbering?: Numbering;
}

/** The keyword of a field instruction: `PAGE` from ` PAGE \* MERGEFORMAT `. */
export function fieldKeyword(instr: string | undefined): string | undefined {
  return instr?.trim().split(/\s+/)[0]?.toUpperCase() || undefined;
}

/** The image part an image node's source embeds (an inline drawing, or a kept `mc:AlternateContent` holding one), through `a:blip r:embed` and the owning part's relationships. */
function imagePartOf(owner: RelationshipsOwner, drawing: Element | undefined): ImagePart | undefined {
  const embed = blipEmbedOf(drawing);
  if (!embed) return undefined;
  return owner.relationshipsPart?.getPart(embed) as ImagePart | undefined;
}

/** The side of a revision an atom's, a row's or a mark's DOM shows, and the author whose colour it takes (ED-005 section 6.3). */
function applyRevision(dom: HTMLElement, r: Revision | null | undefined, authors: AuthorStyles | undefined, prefix = 'rev'): void {
  dom.classList.toggle(`${prefix}-ins`, !!r?.insertion);
  dom.classList.toggle(`${prefix}-del`, !!r?.deletion);
  dom.classList.toggle(`${prefix}-move`, !!(r?.insertion?.move || r?.deletion?.move));
  const who = r?.deletion ?? r?.insertion;
  if (who) { dom.dataset['author'] = who.author ?? ''; authors?.note(who.author); }
  else delete dom.dataset['author'];
}

const REVISION_MARKS = new Set(['insertion', 'deletion', 'formatChange']);

/** Whether a paragraph holds a tracked change of any kind, preserved ones included: the change bar's class, read from the node alone. */
function holdsChange(n: PMNode): boolean {
  const a = n.attrs as ParagraphAttrs & { sectPr: { sectPrChange?: unknown } | null; markRPr: { rPrChange?: unknown } | null };
  if (a.markRevision || a.pPrChange || a.sectPr?.sectPrChange || a.markRPr?.rPrChange) return true;
  let found = false;
  n.descendants((child) => {
    if (found) return false;
    if (child.isText ? child.marks.some((m) => REVISION_MARKS.has(m.type.name)) : !!child.attrs['revision'] && !isEmptyRun(child)) found = true;
    return !found;
  });
  return found;
}

/**
 * A paragraph's classes for tracked changes (ED-005 sections 6.3 and 6.4): `has-change` for the
 * bar; `rev-mark-ins`, `rev-mark-del` for its mark's pilcrow and the views' joins; `rev-all-ins`,
 * `rev-all-del` when every item of it is inserted, or deleted, so a view can hide it whole.
 */
function applyParagraphRevision(dom: HTMLElement, n: PMNode, authors: AuthorStyles | undefined): void {
  const r = n.attrs['markRevision'] as Revision | null;
  dom.classList.toggle('has-change', holdsChange(n));
  dom.classList.toggle('rev-props', !!n.attrs['pPrChange']);
  applyRevision(dom, r, authors, 'rev-mark');
  dom.classList.toggle('rev-all-del', !!r?.deletion && allOn(n, 'deletion'));
  dom.classList.toggle('rev-all-ins', !!r?.insertion && allOn(n, 'insertion'));
}

/** Whether a row, its cells or its paragraphs hold a tracked change: the row's bar, at the table's edge. */
function rowHoldsChange(n: PMNode): boolean {
  if (n.attrs['revision'] || (n.attrs['trPr'] as { trPrChange?: unknown } | null)?.trPrChange) return true;
  let found = false;
  n.descendants((child) => {
    if (found) return false;
    if (child.type.name === 'table_cell') {
      const tcPr = child.attrs['tcPr'] as Record<string, unknown> | null;
      if (tcPr?.['tcPrChange'] || tcPr?.['cellIns'] || tcPr?.['cellDel'] || tcPr?.['cellMerge']) found = true;
      return !found;
    }
    if (child.type.name === 'paragraph') { if (holdsChange(child)) found = true; return false; }
    return !found;
  });
  return found;
}

export function nodeViews(session: EditorDocument, onRevoke: RevokeRegistrar, options: NodeViewOptions = {}): Partial<Record<string, NodeViewConstructor>> {
  const owner: RelationshipsOwner = options.part ?? session.main;
  const ids = options.ids ?? session.ctx.ids;
  const controls = options.controls ?? session.controlEntry;
  const authors = options.authors;
  const urls = new Map<string, Promise<string>>();
  const imageUrl = (part: ImagePart): Promise<string> => {
    const key = part.partName.toString();
    let p = urls.get(key);
    if (!p) {
      p = part.getBytes().then((bytes) => {
        const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: part.contentType }));
        onRevoke(() => URL.revokeObjectURL(url));
        return url;
      });
      urls.set(key, p);
    }
    return p;
  };

  return {
    paragraph(node): NodeView {
      const dom = document.createElement('p');
      // A list paragraph lays out at its level's indent (ED-003 section 5.4), read through the
      // session's numbering definitions; recomputed only when the style, the list or the indent changed.
      let indentKey = '';
      let indent: ReturnType<typeof levelIndentOf> = null;
      let styleClass = '';
      // The paragraph's own breaks and keeps, from its source element (or the one its properties were copied
      // from, Enter's first half; ED-005 section 12.102), for printing (ED-006 section 3.4): read when the node
      // comes from another source, never per keystroke.
      let paginationSource: string | null | undefined;
      const apply = (n: PMNode): void => {
        const sourceId = (n.attrs['sourceId'] ?? n.attrs['copiedFrom']) as string | null;
        if (sourceId !== paginationSource) {
          paginationSource = sourceId;
          const pPr = sourceId ? (session.ctx.ids.get(sourceId)?.value as wml.P | undefined)?.pPr : undefined;
          const attrs = paragraphPagination(pPr);
          for (const key of ['pageBreakBefore', 'keepNext', 'keepLines', 'widowControl']) {
            if (attrs[key] !== undefined) dom.dataset[key] = attrs[key]; else delete dom.dataset[key];
          }
        }
        const style = n.attrs['pStyle'] as string | null;
        // Its own classes toggled, never the list replaced: a node decoration's class (a section's
        // `sect-N`) is not put back by ProseMirror when the decoration itself is unchanged.
        const next = style ? `s-${cssId(style)}` : '';
        if (next !== styleClass) { if (styleClass) dom.classList.remove(styleClass); if (next) dom.classList.add(next); styleClass = next; }
        applyParagraphRevision(dom, n, authors);
        const key = JSON.stringify([style, n.attrs['numPr'], n.attrs['ind']]);
        if (key !== indentKey) { indentKey = key; indent = levelIndentOf(n.attrs as ParagraphAttrs, options.numbering ?? numberingParts(session)); }
        const inline = paragraphInlineStyle(n.attrs, indent, session.pkg.propertyResolverOrUndefined);
        if (inline) dom.setAttribute('style', inline); else dom.removeAttribute('style');
      };
      apply(node);
      return {
        dom,
        contentDOM: dom,
        update(n) {
          if (n.type !== node.type) return false;
          apply(n);
          return true;
        },
      };
    },

    table(node): NodeView {
      return new DocxTableView(node, session, authors);
    },

    table_row(node): NodeView {
      const dom = document.createElement('tr');
      const apply = (n: PMNode): void => {
        // Rows inside row-level controls: the groups' labels, outermost first, as the gutter's label
        // (ED-001 section 4.1's display), read from the controls map so an edit of the properties shows.
        const stack = n.attrs['sdt'] as { groupId: string; first: boolean; last: boolean }[] | null;
        if (stack && stack.length > 0) {
          const displays = stack.map((g) => controlDisplay(controls(g.groupId)?.sdtPr ?? null));
          const label = displays.map((d) => d.label).join(' > ');
          dom.dataset['sdt'] = label;
          dom.dataset['sdtDepth'] = String(stack.length);
          dom.dataset['sdtRole'] = displays[0]!.role;
          applyAppearance(dom, displays[0]!, 'sdt');
          dom.title = displays.map((d) => d.title).join(' > ');
          dom.classList.toggle('sdt-first', stack[0]!.first);
          dom.classList.toggle('sdt-last', stack[0]!.last);
          // The label the first cell's badge shows (editor.css): a pseudo-element on the row itself would lay out as a cell.
          dom.style.setProperty('--sdt-row-label', cssString(label));
          dom.style.setProperty('--sdt-row-name', cssString(displays[0]!.name));
        } else {
          delete dom.dataset['sdt'];
          delete dom.dataset['sdtDepth'];
          delete dom.dataset['sdtRole'];
          delete dom.dataset['sdtAppearance'];
          delete dom.dataset['sdtName'];
          dom.style.removeProperty('--sdt-own');
          dom.removeAttribute('title');
          dom.classList.remove('sdt-first', 'sdt-last');
          dom.style.removeProperty('--sdt-row-label');
          dom.style.removeProperty('--sdt-row-name');
        }
        // A row inserted or deleted as a tracked change, tinted; any change in it, the bar at the table's edge (ED-005 section 6.3).
        applyRevision(dom, n.attrs['revision'] as Revision | null, authors, 'rev-row');
        dom.classList.toggle('has-change', rowHoldsChange(n));
        // For printing (ED-006 section 3.4): a row that cannot split, and a header row.
        const pagination = rowPagination(n.attrs['trPr']);
        if (pagination.cantSplit) dom.dataset['cantSplit'] = '1'; else delete dom.dataset['cantSplit'];
        if (pagination.header) dom.dataset['header'] = '1'; else delete dom.dataset['header'];
        // A row's stated height (`w:trHeight`, twips): a minimum is the row's CSS height, which a table row cannot go under;
        // an exact height is drawn the same way, since a browser's table row grows with its content (docs/developer/word-divergences.md).
        const height = rowHeightOf(n.attrs['trPr']);
        if (height?.val) dom.style.height = `${height.val / 15}px`; else dom.style.removeProperty('height');
      };
      apply(node);
      // Called for a change of the node or of its decorations: the current-control mark carries the properties (plugins/controls.ts).
      return { dom, contentDOM: dom, update: (n) => { if (n.type !== node.type) return false; apply(n); return true; } };
    },

    table_cell(node): NodeView {
      const dom = document.createElement('td');
      const apply = (n: PMNode): void => {
        const colspan = n.attrs['colspan'] as number;
        const rowspan = n.attrs['rowspan'] as number;
        if (colspan !== 1) dom.setAttribute('colspan', String(colspan)); else dom.removeAttribute('colspan');
        if (rowspan !== 1) dom.setAttribute('rowspan', String(rowspan)); else dom.removeAttribute('rowspan');
        const colwidth = n.attrs['colwidth'] as number[] | null;
        if (colwidth) dom.setAttribute('data-colwidth', colwidth.join(',')); else dom.removeAttribute('data-colwidth');
        const style = cellInlineStyle(n.attrs['tcPr'], session.pkg.propertyResolverOrUndefined);
        if (style) dom.setAttribute('style', style); else dom.removeAttribute('style');
        const stack = n.attrs['sdt'] as { groupId: string }[] | null;
        if (stack && stack.length > 0) {
          const displays = stack.map((g) => controlDisplay(controls(g.groupId)?.sdtPr ?? null));
          dom.dataset['sdt'] = String(stack.length);
          dom.dataset['sdtLabel'] = displays.map((d) => d.label).join(' > ');
          dom.dataset['sdtRole'] = displays[displays.length - 1]!.role;
          applyAppearance(dom, displays[displays.length - 1]!, 'sdt');
          dom.title = displays.map((d) => d.title).join(' > ');
        } else {
          delete dom.dataset['sdt'];
          delete dom.dataset['sdtLabel'];
          delete dom.dataset['sdtRole'];
          delete dom.dataset['sdtAppearance'];
          delete dom.dataset['sdtName'];
          dom.style.removeProperty('--sdt-own');
          dom.removeAttribute('title');
        }
      };
      apply(node);
      return { dom, contentDOM: dom, update: (n) => { if (n.type !== node.type) return false; apply(n); return true; } };
    },

    image(node): NodeView {
      const dom = document.createElement('img');
      dom.className = 'image';
      dom.alt = '';
      const w = node.attrs['width'] as number | null;
      const h = node.attrs['height'] as number | null;
      if (w) dom.style.inlineSize = `${w}pt`;
      if (h) dom.style.blockSize = `${h}pt`;
      const id = node.attrs['sourceId'] as string | null;
      const part = id ? imagePartOf(owner, ids.get(id)) : undefined;
      if (part && typeof URL.createObjectURL === 'function') void imageUrl(part).then((url) => { dom.src = url; });
      else dom.classList.add('missing');
      applyRevision(dom, node.attrs['revision'] as Revision | null, authors);
      return { dom, ignoreMutation: () => true };
    },

    opaqueBlock(node): NodeView {
      const kind = String(node.attrs['kind']);
      const id = node.attrs['sourceId'] as string | null;
      const el = id ? ids.get(id) : undefined;
      const dom = document.createElement('div');
      dom.className = `opaque-block kind-${kind}`;
      dom.contentEditable = 'false';
      dom.dataset['kind'] = kind;
      if (BLOCK_MARKERS.has(kind)) {
        // A marker between blocks (a bookmark, a move's or a comment's range): a thin rule, its kind as the tooltip (ED-005 section 6.3).
        dom.classList.add('marker');
        if (kind.startsWith('move')) dom.classList.add('move-marker');
        dom.title = kind;
        return { dom, ignoreMutation: () => true };
      }
      if (kind === 'tbl' && el) {
        const table = document.createElement('table');
        table.className = 'opaque-table';
        for (const row of rowsOf(el.value as object)) {
          const tr = document.createElement('tr');
          for (const cell of cellsOf(row.element.value as object)) {
            const td = document.createElement('td');
            td.textContent = textOf(cell.element);
            tr.appendChild(td);
          }
          table.appendChild(tr);
        }
        dom.appendChild(table);
      } else {
        const label = document.createElement('span');
        label.className = 'opaque-label';
        label.textContent = kind;
        dom.appendChild(label);
        const text = el ? textOf(el) : '';
        if (text) { const body = document.createElement('div'); body.textContent = text; dom.appendChild(body); }
      }
      return { dom, ignoreMutation: () => true };
    },

    opaqueInline(node): NodeView {
      const kind = String(node.attrs['kind']);
      const dom = document.createElement('span');
      dom.className = `opaque-inline kind-${kind}${INVISIBLE_KINDS.has(kind) ? ' invisible' : ''}`;
      dom.contentEditable = 'false';
      dom.dataset['kind'] = kind;
      dom.title = kind;
      if (kind === 'commentReference') dom.textContent = '\u{1F4AC}';
      else if (kind === 'fldSimple') {
        // A simple field shows its cached result (ED-002 section 7.4), or the value the view knows.
        const id = node.attrs['sourceId'] as string | null;
        const el = id ? ids.get(id) : undefined;
        const instr = (el?.value as wml.CTSimpleField | undefined)?.instr;
        const keyword = fieldKeyword(instr);
        const known = keyword ? options.fields?.[keyword] : undefined;
        dom.textContent = known ?? (el ? textOf(el) : '') ?? '';
        dom.classList.add('field');
        if (instr) dom.title = t('view.field', { instruction: instr.trim() });
        if (dom.textContent === '') dom.textContent = '⬚';
      }
      else if (!INVISIBLE_KINDS.has(kind)) dom.textContent = '⬚';
      applyRevision(dom, node.attrs['revision'] as Revision | null, authors);
      return { dom, ignoreMutation: () => true };
    },

    hardBreak(node): NodeView {
      const kind = String(node.attrs['kind']);
      if (kind === 'line') {
        const br = document.createElement('br');
        applyRevision(br, node.attrs['revision'] as Revision | null, authors);
        return { dom: br };
      }
      const dom = document.createElement('span');
      dom.className = `page-break kind-${kind}`;
      dom.contentEditable = 'false';
      dom.textContent = kind === 'column' ? t('view.columnBreak') : t('view.pageBreak');
      applyRevision(dom, node.attrs['revision'] as Revision | null, authors);
      return { dom, ignoreMutation: () => true };
    },

    sdtBlock(node): NodeView {
      const dom = document.createElement('div');
      dom.className = 'sdt-block';
      applyControl(dom, node);
      return { dom, contentDOM: dom, update: (n) => { if (n.type !== node.type) return false; applyControl(dom, n); return true; } };
    },

    sdtInline(node): NodeView {
      const dom = document.createElement('span');
      dom.className = 'sdt-inline';
      applyControl(dom, node);
      return { dom, contentDOM: dom, update: (n) => { if (n.type !== node.type) return false; applyControl(dom, n); return true; } };
    },

    footnoteRef(node): NodeView {
      const dom = document.createElement('sup');
      dom.className = `footnote-ref ${String(node.attrs['kind'])}`;
      dom.contentEditable = 'false';
      dom.title = t(node.attrs['kind'] === 'endnote' ? 'view.noteRef.endnote' : 'view.noteRef.footnote', { id: String(node.attrs['footnoteId']) });
      applyRevision(dom, node.attrs['revision'] as Revision | null, authors);
      return { dom, ignoreMutation: () => true };
    },
  };
}

export function markViews(session?: EditorDocument, authors?: AuthorStyles): Partial<Record<string, MarkViewConstructor>> {
  const span = (className: string, style: string, tag = 'span') => {
    const dom = document.createElement(tag);
    if (className) dom.className = className;
    if (style) dom.setAttribute('style', style);
    return { dom };
  };
  /**
   * A tracked change's mark, the element the schema's `toDOM` makes (ED-005 section 6.3), its
   * author's rule written once, and the author and the date as its tooltip (no words of its own).
   */
  const revision = (tag: string, className: string) => (mark: Mark) => {
    const dom = document.createElement(tag);
    const author = mark.attrs['author'] as string | null;
    const date = mark.attrs['date'] as string | null;
    dom.className = mark.attrs['move'] ? `${className} move` : className;
    dom.dataset['author'] = author ?? '';
    authors?.note(author);
    const title = [author ?? '', date ? formatDate(date) : ''].filter(Boolean).join(', ');
    if (title) dom.title = title;
    return { dom };
  };
  return {
    insertion: revision('ins', 'insertion'),
    deletion: revision('del', 'deletion'),
    formatChange: revision('span', 'format-change'),
    /** A hyperlink: the address (from its relationship) or the bookmark as the tooltip; Ctrl+click follows it (plugins/links.ts). */
    link(mark: Mark) {
      const dom = document.createElement('a');
      dom.className = 'link';
      const address = hyperlinkTarget(session, mark.attrs['rId'] as string | null);
      const anchor = mark.attrs['anchor'] as string | null;
      const target = address ? (anchor ? `${address}#${anchor}` : address) : anchor ? `#${anchor}` : '';
      if (target) { dom.title = t('view.link.follow', { target }); dom.dataset['target'] = target; }
      return { dom };
    },
    font(mark: Mark) {
      const decl: string[] = [];
      if (mark.attrs['name']) decl.push(`font-family: ${fontStack(mark.attrs['name'] as string)}`);
      if (mark.attrs['size']) decl.push(`font-size: ${mark.attrs['size'] as number}pt`);
      const color = mark.attrs['color'] as string | null;
      if (color && color !== 'auto' && color.startsWith('#')) decl.push(`color: ${color}`);
      if (mark.attrs['highlight']) decl.push(`background-color: ${highlightCss(mark.attrs['highlight'] as string)}`);
      return span('font', decl.join('; '));
    },
    runStyle(mark: Mark) {
      return span(`r-${cssId(String(mark.attrs['id']))}`, '');
    },
    underline(mark: Mark) {
      return span('underline', `text-decoration: ${UNDERLINE_CSS[String(mark.attrs['val'])] ?? 'underline'}`, 'u');
    },
    strike(mark: Mark) {
      return span('strike', mark.attrs['double'] ? 'text-decoration: line-through double' : '', 's');
    },
    vertAlign(mark: Mark) {
      return span('', '', mark.attrs['val'] === 'subscript' ? 'sub' : 'sup');
    },
    rtl() {
      return span('rtl', 'direction: rtl; unicode-bidi: isolate', 'bdi');
    },
  };
}
