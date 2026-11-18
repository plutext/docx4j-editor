// SPDX-License-Identifier: GPL-3.0-or-later
//
// The ProseMirror schema of the editable projection (ED-002 section 3.3; ED-001 section 3).
// Every attribute holds direct formatting or a reference into the document layer, never
// effective formatting; `sourceId` is the id of the element a node came from (ids.mts), and the
// `runSource` mark is the same for the run a piece of inline content came from.
//
// Attribute defaults are `null`, not `undefined`: y-prosemirror drops null attributes when it
// writes a node into Yjs and restores the default when it reads it back, so a null default is
// what survives the round trip unchanged.
//
// Inline atoms (tab, break, image, footnote reference, opaque item) and the inline content
// control are created with NO marks: y-prosemirror stores marks only on text, so a mark on an
// element node never reaches Yjs and is lost on the first rebuild from the store
// (docs/developer/upstream-limitations.md item 2). (A `marks` entry in a NodeSpec would restrict the marks
// inside the node's content, which is not what is wanted, so this is a rule for the code, not
// the schema.) An atom carries the run it came from as the `runSource`
// attribute and, when it was created in the editor, its formatting as the `fmt` attribute
// (the string `formattingKey` gives for a set of marks); export groups it into a run from those.
import { Schema, type NodeSpec, type MarkSpec, type DOMOutputSpec } from 'prosemirror-model';

/** The attributes every inline atom shares (UPSTREAM-2). */
const atomAttrs = {
  /** The run this atom came from (ids.mts); null when created in the editor. */
  runSource: { default: null },
  /**
   * The kept `mc:AlternateContent` inside that run whose selected branch this atom came from
   * (ED-009): the element's id, or null for an item of the run itself. The export writes the
   * branch's content back into the element and the other branches as they were.
   */
  alt: { default: null },
  /** The formatting of an atom created in the editor, as `formattingKey` writes it; null means "the source run's". */
  fmt: { default: null },
  /**
   * The atom's run inserted or deleted as a tracked change (ED-005 section 5.2's last row): a
   * `Revision`, since an atom carries no marks through Yjs (UPSTREAM-2); null when it is neither.
   */
  revision: { default: null },
};

/**
 * One side of a tracked change (ED-005 section 5.2): who made it and when, and for a move the
 * move's name (`w:name` of its range, the same on both halves; '' for a move with no named range).
 * `date` is Word's `w:date` as an ISO string (`2026-09-28T05:36:00Z`), or null when it has none.
 * The mark attributes are these primitives, so a mark compares by value and Yjs keeps them flat.
 */
export interface RevisionInfo {
  author: string | null;
  date: string | null;
  move: string | null;
}

/**
 * A paragraph mark's, a row's or an atom's tracked change: an insertion, a deletion, or both
 * (another author's insertion deleted), as `w:ins` and `w:del` can both sit in a `w:rPr` or a
 * `w:trPr`. A move is an insertion (`w:moveTo`) or a deletion (`w:moveFrom`) whose `move` is set.
 */
export interface Revision {
  insertion?: RevisionInfo;
  deletion?: RevisionInfo;
}

/** A paragraph's properties change (`w:pPrChange`): who and when, and the properties before it as plain JSON (`toPlain` of the old `w:pPr`). */
export interface PropertiesChange {
  author: string | null;
  date: string | null;
  before: unknown;
}

/** The editable fields of `w:ind`, in twips; `start`/`end` read the transitional `left`/`right` as aliases. */
export interface IndentAttrs {
  start: number | null;
  end: number | null;
  firstLine: number | null;
  hanging: number | null;
}

/** The editable fields of `w:spacing` on a paragraph, in twips (`line` per `lineRule`). */
export interface SpacingAttrs {
  before: number | null;
  after: number | null;
  line: number | null;
  lineRule: string | null;
}

/**
 * The paragraph's own `w:numPr` (ED-003 section 5.2): `numId` as docx4j's string key (`'0'` is
 * legal and means "no list, whatever the style says"; null where the element states only a
 * level, which takes the style's list at that level), `ilvl` 0 where nothing states one. A
 * style's numbering is effective, not an attribute; the emulator is asked for it.
 */
export interface NumPrAttrs {
  numId: string | null;
  ilvl: number;
}

/**
 * One stop of the paragraph's own `w:tabs` (ED-003 section 8.5, step F4): `pos` in twips, `val`
 * as the element states it (`left`, `center`, `right`, `decimal`, `bar`, `num`, `start`, `end`,
 * and `clear`, which cancels a style's stop at that position and is the paragraph's statement
 * as much as the others), `leader` as stated or null. A style's stops are effective, not an
 * attribute; `paragraphTabs` asks the resolver for the merged list.
 */
export interface TabAttrs {
  pos: number;
  val: string;
  leader: string | null;
}

export interface ParagraphAttrs {
  sourceId: string | null;
  pStyle: string | null;
  jc: string | null;
  ind: IndentAttrs | null;
  spacing: SpacingAttrs | null;
  bidi: boolean;
  numPr: NumPrAttrs | null;
  /** The paragraph's own `w:tabs`, in document order; null when it states none. */
  tabs: TabAttrs[] | null;
  /** The paragraph's `w:sectPr` as plain JSON (PARENT stripped, TYPE_NAME kept), when it ends a section. */
  sectPr: unknown | null;
  /** The paragraph mark's `w:rPr` as plain JSON, less its tracked change, which is `markRevision`. */
  markRPr: unknown | null;
  /** The paragraph mark inserted or deleted as a tracked change (ED-005 section 5.2), read out of the mark's `w:rPr`. */
  markRevision: Revision | null;
  /** `w:pPrChange`: the paragraph's properties before a tracked change to them. */
  pPrChange: PropertiesChange | null;
  /**
   * The source id of the paragraph this one was made with a copy of (Enter's first half; ED-005 section
   * 12.102): the export takes that element's `w:pPr` as its base, so the properties outside the editable
   * subset (a border, a keep) survive as Word keeps them on the new mark. Session-local, like `sourceId`.
   */
  copiedFrom: string | null;
}

export interface FontAttrs {
  name: string | null;
  /** Points. */
  size: number | null;
  /** '#RRGGBB', 'auto' or a theme colour name, as `readRunOptions` reports it. */
  color: string | null;
  /** '#RRGGBB' of a highlight colour. */
  highlight: string | null;
}

const paragraphDOM: DOMOutputSpec = ['p', 0];

export const nodes: Record<string, NodeSpec> = {
  doc: {
    content: 'block+',
    attrs: {
      /** The body's final `w:sectPr` as plain JSON; never removed (ED-001 section 3.1). */
      finalSectPr: { default: null },
    },
  },

  paragraph: {
    group: 'block',
    content: 'inline*',
    attrs: {
      sourceId: { default: null },
      pStyle: { default: null },
      jc: { default: null },
      ind: { default: null },
      spacing: { default: null },
      bidi: { default: false },
      numPr: { default: null },
      tabs: { default: null },
      sectPr: { default: null },
      markRPr: { default: null },
      markRevision: { default: null },
      pPrChange: { default: null },
      copiedFrom: { default: null },
    },
    parseDOM: [{ tag: 'p' }],
    toDOM: () => paragraphDOM,
  },

  /** A block-level `w:sdt`: content editable in E1, the control itself not (ED-002 section 8 item 7). */
  sdtBlock: {
    group: 'block',
    content: 'block+',
    defining: true,
    attrs: {
      sourceId: { default: null },
      /** `w:sdtPr` as plain JSON. */
      sdtPr: { default: null },
    },
    toDOM: () => ['div', { class: 'sdt-block' }, 0],
  },

  /**
   * `w:tbl` (ED-003 section 3.2): the shape `prosemirror-tables` expects (`tableRole`, and cells
   * with `colspan`, `rowspan` and `colwidth`), so its map, selections and commands apply
   * unchanged. `tblPr` and `tblGrid` ride on the copy at export; the editable subset is the
   * style id and the grid widths. A control around a whole table is `sdtBlock` containing `table`.
   */
  table: {
    group: 'block',
    content: 'table_row+',
    tableRole: 'table',
    isolating: true,
    attrs: {
      sourceId: { default: null },
      /** `w:tblPr` as plain JSON. */
      tblPr: { default: null },
      /** `w:tblGrid`'s column widths in twips. */
      tblGrid: { default: null },
    },
    toDOM: () => ['table', ['tbody', 0]],
  },

  /**
   * `w:tr`. A row inside row-level content controls carries them as the `sdt` stack (ED-001
   * section 4.1 option 2, decision 4): `[{ groupId, first, last }]`, outermost first, the group's
   * element in the ids map and its `sdtPr` in the `controls` map (E2.b). Content `table_cell*`,
   * as `prosemirror-tables` allows: a row whose cells are all continuations of merges above has
   * no cell nodes of its own.
   */
  table_row: {
    content: 'table_cell*',
    tableRole: 'row',
    attrs: {
      sourceId: { default: null },
      /** `w:trPr` as plain JSON, less its tracked insertion or deletion, which is `revision`. */
      trPr: { default: null },
      sdt: { default: null },
      /** The row inserted or deleted as a tracked change (`w:trPr/w:ins`, `w:del`; ED-005 section 5.2): a `Revision`. */
      revision: { default: null },
    },
    toDOM: () => ['tr', 0],
  },

  /**
   * `w:tc`. `colspan` is `w:gridSpan`; `rowspan` folds a `w:vMerge` run into its restart cell,
   * whose continuation cells are the `merged` side data (their source ids) and are re-emitted
   * by identity when the cell is unchanged. `colwidth` is in CSS pixels, one per spanned column,
   * the unit `prosemirror-tables`' resizing writes; export converts to twips (15 per pixel).
   */
  table_cell: {
    content: 'block+',
    tableRole: 'cell',
    isolating: true,
    attrs: {
      sourceId: { default: null },
      colspan: { default: 1 },
      rowspan: { default: 1 },
      colwidth: { default: null },
      /** `w:tcPr` as plain JSON, `gridSpan`, `vMerge` and `tcW` included as imported. */
      tcPr: { default: null },
      /** The cell-level control stack, as a row's. */
      sdt: { default: null },
      /** Source ids of the `w:tc` continuation cells this cell's `rowspan` folds in. */
      merged: { default: null },
    },
    toDOM: (node) => {
      const attrs: Record<string, string> = {};
      if ((node.attrs['colspan'] as number) !== 1) attrs['colspan'] = String(node.attrs['colspan']);
      if ((node.attrs['rowspan'] as number) !== 1) attrs['rowspan'] = String(node.attrs['rowspan']);
      const colwidth = node.attrs['colwidth'] as number[] | null;
      if (colwidth) attrs['data-colwidth'] = colwidth.join(',');
      return ['td', attrs, 0];
    },
  },

  /**
   * A block the editor does not edit: block-level tracked changes, math, altChunk,
   * bookmarks and permissions at block level, anything unknown. Emits its original element on
   * export; reported in diagnostics (ED-001 section 11 item 8).
   */
  opaqueBlock: {
    group: 'block',
    atom: true,
    selectable: true,
    attrs: {
      sourceId: { default: null },
      /** The element's local name (`tbl`, `bookmarkStart`, `oMathPara`, ...). */
      kind: { default: 'unknown' },
      snapshotHtml: { default: null },
    },
    toDOM: (node) => ['div', { class: 'opaque-block', 'data-kind': String(node.attrs['kind']) }],
  },

  text: {
    group: 'inline',
  },

  /** `w:br` (kind 'line', 'page' or 'column') and `w:cr` (kind 'line'). */
  hardBreak: {
    group: 'inline',
    inline: true,
    selectable: false,
    attrs: {
      ...atomAttrs,
      kind: { default: 'line' },
      clear: { default: null },
    },
    toDOM: () => ['br'],
  },

  /** `w:tab`. */
  tab: {
    group: 'inline',
    inline: true,
    selectable: false,
    attrs: { ...atomAttrs },
    // A tab inserted or deleted as a tracked change carries its side and author for the view (ED-005 section 6.3); a tab has no node view.
    toDOM: (node) => {
      const r = node.attrs['revision'] as Revision | null;
      if (!r) return ['span', { class: 'tab' }, '\t'];
      const who = r.deletion ?? r.insertion;
      return ['span', { class: `tab${r.insertion ? ' rev-ins' : ''}${r.deletion ? ' rev-del' : ''}`, 'data-author': who?.author ?? '' }, '\t'];
    },
  },

  /** `w:drawing` with `wp:inline`: shown, deletable, not editable. Anchored drawings are opaque. */
  image: {
    group: 'inline',
    inline: true,
    atom: true,
    draggable: false,
    attrs: {
      ...atomAttrs,
      sourceId: { default: null },
      /** A blob URL in the browser; null in Node. */
      src: { default: null },
      /** Points. */
      width: { default: null },
      height: { default: null },
    },
    toDOM: (node) => ['img', { src: node.attrs['src'] ?? '', width: node.attrs['width'], height: node.attrs['height'] }],
  },

  /** `w:footnoteReference` / `w:endnoteReference`; numbered by a decoration, the note's text is E1.e. */
  footnoteRef: {
    group: 'inline',
    inline: true,
    atom: true,
    attrs: {
      ...atomAttrs,
      sourceId: { default: null },
      footnoteId: { default: null },
      kind: { default: 'footnote' },
    },
    toDOM: (node) => ['sup', { class: 'footnote-ref', 'data-id': String(node.attrs['footnoteId']) }],
  },

  /** A run-level `w:sdt`: an inline node with inline content, so nesting is ordinary. */
  sdtInline: {
    group: 'inline',
    inline: true,
    content: 'inline*',
    attrs: {
      sourceId: { default: null },
      sdtPr: { default: null },
    },
    toDOM: () => ['span', { class: 'sdt-inline' }, 0],
  },

  /**
   * An inline item the editor does not edit: a field char, instruction text, a symbol, a picture,
   * an object, a comment reference, a bookmark or permission marker, a proofing mark, ... Emits
   * its original element (an item inside a run, or a paragraph-level item) on export.
   */
  opaqueInline: {
    group: 'inline',
    inline: true,
    atom: true,
    selectable: false,
    attrs: {
      ...atomAttrs,
      sourceId: { default: null },
      kind: { default: 'unknown' },
    },
    toDOM: (node) => ['span', { class: 'opaque-inline', 'data-kind': String(node.attrs['kind']) }],
  },
};

export const marks: Record<string, MarkSpec> = {
  /**
   * `w:rStyle`. First, so its element is the outermost of the formatting marks': a character style sits under the
   * run's direct formatting (ECMA-376 section 17.7.2), and in the view an inner element's rule beats what it inherits.
   * After the font mark, a style's size and font won over the run's own (ED-005 section 12.132).
   */
  runStyle: { attrs: { id: { default: null } }, toDOM: () => ['span', { class: 'run-style' }, 0] },
  bold: { attrs: {}, toDOM: () => ['strong', 0] },
  italic: { attrs: {}, toDOM: () => ['em', 0] },
  strike: { attrs: { double: { default: false } }, toDOM: () => ['s', 0] },
  /** `w:u`: the Office JS underline name (`Single`, `Double`, `Wave`, ...). */
  underline: { attrs: { val: { default: 'Single' } }, toDOM: () => ['u', 0] },
  /** `w:vertAlign`: 'subscript' or 'superscript'. */
  vertAlign: { attrs: { val: { default: 'superscript' } }, toDOM: (mark) => [mark.attrs['val'] === 'subscript' ? 'sub' : 'sup', 0] },
  /** Font name, size, colour and highlight as one mark; a field is null when not set directly. */
  font: {
    attrs: { name: { default: null }, size: { default: null }, color: { default: null }, highlight: { default: null } },
    toDOM: () => ['span', { class: 'font' }, 0],
  },
  /** `w:rtl`. */
  rtl: { attrs: {}, toDOM: () => ['bdo', { dir: 'rtl' }, 0] },
  /**
   * The run this inline content came from (ids.mts); absent on typed text. `alt` (ED-009) is the
   * kept `mc:AlternateContent` inside that run whose selected branch the text came from, or null:
   * text with a different `alt` never merges, so a branch's text keeps its edge, and the export
   * writes it back into the element's selected branch.
   */
  runSource: {
    attrs: { sourceId: { default: null }, alt: { default: null } },
    // Not shown; only carried.
    toDOM: () => ['span', { class: 'run' }, 0],
  },
  /** `w:hyperlink`: `sourceId` of the hyperlink, its relationship id and/or anchor. */
  link: {
    attrs: { sourceId: { default: null }, rId: { default: null }, anchor: { default: null } },
    inclusive: false,
    toDOM: () => ['a', { class: 'link' }, 0],
  },
  /**
   * Text inserted as a tracked change (`w:ins`, or `w:moveTo` when `move` is set; ED-005 section
   * 5.2). Inclusive, as the run it came from is: text typed at its end joins its run, and so its
   * `w:ins` (the tracker of step C3 decides what typing with tracking on makes).
   */
  insertion: {
    attrs: { author: { default: null }, date: { default: null }, move: { default: null } },
    toDOM: (mark) => ['ins', { class: mark.attrs['move'] !== null ? 'insertion move' : 'insertion', 'data-author': String(mark.attrs['author'] ?? '') }, 0],
  },
  /**
   * Text deleted as a tracked change, kept as text (`w:del` with `w:delText`, or `w:moveFrom`).
   * Not inclusive: text typed at its end is not deleted text, and export takes it out of the
   * deletion (ED-005 section 12.3). Two may overlap (`excludes: ''`, ED-005 section 12.7 item 5
   * (a)): y-prosemirror then keys each by a hash of its attributes, so two authors deleting
   * overlapping text at once both keep their deletion. The export writes the earlier
   * (`deletionOf`); the listing shows both, and rejecting one leaves the other.
   */
  deletion: {
    attrs: { author: { default: null }, date: { default: null }, move: { default: null } },
    inclusive: false,
    excludes: '',
    toDOM: (mark) => ['del', { class: mark.attrs['move'] !== null ? 'deletion move' : 'deletion', 'data-author': String(mark.attrs['author'] ?? '') }, 0],
  },
  /** A run's formatting changed as a tracked change (`w:rPrChange`): `before` is the old `w:rPr` as the JSON string of its plain form. */
  formatChange: {
    attrs: { author: { default: null }, date: { default: null }, before: { default: null } },
    toDOM: (mark) => ['span', { class: 'format-change', 'data-author': String(mark.attrs['author'] ?? '') }, 0],
  },
  /** A comment's range; several may overlap, hence `excludes: ''`. Read-only in E1. */
  comment: {
    attrs: { id: { default: null } },
    excludes: '',
    inclusive: false,
    toDOM: (mark) => ['span', { class: 'comment', 'data-comment': String(mark.attrs['id']) }, 0],
  },
};

/** The marks that describe formatting, in the order they are compared for run grouping on export. */
export const FORMATTING_MARKS = ['bold', 'italic', 'strike', 'underline', 'vertAlign', 'font', 'runStyle', 'rtl'] as const;

export const schema = new Schema({ nodes, marks });

export type EditorSchema = typeof schema;
