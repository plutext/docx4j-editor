# A left-aligned table's edge: ED-005 section 12.132 item 5's Word check

Built by `node packages/editor-model/test/fixtures/build/table-indent-for-word.mjs` from the repository root, written by hand
as XML with no engine, and rebuilt byte for byte. To be run in Word 365 through the portfolio's Word bridge (Jason,
2026-11-18: "build a small test file then ask portfolio agent to have Word run it via the bridge"), round `table-indent`. Each
job opens one file and saves it as docx and as PDF. Word's saves go beside the inputs as `ind15-mode15-word365.docx`, `.pdf`
and so on.

## The question

Where does Word put a left-aligned table, and so its cells' text, against the body text?

The editor outdents every left-aligned table by the table's default left cell margin (`tableVars`, `view/tables.ts`), so the
text of a cell with that margin lines up with the body. That is Word 2010's placement, and the editor uses it in every
compatibility mode. Word 365's PDF of a share application form in mode 15 put the edge of a table at the margin; its cells
set their own left margin to 0, and the editor drew the table 108 twips (5.4pt) left of Word's.

These candidate rules fit that one file differently:

- **A**: no outdent; the edge at the margin plus `w:tblInd`.
- **B**: outdented by the first row's first cell's own left margin (its `w:tcMar`, else the table's).
- **B′**: each row outdented by its own first cell's left margin, so every row's first text is at the margin.
- **C**: outdented by the table's default left margin, whatever a cell says (the editor's rule).

## The files

The same body in each: a reference paragraph `T0` at the margin, then seven bordered tables, each after a caption paragraph
that says what it holds. Every text starts with "T", so a glyph's side bearing is the same everywhere. A4, margins of 1440
twips, Arial 10pt, `TableNormal`'s cell margins 108 twips.

| File | Settings |
|---|---|
| `ind15-mode15.docx` | compatibility mode 15 (Word 2013 and later; a new document in Word 365) |
| `ind14-mode14.docx` | compatibility mode 14 (Word 2010) |
| `ind00-no-mode.docx` | no compatibility setting at all, as many generated files have |

| Table | `w:tblInd` | The table's left cell margin | The first cell's own (`w:tcMar`) |
|---|---|---|---|
| T1 | none | 108 (default) | none |
| T2 | none | 400 (`w:tblCellMar`) | none |
| T3 | none | 108 | 0, in both cells (the form's case) |
| T4 | none | 108 | 400 |
| T5 | 720 | 108 | none |
| T6 | 720 | 108 | 400 |
| T7 | none | 108 | row 1: 400; row 2: none |

## What to read from Word's saves

Each first cell's text against the body text, in points, from the PDF: `pdftotext -bbox`, the `xMin` of each cell's first
word (`T1`, `T2`, …, `T7` twice) less the `xMin` of `T0`. Each rule predicts this (the cell border adds about a point in the
editor):

| Table | A | B | B′ | C | The editor now (mode 15, Chromium) |
|---|---|---|---|---|---|
| T1 | 5.4 | 0 | 0 | 0 | 0.8 |
| T2 | 20 | 0 | 0 | 0 | 0.8 |
| T3 | 0 | 0 | 0 | −5.4 | −4.6 |
| T4 | 20 | 0 | 0 | 14.6 | 15.4 |
| T5 | 41.4 | 36 | 36 | 36 | 36.7 |
| T6 | 56 | 36 | 36 | 50.6 | 51.3 |
| T7 row 1 | 20 | 0 | 0 | 14.6 | 15.4 |
| T7 row 2 | 5.4 | −14.6 | 0 | 0 | 0.8 |

The three files answer whether the rule depends on the mode. Word's saved `document.xml` says whether Word wrote anything
of its own (a `w:tblInd`, a compatibility setting) when it saved.

## Run 1: Word 365 (2026-11-18, bridge 958f35f, Word 16.0.20326.20318)

Round `table-indent`, run by the portfolio session with the VM's go (323), between the controls `tind-mu-control-a` and `-b`.
All five jobs opened silently and saved completely, with no dialogs, repairs or Compatibility Checker, and each PDF has one
page. Word's saves are beside the inputs as `*-word365.docx` and `.pdf`. The readings below are this repository's, from
`pdftotext -bbox`; the portfolio's agreed to the tenth of a point.

| Table | mode 15 | mode 14 | no mode | A | B | B′ | C |
|---|---|---|---|---|---|---|---|
| T1 | 5.8 | 0.0 | 0.0 | 5.4 | 0 | 0 | 0 |
| T2 | 20.2 | 0.0 | 0.0 | 20 | 0 | 0 | 0 |
| T3 | 0.5 | 0.0 | 0.0 | 0 | 0 | 0 | −5.4 |
| T4 | 20.2 | 0.0 | 0.0 | 20 | 0 | 0 | 14.6 |
| T5 | 41.8 | 36.0 | 36.0 | 41.4 | 36 | 36 | 36 |
| T6 | 56.2 | 36.0 | 36.0 | 56 | 36 | 36 | 50.6 |
| T7 row 1 | 20.2 | 0.0 | 0.0 | 20 | 0 | 0 | 14.6 |
| T7 row 2 | 5.8 | −14.6 | −14.6 | 5.4 | −14.6 | 0 | 0 |

1. **Mode 15 is rule A:** no outdent; the table's edge at the margin plus `w:tblInd`, each cell's text its own margin in.
   Every reading is about 0.4pt right of A's, the 0.5pt left border drawn inside the table.
2. **Mode 14, and a file with no mode, are rule B:** the table outdented by the first row's first cell's left margin (its
   own `w:tcMar`, else the table's), so that cell's text is at the margin. Every row moves with the first: T7's second row,
   with the default margin, stands 14.6pt out in the page margin. B′ and C are out.
3. **No mode is read as mode 12,** and Word's save writes `compatibilityMode` 12 into it. Mode 14 and mode 12 place tables
   alike here.
4. **Word's saves keep every `w:tblInd`, `w:tblCellMar` and `w:tcMar` as written** and add none. Mode 15's save adds the
   four compatibility settings a new Word 365 document has, and every save adds `useWord2013TrackBottomHyphenation`.

The editor follows A in mode 15 and B below it since ED-005 section 12.132 item 5 (`tableVariables`, `view/tables.ts`, with the
session's `compatibilityMode`). `e2e/tables.spec.ts` holds the editor's placement in all three files to Word's readings, within
a point; `test/tables.test.mts` holds the mode read from each file and from Word's save of the one with none.
