# A table cell's own shading: ED-005 section 12.128's Word check

Built by `node packages/editor-model/test/fixtures/build/row-shading-for-word.mjs` from the repository root, after
`npm run build -w @docx4j/editor-model`. `shd0` to `shd2` are XML written by hand, with no engine, and rebuild byte for byte.
`shd3` is the model's own save of `shd0` after four `row.shading` commands. Its parts rebuild byte for byte, but the zip's entry
dates are the time of the save. To be run in Word 365 through the portfolio's Word bridge (Jason,
2026-11-15: "set up the Word bridge round"), round `row-shading` with the control pair `shd-mu-control-a` and `-b`. Each job is
opened, then saved as docx and as PDF. Word's saves go beside the inputs as `shd1-style-forms-word365.docx`, `.pdf` and so
on. Every input is in compatibility mode 15.

## The questions

The editor's Row shading writes a cell's own `w:shd` in Word's form, `w:val="clear" w:color="auto" w:fill="RRGGBB"`. Its
No colour takes the cell's own `w:shd` away, so a table style's shading shows again. A bridge job can only open and save
a file, so the round can't press Word's No Color. It measures what Word makes of each form a "no colour" could take
instead:

1. **Under a table style that shades the cell**, which forms show the style's colour and which hide it?
2. **Does Word's save keep each form**, or rewrite or drop it?
3. **A vertical merge** is one cell to the person but several `w:tc`s in the file. Which `w:tc`'s shading does Word draw the
   merged cell from? The editor writes the same shading on every `w:tc` of a merge, so that every reading agrees.
4. **Does Word read a file the editor's `row.shading` wrote** as the editor draws it?

The editor's own reading, for comparison: the forms with no colour of their own (`fill auto`, `val nil`) draw no background
of their own, so the style's colour shows (`shadingCss` in `view/tables.ts`).

## The files

Every table has two columns of 4513 twips, and single borders on every edge.

**`shd1-style-forms.docx`**: two styled tables. Each row is one form, named in its first cell, with the XML in its second.
In both styles the shading is orange (`FFC000`).

| Row | The cells' own `w:shd` | The style's shading under it |
|---|---|---|
| A1 | none | `ShadedTable`: the style's own `w:tcPr` shades every cell |
| A2 | `w:val="clear" w:color="auto" w:fill="auto"` | the same |
| A3 | `w:val="nil"` | the same |
| A4 | `w:fill="FFFFFF"` (with clear, auto) | the same |
| A5 | `w:fill="FFFF00"` (the editor's form) | the same |
| B1, B2 | none | `BandedTable`: `band1Horz` shades the odd rows, a band of one row, `w:tblLook` banding on |
| B3 | `fill auto` | a banded row |
| B4 | none | an unbanded row |
| B5 | `val nil` | a banded row |
| B6 | none | an unbanded row |

**`shd2-merged.docx`**: three unstyled tables, two by two, with the second column merged down.

| Table | The restart `w:tc` | The continuation `w:tc` |
|---|---|---|
| C | yellow | none |
| D | none | yellow |
| E | yellow | yellow (the editor's form) |

**`shd3-editor.docx`**: `shd0-base.docx` after the editor's commands, run through the model's headless runner and saved by the
model. Each cell's text says what the editor shows.
- Table 1: row 1 yellow, set at a caret. Rows 2 and 3 are grey (`D9D9D9`), set from a cell selection of their second column.
- Table 2: the second column is merged down. Row 1 is light blue (`00B0F0`), set at a caret, so the merged cell is light blue
  down to its foot. Row 2's first cell is white.
- Table 3: the orange style. Row 1 held its own yellow, which No colour took away, so it is orange like row 2.

## What to read from Word's saves

- **shd1's PDF.**
  - A1 and B1 should be orange: the style shows through a cell with no shading of its own. This is the check that the
    styles are read.
  - A4 should be white and A5 yellow.
  - **A2, A3, B3 and B5 answer question 1.** Orange means that form leaves the style's colour, as the editor's No colour
    does. White means it hides the style's colour, so it is what a No colour that hides a band would write.
- **shd1's saved `document.xml` answers question 2:** each row's `w:shd` as Word wrote it back.
- **shd2's PDF answers question 3.**
  - C yellow down to the foot means Word reads the restart.
  - D yellow means it reads each `w:tc` for its own row, or the last.
  - E should be yellow throughout.
- **shd3's PDF and save answer question 4.** The colours should be as the cells' text says, and the file should open with no
  repair.

## Run 1: Word 365 (2026-11-15, bridge 22479f7, Word 16.0.20326.20318)

Round `row-shading`, with the control pair `shd-mu-control-a` and `-b`. All five jobs opened silently and saved completely,
with no dialogs or repairs, and each PDF has one page. Word's saves are beside the inputs as `*-word365.docx` and `.pdf`. The
PDFs' colours were read by rendering the pages.

| Row | The cells' own `w:shd` | Word drew | Word's save |
|---|---|---|---|
| A1, B1 | none, over the style's orange | orange | none |
| A2, B3 | `clear`, `auto`, `fill auto` | **none**: the style's orange is hidden | as written |
| A3, B5 | `nil` | **orange**: the style shows | **dropped** |
| A4 | white | white | as written |
| A5 | yellow | yellow | as written |
| B2, B4, B6 | none, unbanded | none | none |

1. **Under a style, `w:val="clear"` with an automatic fill is an explicit no colour,** and hides the style's or the band's
   shading. `w:val="nil"` is no shading at all, and Word's save drops it.
   - The editor drew `fill auto` as no shading of its own, so the style's colour showed. That was wrong. It now draws it as
     `transparent` (`shadingCss`).
   - The editor's No colour removes the cell's own `w:shd`, as `nil` does, so the style's colour shows again. To hide the
     style's colour it would write `clear`, `auto`, `fill auto`. Whether Word's own No Color writes that is still not
     measured: a save job can't press it, and core-ts CR-002 section 67's Office JS round measures Office JS's setter.
2. **Word's save keeps every form but `nil`.** It reorders the attributes (`w:val` first) and adds docProps, fontTable,
   theme and webSettings.
3. **A merged cell is drawn from its restart `w:tc`.**
   - C (the restart yellow) is yellow down to the foot.
   - D (only the continuation yellow) shows no yellow.
   - E (both yellow, the editor's form) is yellow.

   The editor draws a merge from the restart node too, so the two agree. Writing the continuations as well (section 12.128
   item 2) changes nothing Word draws. It keeps the file saying one thing.
4. **Word reads the editor's file as the editor draws it** (shd3): yellow, grey, light blue to the merge's foot, white, and the
   orange style's two rows. Word's save keeps every `w:shd` value.

`test/tables.test.mts` pins Word's saves. `e2e/tables.spec.ts` pins the editor's drawing of shd1 and shd2 to Word's PDFs.
