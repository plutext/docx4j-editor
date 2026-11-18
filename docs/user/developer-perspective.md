# The Developer perspective: what E4 added

The Developer perspective shows the document as its XML and lets a developer read and change it
below the editor's own commands. This page covers what phase E4 added: reveal codes in Monaco,
Select by XPath, and the parts panel made editable, and the normalisation report that came after it.
The console and macros are left to the whole guide, when it is written; the ask box is in `Help_AskBox.md`.

## Reveal codes in Monaco

The **Reveal codes** tab shows the XML of the block the caret is in (or of the blocks selected, or
of the run under the caret), as the editor would write it, in the Monaco code editor: coloured,
folded by element, with line numbers, read-only until Edit. Inside it Ctrl+F and Ctrl+H are
Monaco's own find and replace, apart from the document's find bar.

- **In a table**, the cell the caret is in is marked, and the other rows and, in its row, the other
  cells are folded, so a large table reads as the one cell with its row around it; the folds open
  with a click on their gutter arrows. The breadcrumbs name the way down (`tbl > tr > tc > p`), and
  moving the caret moves the mark.
- **Edit** makes the text editable, keeping the folds. The **Styles** tab edits a style's XML the
  same way.
- **Apply** writes the XML back as the block or the run; **Cancel** leaves it. A text that does not
  parse is refused, its line marked in the editor.
- Ticking **API script** shows the content-API script that makes the block, which **Open in console**
  takes to the console.

## Select by XPath

The **XPath** tab selects nodes of a part by an XPath expression and lists the hits: each with its
address, its node and its text. Clicking a hit selects it on the page, in the body or in a note or
header that is shown. **Part** chooses the part; **Main document** is the default. **Open in
console** turns the expression into a console script that selects the same.

- **XPath 1.0** is the browser's own, and the default: `//w:p[w:pPr/w:pStyle]` over a body of 255
  pages takes about 0.6 seconds in Chromium.
- **XPath 3.1** is FontoXPath's, an option for its functions (`matches()`, `string-join()`, …). It is
  much slower on a long document: the same expression takes about 25 seconds in Chromium. The count
  line gives each run's time.
- The expression runs over a copy of the part as the editor would write it, never over the document
  the editor holds, so a query changes nothing. The prefixes are the object model's (`w:`, `r:`,
  `w14:`, …), whatever the file itself declares.

## The parts panel

The **Parts** tab lists the package's parts by their relationships, from `/_rels/.rels` down, with
the parts no relationship reaches at the end. Choosing a part shows its XML: in Monaco, with Apply
and Revert, where it can be edited here.

**What is editable here:**

- **Checked that nothing would be lost** ("typed"): the styles, numbering, settings, theme and font
  table parts, the document properties, and other parts the object model reads. An Apply reads the
  text through the object model and refuses it if an element or attribute would be dropped.
- **Kept as XML** ("bytes"): custom XML data parts and other parts the engine keeps as they are. An
  Apply checks only that the text parses.

**What is read-only here, and where it is edited instead:**

- the main document, the notes, the headers and the footers: on the page, or a block's XML in Reveal
  codes;
- the comments and their side parts: in the comments margin;
- the relationships: the engine writes them at save;
- binary parts (images, fonts).

**Apply and Revert:**

- **Apply** checks the text first and writes it only if nothing is wrong. A refusal lists each
  problem with its path in the part and its line: a text that does not parse, a value the part cannot
  hold, an element or attribute that would not be saved. The problems are announced to a screen
  reader as they appear.
- After an Apply the page redraws: a style's new colour shows at once, and list labels follow the
  numbering part. If the document still names a style or list the part no longer defines, the status
  says which.
- **A part edit has no undo.** **Revert** puts back the part as it was when the pane opened it, which
  is another Apply.
- **In a shared document** a part edit reaches everyone, entry by entry (a style, a list), so two
  people editing different styles do not overwrite each other. If someone changes the part you are
  editing, the pane says so and offers to reload it or to apply yours over it.

## What first marked a document changed

The dot after the document's name says it has unsaved changes. What first put it there since the open or
the last save is kept (ED-011 section 9.9): in the session's record in IndexedDB (`docx4j-editor`, store
`sessions`, field `firstEdit`), and in a log of the last twenty in `localStorage`, which outlives the session.
To read the log in the browser's console:

```js
JSON.parse(localStorage.getItem('docx4j-editor.editCauses'))
```

Each entry names the session and document, the time, the transaction's origin (`y-sync$` is an edit in a
view, `undo` the undo manager, `local` a command run without a view), the store entries and body blocks
it changed, the last key or pointer input before it (a character is only ever `character`), where the focus
was, and the stack that made it. Nothing of the document's text is kept.

## The normalisation report

A save writes each value as the file has it, so a lower-case colour stays lower case. The report, at the
end of the **Diagnostics** tab, shows what writing the values in their normal form would change. It is
made when you press **Run the report**, and it changes nothing.

- **Values not in normal form** are listed by part, then by element and attribute, then by value, with
  how often each occurs: `w:color/@w:val`, `ffffff` would be written as `FFFFFF`, 77 times. A value with
  no normal form is left out: a highlight of `purple`, which is not on Word's list.
- **Notes from reading the file** are what the engine recorded while it read the parts, by code, in its
  own words. `mce.branches_differ` says a kept `mc:AlternateContent`'s branches hold different text.
- **It reads the document as it is when you run it**, your edits included, in every part. It does not
  follow later edits: press the button again.
- **The 255-page test document takes two or three seconds.**
