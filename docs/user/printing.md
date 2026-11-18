# Printing

**Ctrl+P** (⌘P on a Mac) prints the document, not the editor around it. So do **Print** in the
application bar and **Print the document** in the palette (Ctrl+Shift+P). The browser's print dialog
opens, where you choose a printer or **Save as PDF**.

The browser lays out the pages, so they will not match Word's page for page: lines break by the
browser's text layout and fonts. Treat the print as a draft. When the dialog closes, the status line
says so.

## What prints

- **The document's text**, with its styles, formatting, lists, tables and pictures.
- **Each section on its own paper**: its size, its orientation and its margins, including a gutter
  and mirror margins. What stands in a margin prints too, as it does from Word: a table's edge, a
  negative indent, a change bar. Within a quarter of an inch of the paper's edge it is cut.
- **Word's page breaks and keeps:**
  - page breaks, and section breaks that start a new page;
  - "page break before";
  - "keep with next" and "keep lines together";
  - widow and orphan control;
  - table rows that cannot split.
- **A table's header rows** are repeated at the top of each page it runs onto.
- **Footnotes and endnotes** print together after the document's text, under the headings
  "Footnotes" and "Endnotes", numbered as in the text. They do not print at the foot of each page.
- **Tracked changes** print as the current view shows them:
  - **All markup**: insertions, deletions and change bars in their authors' colours;
  - **As accepted**: the text as if every change were accepted;
  - **Original**: the text before the changes.

  Choose the view in the Review perspective, or with **Show:** in the palette, before printing.
- **Content controls** print their content, or their placeholder text, without their boundaries
  or tags, as Word prints them.

## What does not print

- **Headers and footers**, and so **page numbers**.
- **Comments**, neither in the margin nor as highlights in the text.
- **What the editor draws to help you edit:** page boundaries, section break labels, search
  highlights, other people's carets and the selection.
- **Text boxes, equations and other content the editor cannot show** print as the placeholder you
  see on screen.
- **Columns** print as one column; line numbers, page borders and the document's background colour
  do not print.

An odd-page or even-page section break prints as an ordinary page break: browsers do not add the
blank page Word adds when the new section would otherwise start on the wrong side.

## The browser's own Print

The browser's menu (**File → Print**) also prints the document alone, with each section on its
paper, but without the footnotes and endnotes. Only the editor's Print prepares those.

## Firefox

Firefox prints the document too, with three differences from Chrome and Edge:

- it does not keep a heading with the paragraph after it;
- it does not apply widow and orphan control;
- it prints every page on the first section's paper size, turning only the orientation.
