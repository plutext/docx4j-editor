# Where the editor differs from Word

What Word does, or what its API (Office JS) reports, that the editor does differently. Each entry
has been measured in Word or decided knowingly; each says what Word does, what the editor does,
why, and whether that is settled. Add an entry when a check in Word finds a difference. When one is
fixed, move it to the last section with the commit that fixed it, rather than deleting it.

Differences in third-party libraries, and how the editor works round them, are in
`upstream-limitations.md`. Differences in the engine are in core-ts's change requests.

## Open: a decision or a fix waiting

### A control bound with tracking on is not itself marked inserted

- **Word:** its XML Mapping Pane's Insert Content Control over a selection, with tracking on, marks
  the new control inserted (`w:customXmlInsRangeStart` and `End` around it) and, inside it, the
  selected text deleted and the node's value inserted (core-ts check 34, 1j, `34b-tracked.docx`).
  Reject All removes the control with its binding, puts the text back, and leaves the part as it
  was, still "Ann" (`34b-tracked-reject-all.docx`, Jason, 2026-10-03).
- **The editor:** writes the deletion and the insertion inside the control, as Word does
  (`check34.test.mts`), and the control unmarked: the model holds no revision on a control. A
  reject puts the old text back in the control, which stays bound, and the text is then written to
  the node, as any edit in a bound control is, so every control bound to that node shows it.
- **Why:** a control's own revision is new to the model (its import, export, listing and
  resolution); not yet decided.
- **Settled:** no (ED-005 section 12.57).

### Typing in a rich text control mapped with `w15:dataBinding` is not written to its node

- **Word:** writes the control's content into its node as escaped Flat OPC, a whole `pkg:package`
  (core-ts check 35, 1e: about 51 KB for a few words).
- **The editor:** writes nothing back from such a control, so its node keeps the value it had. It never writes over
  a node's Flat OPC from another control. The other direction works since core-ts 0.8.0: a node's Flat OPC fills the
  control through the engine, offered at open by the stale banner and done on a bind (ED-005 proposal 46, sections
  12.119 and 12.124).
- **What it costs:** Word 365 fills such a control from its node when it opens the file, and keeps the mapping on save
  (proposal 46's check 46c, in compatibility mode 15). So an edit made in the editor inside the control is replaced by
  the node's content when the file is opened in Word. Word 2010 drops the `w15:dataBinding` when it saves (check 35
  part 2), so through Word 2010 the control is unbound.
- **Why:** writing the control's content as Flat OPC is proposal 46's reverse direction, which Jason's decision (c)
  left out (`TODOs_OpenDoPE.md`).
- **Settled:** no (ED-005 section 12.58; proposal 46 decision (c)).

### A stale `w:cnfStyle` cache after a table is edited

- **Word:** writes `w:cnfStyle` on rows, cells and paragraphs as a cache of the conditional formats
  they are under (Word 2007 and 2010 did; `headers-footers.docx` and the 255-page document have them).
  Whether Word trusts a cache that no longer matches the position (a header row's, after a row is
  added above it) or recomputes from the position is not measured.
- **The editor:** leaves the caches as they were when it edits a table, and draws as the engine
  resolves (core-ts CR-007), which takes a cache over the position, as docx4j does. A row added above
  a header row with a cache leaves both drawn as header rows.
- **Why:** the engine's rule, not measured in Word.
- **Settled:** no. A check in Word would decide whether the engine should prefer the position, and
  whether the editor should drop or rewrite a table's caches when it edits its rows.

## Deliberate

### A bound control is not refreshed from its node at open

- **Word:** fills a mapped control from its node when it opens the file: a text control's text
  (core-ts check 34), and a picture control's picture, in Word 2010 and Word 15 (check 38,
  2026-10-08: the template showed the node's picture, its placeholder part gone from the save).
  Word 365 fills a `w15`-mapped rich text control from a node holding Flat OPC the same way (ED-005 proposal 46's
  check 46c).
- **The editor:** shows the file as it is at open, and refreshes no bound control unasked. Instead it counts the bound
  controls a refresh would change and, when there are any, offers a banner, once per document: Refresh from the data
  or Keep as it is (ED-005 section 12.89).
  - It counts text, XHTML and picture controls, and Word's `w15` fragment controls (section 12.119).
  - It doesn't count a date picker or a check box, since the editor doesn't refresh those yet (`TODOs_OpenDoPE.md`).
  - Otherwise a bound control is refreshed only on an occasion of its own: a bind, a Data tab edit or Apply, a bound
    control's write-back, the Control tab's XHTML switch (sections 12.54, 12.75, 12.76, 12.124).
- **Why:** an unedited document round-trips byte for byte (ED-001 section 3), and a refresh is an
  edit; a template's placeholder content is what its author sees. The processor's bind, or Word,
  fills it for an instance. The banner leaves the choice to the person (Jason, 2026-10-09).
- **Settled:** yes (the round-trip invariant).

### A replacement is one entry in the review panel

- **Word:** its Reviewing Pane lists a deletion and the insertion typed over it as two revisions,
  Deleted and Added, and Next steps to each. Seen by Jason in current Word (2026-09-30).
- **The editor:** the review panel lists them as one entry, "Replaced “7 days” with “14 days”",
  whose Accept and Reject resolve both halves in one step. Split shows them apart.
- **Why:** one edit, one decision (ED-005 proposal 41, section 9.7).
- **Unchanged:** the file keeps a `w:del` and a `w:ins`, as Word writes them; the agent's
  `get_tracked_changes` lists the two halves apart, as Office JS does.

### The Data tab's bind across paragraphs or over a picture is refused

- **Word:** its XML Mapping Pane's Insert Content Control over a selection from one paragraph into
  the next replaces the selected part of the first paragraph alone, leaving the second as it was
  (core-ts check 34, 1f); over a selection holding a picture it replaces the text and moves the
  picture to the end of the paragraph (1g).
- **The editor:** refuses both, saying why in the status line (`plainTextControlRefusal`).
- **Why:** Jason's decision (2026-10-03): neither result is what the selection asked for.
  `check34.test.mts` pins Word's result beside the refusal.

### A plain text control on a node mapped to rich text is refused

- **Word:** its XML Mapping Pane's Insert Content Control > Plain Text, on a node a rich text control
  is mapped to, nests the control inside that rich text control with the alert "The custom XML node
  is already mapped to a rich text content control, so it can't be mapped to a plain text content
  control", and saves it with a `w:dataBinding` it then ignores (core-ts check 35, 1e).
- **The editor:** refuses, with Word's reason in the status line, and binds no plain text control to
  a node a rich text control is mapped to (`dataBindRefusal`).
- **Why:** a binding Word 15 ignores would be written back by the editor, over the rich text's
  content. Word 2010 does not ignore it: it drops the rich text control's `w15:dataBinding` and shows
  the node's Flat OPC in the nested control as raw XML (check 35 part 2, `35c-word2010.docx`).

### A plain text bind at a caret in an empty paragraph makes a run-level control

- **Word:** its XML Mapping Pane's Insert Content Control with the caret in an empty paragraph makes
  a block-level plain text control around the paragraph (core-ts check 34, 1d, `34a-mapped.docx`).
- **The editor:** the Data tab's bind (`contentControl.insertBound`) makes a run-level control inside
  the paragraph there, as it does at any other caret; the paragraph stays. `check34.test.mts` pins
  both results.
- **Why:** a block control around a document's last paragraph holds the document's final paragraph
  mark, which no key can leave, so everything typed after the bind went into the control (Jason,
  2026-10-08). Word avoids that by keeping a paragraph after the control, which the editor now does
  for a block control it makes at the end of a document, a cell or a block control
  (`keepParagraphAfter` in `commands/controls.mts`); for a bind, Jason chose the run-level control.
- **Settled:** yes (Jason's decision, 2026-10-08; ED-005 section 12.71).

### A block-level rich text control mapped to an empty node stays a block

- **Word:** mapping a block-level rich text control holding two paragraphs to an empty node left a
  run-level control, showing the placeholder, at the start of the next paragraph (core-ts check 35,
  1b).
- **The editor:** shows the placeholder in one paragraph inside the block-level control.
- **Why:** Word's result joins the control to a paragraph it was not in.

### Text deleted by a tracked bind keeps its formatting

- **Word:** a plain text control gives its content one formatting, so the selected text it keeps
  deleted inside the new control loses its own ("brown" its bold; check 34, 1j), and Reject All puts
  it back without it (`34b-tracked-reject-all.docx`).
- **The editor:** keeps the deleted text's formatting, which a reject then puts back as it was.
- **Why:** a reject should restore what was there.

### Style names as the styles part names them

- **Word:** its interface shows its built-in styles capitalised ("Heading 2" for the style named
  `heading 2`).
- **The editor:** the style box and the review panel show the name the styles part holds
  ("heading 2").
- **Why:** the two agree with each other and with the file (ED-005 section 12.40 item 2).

### The review panel's snippet is plain text

- **Word:** its Reviewing Pane shows inserted and deleted text with its formatting.
- **The editor:** shows it plain. The formatting a change made is worded instead ("Font: Bold").
- **Why:** Jason's decision when the panel gained values (ED-005 section 12.32 item 2).

### A paragraph after a nested table shows its own style

- **Word:** Office JS reads the style of a paragraph that follows a nested table in a table cell as `''`
  (`styleBuiltIn` `Other`) unless the paragraph is Normal, for `Paragraph.style` and its range's alike
  (core-ts checks 52 and 53, Word 365, 2026-10-22). Word's own saves keep the `w:pStyle`, and its
  Styles pane shows it.
- **The editor:** the style box, the outline, the review panel and the export take the paragraph's
  `w:pStyle`, so a numbered item or a heading after a nested table is one; the paragraph the cell
  rule appends after a nested table carries the copied style the same way (ED-005 section 12.109).
- **Why:** following the reading would hide a style the paragraph has. core-ts reads the own style too,
  its CR-002 section 57 (Jason's decision, 2026-10-22); the two agree with each other and the file.
- **Settled:** yes. An Office JS reading, not a behaviour of Word's editing or its files.

### A whole paragraph pasted at a paragraph's end leaves the target as it was

- **Word:** pasting a paragraph with its mark at the end of a paragraph gives the target paragraph the
  pasted paragraph's properties (a plain target became a numbered List Paragraph) and leaves the target's
  own mark as an empty paragraph after it (Word 365, ED-010 section 9.1, cases 3, 5, 7 and 8).
- **The editor:** the target keeps its text and properties, the pasted paragraph follows it with its own,
  and no empty paragraph is made (ProseMirror's fitting of a closed slice).
- **Why:** a person's own paragraph does not change its style because of what was pasted after its text,
  and the empty paragraph is a leftover (ED-010 proposal P8, accepted 2026-10-22).

### A pasted list item continues its list

- **Word:** a numbered item copied and pasted elsewhere starts a new list, the definition copied
  (`abstractNum` with a new `nsid`, a new `num`), so the pasted item reads 1 (ED-010 section 9.1, case 8).
- **The editor:** the pasted item keeps its `numId`, so it continues the list it came from.
- **Why:** Word's new list needs a numbering-part write in the paste transaction, as `list.restart`'s
  `prepare` half makes one; a step of its own if wanted (ED-010 proposal P9, accepted 2026-10-22).

### A cut and paste with tracking on is a deletion and an insertion, not a move

- **Word:** cutting a paragraph and pasting it elsewhere with tracking on records a move: `moveFrom` at the
  original place and `moveTo` at the paste, paired by name, a bookmark inside the cut text left in the
  `moveFrom` and gone once the move is accepted (Word 365, ED-010 section 9.3, 2026-10-29).
- **The editor:** the cut is a tracked deletion and the paste a tracked insertion by the same author; a
  bookmark in the cut text stays with the deleted original and goes when the deletion is accepted, as in
  Word's accepted file. The review panel lists two entries where Word's pane lists one move.
- **Why:** the tracker records one transaction at a time and the paste does not know it follows a cut;
  pairing the two into a move would need the cut remembered across the clipboard, a step of its own if
  wanted.

### Revisions in notes, headers and footers are not listed

- **Word:** its Reviewing Pane lists them with the body's.
- **The editor:** the review panel lists the body's revisions only, and says so at its foot. Notes
  and headers keep their revisions and show them.
- **Why:** E4's scope (ED-005 section 6.9).

### A formatting change is drawn in the text

- **Word:** shows a formatting change in a balloon in the margin ("Formatted: Font: Bold").
- **The editor:** underlines the changed text with a dotted line in the author's colour. The review
  panel words the change.
- **Why:** E4 leaves out balloons (ED-005 section 6.3, proposal 26).

### Printing is the browser's (ED-006 level 1)

- **Word:** prints its own pages: its line breaking and fonts, headers and footers, page numbers,
  footnotes at the foot of each page, comments in balloons when markup is printed. An odd-page or
  even-page section break adds a blank page when the parity requires it.
- **The editor:** prints through the browser (ED-001 decision 25). Pages break where the browser's
  layout puts them; headers, footers, page numbers and comments are not printed; footnotes and
  endnotes follow the body. An odd-page or even-page section break prints as a plain page break: neither
  Chromium nor Firefox adds the blank page for `break-before: right` or `left` (ED-006 section 9.1,
  case F). Firefox also ignores keep with next, orphans and widows, and prints every page on the first
  page's paper size, turning only the orientation.
- **Why:** the editor lays out no pages (ED-001 Appendix D). Headers, footers and page numbers are
  level 2's (ED-006 section 2.2); Word's own pages would be level 3's.
- **Settled:** yes, for level 1 (ED-006, accepted 2026-10-06).

### The As accepted and Original views are approximations

- **Word:** No Markup lays out two paragraphs that a deleted mark joins as one paragraph, in the
  first one's properties (a heading's, after a heading). Original shows the original formatting,
  with the change bar and the balloon still drawn. Measured in Word 15 (ED-005 section 12.33 items
  3 and 4).
- **The editor:**
  - As accepted draws the joined pair inline, in the formatting of what contains them;
  - Original shows today's formatting, still marked.
- **Why:** the three views are CSS over one editable document, not a second rendering (ED-005
  proposal 22). Since the heading's deleted mark carries its properties to the next paragraph (the
  Fixed section), the joined line after a heading is drawn as Word's No Markup draws it.

### Office JS's hyphen characters in text from elsewhere

**Word:** its writers turn U+001E into `w:noBreakHyphen` and U+001F into `w:softHyphen` (core-ts check 55).
**The editor:** text pasted, dropped or scripted into a document keeps them as the characters Word shows them as, U+2011 and
U+00AD. The edit layer holds those two elements only as opaque atoms of a source element, and a new one has none. The agent's
reading and a file's own hyphens are unaffected. **Settled** until the edit layer can make such an atom (ED-005 section 12.114).

### A picture control whose node holds a data URL

- **Word:** reads a mapped picture control's node as bare base64 only (Word 365, 16.0.20326.20318, through the portfolio's
  Word bridge, 2026-11-12; `editor-model/test/fixtures/pictures-dataurl/`, ED-005 section 12.125):
  - over a data URL (`data:image/png;base64,...`) it keeps what the control holds, its own picture or its placeholder, with
    no message;
  - over bare base64 it swaps the control's picture for the node's on open;
  - it writes nothing back to the node.
- **The editor:** reads a data URL as the base64 after its comma, as the engine's bind step does, so it fills the control
  from the node, and its stale banner offers a control whose picture differs from it.
- **The same file shows the same picture in both** when the editor filled the control. Edit the node afterwards and Word goes
  on showing the old picture. A file meant for Word should hold bare base64. **Open:** the editor follows the engine until
  Jason decides whether it should read a data URL as Word does.

### Row shading on a row a vertical merge reaches colours the merged cell

- **Word:** Office JS's `TableRow.shadingColor`, on a row a vertical merge reaches from above, writes that row's own `w:tc`s,
  the merge's continuation among them, and never the restart in the row above. Word draws a merged cell from its restart,
  so the merged cell keeps its colour. On the row a merge starts in, it writes the restart and no continuation below
  (core-ts's Office JS round, Word 365 16.0.20326.20318, 2026-11-17; core-ts CR-002 section 67.3). Word's own Shading
  button was not measured.
- **The editor:** Row shading colours a merged cell whole from any row it reaches, and writes the colour on every `w:tc` of
  the merge, so the file says one thing whichever `w:tc` a reader draws from (ED-005 section 12.128 item 2).
- **Why:** a person who shades a row expects the merged cell in it to change colour, which Word's setter doesn't do from a
  lower row.
- **Settled:** yes (Jason, 2026-11-18: "yes keep").

### A picture control holding no picture, mapped to a picture

- **Word:** fills a mapped picture control only by swapping the picture it holds (Word 365, 16.0.20326.20318, through the
  portfolio's Word bridge, 2026-11-12; `editor-model/test/fixtures/pictures-dataurl/` pdu1, pdu5 and pdu7, ED-005 section
  12.125). Placeholder text with the flag is kept, an empty control stays empty, and text without the flag is content Word
  repairs.
- **The editor:** fills it with the node's picture at its natural size, as the engine's bind does, and its stale banner
  counts it. A picture control the editor makes now holds Word's form, a picture Word fills (ED-005 section 12.126), so such a
  control comes from another producer, or from the editor before that.
- **Settled** for the engine by Jason (core-ts CR-005 section 10.2, E2, 2026-11-13: "E2 recorded as a difference from Word");
  the editor follows the engine.

## Fixed

### A new picture control is never filled by Word

- **Word:** fills a mapped picture control only by swapping the picture it holds, on open and with the placeholder flag or
  without it (Word 365, 16.0.20326.20318, through the portfolio's Word bridge, round `picture-placeholder`, 2026-11-12;
  ED-005 section 12.125). With no picture there is nothing to swap: placeholder text with the flag is kept, an empty
  control stays empty, and text without the flag is content Word repairs. Given Word's own form (the flag and a
  placeholder picture, as docx4j writes it) it shows the node's picture.
- **The editor, before:** made a new picture control holding placeholder text (`contentControl.insert`), so a picture control
  it made and saved before a refresh put a picture in showed only that text in Word.
- **Fixed** in ED-005 section 12.126, merged at the core-ts 0.9.0 install (1326b7a): a new picture control holds Word's form,
  the flag and docx4j's placeholder picture from core-ts's exported `showPicturePlaceholder`.

### A left-aligned table's edge

- **Word:** in compatibility mode 15 puts a left-aligned table's edge at the margin plus `w:tblInd`, so a cell's text is its
  margin in from the body text. Below mode 15 (14, and 12 for a file that states none) it outdents the table by the first
  row's first cell's left margin, so that cell's text lines up with the body text (Word 365, round `table-indent`,
  2026-11-18; `test/fixtures/table-indent/`).
- **The editor, before:** outdented every left-aligned table by the table's default left cell margin, whatever the mode or
  the first cell's own margin. A share application form in mode 15, its cells' own margin 0, was drawn 108 twips left of
  Word's.
- **Fixed** in ED-005 section 12.132 item 5: `tableVariables` follows the mode and the first cell's own margin.

### A run's own font, size and bold under a character style

- **Word:** draws a run's own formatting over its character style, and a run's own `<w:b w:val="0"/>` plain under a
  bold style (a share application form and Word 365's PDF of it, 2026-11-18).
- **The editor, before:** drew the character style's font, size, colour and bold over the run's own, since the style's
  element was inside the run's formatting's, and never drew a run's own "off".
- **Fixed** in ED-005 section 12.132: the character style's mark is the schema's first, the outermost; a run's own bold,
  italic, caps or small caps switched off is a decoration.
- **And the line's height:** the style's element, outermost, still had the style's size, so a line under a 14pt style
  with the runs' own 11 points was 16.1pt tall where Word's is 12.6pt. **Fixed** in ED-005 section 12.133: the style's
  size and family are drawn on its text, where the run states none of its own.

### A table style's conditional formatting is not drawn

- **Word:** draws a table style's conditional formatting (`w:tblStylePr`: the header row, the first
  column, banded rows) as the table's `w:tblLook` turns it on, the cells' shading and borders and
  their text's formatting alike.
- **The editor, before:** drew the table style's own properties and each cell's `w:tcPr`, with no
  conditional formatting and none of the table style's text formatting (ED-003 sections 3.8 and 3.9).
- **Fixed** in 0ebc853 (ED-005 section 12.65), on core-ts 0.3.2's CR-007:
  - each cell is drawn by the conditional formats that apply to it, its shading, borders and margins
    merged in Word's order;
  - its paragraphs' text is resolved with the table level, a paragraph's own `w:cnfStyle` honoured.
- **Left:** a conditional format's inside borders (`insideH`, `insideV`) and its table-level and
  row-level properties (`w:tblPr`, `w:trPr`) are not drawn.

### A comment on part of a pending insertion, made by the agent

- **Word:** keeps the anchor narrow: the start marker inside the `w:ins`, the `w:ins` closed at the
  range's end, the end marker and a plain reference run after it, and a second `w:ins` for the rest
  of the insertion. Word 2010 and Word 15 alike (ED-005 section 12.33 item 2).
- **Fixed** 2026-10-02 with core-ts 0.3.1 (ED-005 section 12.50): the engine's `Range.insertComment`
  writes Word's form, so the agent's comment stays on its words, and the tool's warning is gone. The
  editor's own comments had the narrow anchor already, their end marker just inside the first
  `w:ins`: a valid form, which Word reads.

### A comment whose text a rejected insertion takes

- **Word:** removes the comment, its reference run (plain, outside the revision) and its entry, when
  its start marker goes with the revision: an insertion rejected, alone or by Reject All, or a
  deletion accepted (current Word, core-ts check 32, 2026-10-02). Where the markers lie outside the
  revision, Word 2010 and Word 15 keep the comment as its reference alone (their Reject All saves of
  the editor's session file, `fixtures/review/`). **The outside form in current Word, measured
  2026-10-21** (Jason's Word 365 Reject All, `review-for-word-word365-rejected.docx`): Word 365
  removes it too, and core-ts 277d118 follows (CR-002 section 54). A comment keeping some text
  beyond the revision is not measured.
- **Fixed** 2026-10-02 (ED-005 section 12.49): the four resolving commands remove such a comment's
  reference, and the save drops the comment; Ctrl+Z brings it back. A comment made in the editor has
  its start marker inside the revision its first word is in, as Word's form does. Held to check 32
  and to both Words' saves by `check32.test.mts` and `word-review.test.mts`.
- **Not followed, in form only:** for a comment the file marked outside the revision, the editor
  writes its markers, collapsed, before the reference, where Word writes the reference alone. Word
  writes collapsed markers for a comment that never had text (a comment at a caret), and the editor
  does not tell the two apart. Both forms are valid, and Word reads either.
- **Open** (2026-10-21, ED-005 section 12.49 item 9): the comment the file marked outside the
  revision, which the editor keeps as Word 2010 and 15 did and Word 365 removes. Whether the editor
  follows Word 365 is Jason's decision; `word-review.test.mts` pins the one difference exactly.

### A comment's reference between two insertions

- **Word:** Office JS lists one author's insertions either side of a plain comment reference run as
  one change (current Word, core-ts check 32).
- **Fixed** 2026-10-02 (ED-005 section 12.49): the editor's listing lets a reference run part
  nothing, as a bookmark or a range marker parts nothing. A comment on part of an insertion is one
  entry in the review panel again.

### An inserted table in the list of tracked changes

- **Word:** Office JS's `body.getTrackedChanges()` groups inserted or deleted rows with whatever
  inserted or deleted content touches them: neither a table's edge nor a row's is a break. Measured
  in Word 16.0.20326.20158 by core-ts checks 30 and 31 (2026-09-30 and 2026-10-01): a table between
  inserted paragraphs is one change with them, rows added at a table's end are one change with the
  text typed after it, and rows either side of a plain row are two.
- **Fixed** 2026-10-01 at Jason's decision (ED-005 section 12.45): the editor's listing does the
  same, with a row's text as Office JS's (its cells with a tab between, then `\r\n`). Held to both
  checks by `check30.test.mts` and `check31.test.mts`, Accept all and Reject all included.
- **Not followed** (core-ts's `docs/word-known-issues.md` entry 7): Office JS lists nothing for a
  middle row by another author, which its own Accept all resolves; the editor lists it.

### Alignment set to what the paragraph would have anyway

- **Word:** Ctrl+L with tracking off, on a paragraph centred directly over a left-aligned style,
  removes the paragraph's `w:jc` (ED-005 section 12.33 item 2).
- **Fixed** in E4.g step G5 (ED-005 section 12.44): `paragraph.alignment` removes the value the
  paragraph resolves to without it. Held to both Words' saves by `test/word-measure.test.mts`.
  Until core-ts 0.3.1, the engine still writes it (a `known` case in `api-cases.mts`).

### A comment on an empty paragraph

- **Word:** writes the comment's reference run alone, with no range markers (ED-005 section 12.33
  item 2).
- **Fixed** in E4.g step G5 (ED-005 section 12.44): the export places no markers in a paragraph
  holding nothing but comment references and markers. Held to both Words' XML.

### A heading's paragraph mark deleted with tracking on

- **Word:** gives the next paragraph the heading's properties, with a `w:pPrChange` recording its
  own (ED-005 section 12.33 item 3).
- **Fixed** in E4.g step G5 (ED-005 section 12.44): the tracker does the same. Held to Word 15's
  save, and accepted and rejected both ways.
- **Only for a mark deleted alone,** its text kept to join the next paragraph. A whole paragraph
  deleted, its text and its mark, hands the next paragraph nothing (current Word, core-ts check 33).
  The tracker had carried then too, until 2026-10-02 (ED-005 section 12.51); `check33.test.mts` holds
  it to Word's saves.

### A kept `mc:AlternateContent` edited: the other branches

- **Word:** draws the first Choice it understands; when it has edited the shape it regenerates the
  Fallback from the Choice, and when it has not it keeps an authored Fallback verbatim, stale or not; a
  `w:altChunk` anywhere makes it regenerate every one; a run-level element holding bare text it resolves
  to a plain run on save (docx4j CR-021 section 8.5, measured 2026-09-19 and 2026-10-17; ED-009 section
  1.2, 2026-10-20).
- **The editor:** draws the first Choice the engine understands (`UNDERSTOOD_PREFIXES`), else the
  Fallback, else nothing (an opaque atom kept whole, since core-ts 0.7.0's CR-029; the first Choice
  before it); writes an edit into that branch and keeps every other branch as it
  found it, in every case, so an edited Choice leaves a Fallback Word would have rewritten (ED-009
  section 3.8; the editor cannot make a text box's VML counterpart). A formatting split writes a later
  piece whose other branches are present and empty. **The branch shown agrees with Word's** since
  core-ts CR-010 section 15 (ED-009 section 9.2 items 6 and 7): `w16sdtdh`, `w16sdtfl`, `w16du` and
  Excel's `x14ac`, which Word 365 understands (measured 2026-10-20) and the engine's set had lacked, are
  understood by `Requires` alone, so a Choice requiring one of them first shows in both ("I prefer
  w16du" there and here, where the editor had shown "I prefer w14").
- **Word's save of the editor's shapes** (ED-009 section 9.2, Word 365, 2026-10-20,
  `test/fixtures/mc/mc-for-word-word365.docx`): every element, edited or not, run-level or
  paragraph-level, becomes the selected branch's items as plain runs and every Fallback is dropped;
  the editor's split is three runs; the tracked deletion in the Choice accepted removes its text alone.
  The editor keeps the element with every branch; a Word 365 save after it removes the question.

### A row of exact height

- **Word:** a row with `w:trHeight w:hRule="exact"` is drawn at that height whatever its content,
  which is clipped past it.
- **The editor:** draws a stated row height as the row's CSS height (since ED-007, 2026-10-14; it had
  drawn none), which a browser's table row treats as a minimum: an exact row whose content needs more
  grows, where Word clips. A row of minimum height (`atLeast`) is drawn as Word draws it.
- **Fit columns removes both** (ED-007 section 3.5 as revised), so a fit table has no such row.

### A cell ending in a nested table

- **Word:** requires a cell's last block to be a paragraph. A file whose `w:tc` ends with a `w:tbl`
  is reported as unreadable content and offered a repair; the repair appends an empty `w:p` after the
  nested table and changes nothing else in shape (Word 365, Jason, 2026-10-21; the file the editor
  wrote and Word's repaired save on the share, `Office 2016/mc-fun/editor-cell-ending-in-table.docx`
  and `-repaired.docx`; ED-005 section 12.109).
- **The editor:** Insert Table at a caret in a cell, and a paste that lands a table at a cell's end,
  put the table after the cell's last paragraph, and the save writes the cell so, with no trailing
  paragraph.
- **Fixed** 2026-10-21 (ED-005 section 12.109): the edit layer appends the paragraph (the view's
  `cellEndsPlugin`, the headless runner, `repairedForSave`), never the export alone, so the file and
  the screen agree and Word opens the save without a repair. Word checks the last block as it reads it,
  through block-level `w:sdt` wrappers and the read branch of an `mc:AlternateContent` (Jason's run of
  twelve probes, 2026-10-22, the share's `Office 2016/cell-last-block/`), and the editor looks through
  block controls the same way since that day; a block-level `mc` in a cell is opaque to the editor and
  written by identity, so a file whose read branch ends in a table is Word's to repair. The rules'
  record across the family is docx4j's `docs/developer/word-table-rules.md`. Word's own Insert Table at a
  cell's end leaves the trailing paragraph in the cell's last paragraph's properties (core-ts check 52,
  2026-10-22), and the editor's appended paragraph copies them since that day; a last paragraph deleted after
  a nested table, which Word keeps emptied, is unmeasured in the editor (ED-005 section 12.109).

### The placeholder's wording

- **Word:** current Word's default placeholder reads "Click or tap here to enter text." (check 34, 1i).
- **The editor, before:** "Click here to enter text.", the engine's old `PLACEHOLDER_TEXT`, which is Word 2010's
  wording.
- **Fixed** in e4dca71 (2026-10-26, ED-008 section 8.8 item 14), on core-ts 0.6.0: the editor's own placeholders take
  Word's form. A control made at a caret and a bound control written an empty value show Word's default placeholder
  for their kind ("Click or tap here to enter text.", a date picker's "Click or tap to enter a date.", a list's "Choose
  an item."), in the `PlaceholderText` style with `w:showingPlcHdr`. They name the default entry in
  `w:placeholder/w:docPart`, and a save writes the entry into the glossary where the file lacks it.
- **Left:** a new picture control; see "A new picture control is never filled by Word".
