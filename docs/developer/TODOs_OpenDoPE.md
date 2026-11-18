# OpenDoPE and bound content controls: still to do

What Word (or the OpenDoPE specification) does that the editor does not do yet. Each item says what
Word does, what the editor does now, and what doing it would take. When one is done, move it to the
section at the end with the note that records it. Revised 2026-11-17.

The design is in:
- ED-003 sections 4 and 11.2: content controls and the Template perspective.
- ED-005 sections 12.54 and 12.55: the bound value shown, and the mapping both ways.
- ED-005 proposals 44 to 46, with their notes: escaped XHTML, pictures, Word fragments.

The notes are ED-005's, in its notes on content controls and data binding. Where the editor
differs from Word on purpose or for now, `docs/developer/word-divergences.md` says so.

## Kinds the editor doesn't bind yet

### A date picker

- **What Word does:**
  - **Showing:** a date picker bound to a node shows the node's value formatted by the control's
    `w:dateFormat` (`d/MM/yyyy`, `dddd, d MMMM yyyy`, …), in the language of its `w:lid`. It keeps
    the ISO value in `w:fullDate`.
  - **Writing back:** a date chosen or typed in the control is written to the node as the
    `w:storeMappedDataAs` attribute says: `dateTime` (the default), `date` or `text`.
  - **The calendar:** `w:calendar` names it (Gregorian by default).
- **What the editor does:** binding a date picker leaves its content as it is (ED-005 section 12.54),
  and an edit in it writes nothing back (section 12.55). The stale banner at open doesn't count it.
- **To do:**
  1. **Measure Word first:** a node holding a date without a time, a date and time, an invalid date,
     and an empty node; a control with no `w:dateFormat`; each `w:storeMappedDataAs`; two `w:lid`s.
  2. **Format:** the node's value by `w:dateFormat`, Word's own pattern language (`d`, `dd`, `ddd`,
     `dddd`, `M`, `MM`, `MMM`, `MMMM`, `yy`, `yyyy`, the time parts, quoted literals), with the
     month and day names from `Intl.DateTimeFormat` in the `w:lid`'s locale. Then set `w:fullDate`.
  3. **Write back:** the text read back through the same pattern into the `w:storeMappedDataAs`
     form; a text that does not parse writes nothing, as Word does (to measure).
  4. **Tests:** against Word's saves of the measurement file, as `check32.test.mts` holds the editor
     to Word's.

### A check box

- **What Word does:** a check box (`w14:checkbox`) bound to a node shows the node's `true`, `false`, `1` or `0`
  as the checked state, with the control's `w14:checkedState` or `w14:uncheckedState` symbol. A click writes it back.
- **What the editor does:** it shows and writes nothing.
- **To do:** measure which values Word reads and which it writes; then the symbol shown, and a click writing the node
  as Word does.

### A repeating section

- **What Word does:** a repeating section (`w15:repeatingSection`) bound to a repeating node adds and removes items as
  the data's nodes come and go.
- **What the editor does:** the item count doesn't follow the data.

## The mapping both ways: what is left

- **Notes and headers:** a bound control in a footnote, an endnote, a header or a footer is neither written back nor
  refreshed. The body's are.
- **A control bound by `od:xpath` alone,** with no `w:dataBinding`, is not written back. Word maps by `w:dataBinding`
  only; an OpenDoPE processor uses the XPaths entry.
- **Rich text written back.** Word writes a `w15:dataBinding` control's content into its node as escaped Flat OPC, a
  whole `pkg:package` (core-ts check 35, 1e: about 51 KB for a few words).
  - The editor writes nothing back from such a control, and it never writes over a node's Flat OPC from another control.
  - **So an edit there is lost in Word.** Word fills a `w15` control from its node when it opens the file (measured in
    ED-005 proposal 46's check 46c), so it replaces an edit made in the editor inside such a control with the node's content.
  - It is proposal 46's reverse direction, which decision (c) left out. The specification offers it, for `od:progid`
    fragments too.
  - Doing it would take the control's content written as Flat OPC, as Word writes one: the document, its styles, the
    glossary.
- **An undo from the content API** (`session.undoManager.undo()` in a script or a test) writes nothing back. The page's
  undo does.

## A control made and bound in one step: what is left

Word's answers came in core-ts check 34, and the editor follows them (ED-005 sections 12.56 and 12.57, held to Word's
saves by `check34.test.mts`). Across paragraphs and over a picture the editor refuses where Word does something else, and
an empty paragraph gets a run-level control where Word's pane makes a block-level one. Both are decided divergences in
`docs/developer/word-divergences.md`.

- **The control itself under tracking:** Word marks the new control inserted (`w:customXmlInsRangeStart`), and its Reject
  All removes the control and leaves the data as it was (`34b-tracked-reject-all.docx`).
  - The editor cannot mark a control (`word-divergences.md`, open): a reject leaves the control bound and writes the old
    text to the node.
  - To do, if Jason decides so: a control-level revision in the model.
- **A rich text control bound by `od:xpath` alone** (no `w:dataBinding` or `w15:dataBinding`): whether the Data tab's bind
  nests inside it or binds it in place is an open question for Jason (2026-10-03). Until it is settled, it is bound in place.
- **Dragging a node onto text with no control there makes nothing.** A drop binds the control under it, rich text and
  pictures included (Map to Selected), where the keyboard's bind would make or nest one.

## Pictures: what is left

Done in proposal 45 and after: see the end. Open:

- **A picture control holding text,** which Word repairs: one made around a selection, one made in a note or a header (the
  text placeholder, since the picture would need a relationship from that story's part), and one typed in (ED-005 section
  12.126 item 2). For Jason.
- **The floating case** (the specification's REQ-061): a floating picture filled through the engine's bind step in a tree
  run. The step is released (since core-ts 0.4.0), so nothing waits on the engine.
- **Writing back:** the editor's own write-back of a picture replaced in a bound control.
- **Tracking:** a picture bind as a tracked change.
- **A room:** the picture's part and drawing are one session's, not replicated to a peer.

## Escaped XHTML: what is left

Done in proposal 44: see the end. Open:

- **A tracked bind:** the conversion as a tracked change.
- **The conversion is the view's** (`view/xhtmlbind.ts`). A session with no view, such as the companion or a tree run,
  doesn't refresh an XHTML control.
- **A value of another level:** what the markup's level means when a node's value changes afterwards (a block-level value
  arriving in an inline control) is not handled. The converter puts what it can in the run-level control.

## Word fragments: what is left

Done in proposal 46: see the end. Open:

- **The reverse direction** (decision (c)): see "Rich text written back" above.
- **Comments in a fragment:** dropped, with a note, until a check in Word holds one.
- **Tracking:** a fragment's fill as a tracked change.
- **The engine's notes** from a fill (`refreshFragmentControls`) aren't shown in the status line. They are its own
  English, not localised (ED-005 sections 12.119 and 12.124).

## Done

- **The bound value shown on binding** (ED-005 section 12.54): plain and rich text, a drop-down's and a combo box's display
  text, a multi-line control's lines, the placeholder for an empty node.
- **The mapping both ways** (section 12.55): typing in a bound control writes its node, and the other controls bound to that
  node follow; a value edited in the Data tab shows in the controls; an undo on the page writes the old value back.
- **A control made and bound in one step** (section 12.56), confirmed and adjusted by core-ts check 34 (section 12.57): a
  control nested inside a rich text control, and the replacement tracked. The block-level control Word makes in an empty
  paragraph was dropped for a run-level one 2026-10-08 (section 12.71).
- **Rich text mapped as Word maps it** (section 12.58, core-ts check 35): the Data tab's bind nests inside a control mapped
  with `w15:dataBinding` and refuses its own node, as Word's pane does; a rich text control with `w:dataBinding` is plain
  text, bound in place, shown and written back; `w15` is declared ignorable for Word 2010.
- **The placeholder style** written at save where a run uses it (section 12.59), through the engine since core-ts 0.3.2.
- **Word's placeholder form** for the editor's own placeholders (ED-008 section 8.8 item 14): an empty node and a control made
  at a caret show Word's default placeholder for their kind. It is named in `w:placeholder/w:docPart` and written into the
  glossary at save where the file lacks it. A file control's own entry is shown as the file holds it.
- **The XPaths part:** the Data tab's XML view, Bind from its context menu and Copy XPath (section 12.64); bindings entered in
  the XPaths part once there is one (section 12.77); its boolean column (section 12.84); its binding attributes unqualified
  (section 12.87).
- **No control inside a control that cannot hold one** (section 12.85).
- **Escaped XHTML** (proposal 44; sections 12.75, 12.79, 12.86 and 12.117).
  - The pieces: the converter in core-ts (Apache-2.0); the engine's bind step; the Control tab's "XHTML content" switch;
    the view's refresh.
  - The Data tab's bind of a node holding markup makes a rich text control: block-level markup around a blank paragraph,
    refused inside one with content, and inline markup at the caret.
  - The Preview tab converts through the engine's bind step and draws in its own styles and lists.
  - Word check 37 kept every control.
- **Pictures** (proposal 45; sections 12.76, 12.117, 12.125 and 12.126).
  - Word's picture control and `od:Handler=picture` with a `width` are shown from the node.
  - The Control tab shows the binding's facts and sets a handler's `width`. The Data tab shows an image node as a
    thumbnail, and its bind there makes Word's picture control.
  - A node holding a data URL is read as its base64 (`imageBase64Of`). Word reads bare base64 only, which
    `word-divergences.md` records.
  - A new picture control holds Word's form, the `w:showingPlcHdr` flag and docx4j's placeholder picture from core-ts's
    `showPicturePlaceholder`, which Word fills from its node (section 12.126; merged at the core-ts 0.9.0 install).
- **Stale bound controls offered a refresh at open** (section 12.89, extended in 12.119). Word refreshes every bound control
  when it opens a file. The editor refreshes none unasked, so that an unedited file saves byte for byte.
  - Instead it counts the controls a refresh would change: text, XHTML, picture, and Word's `w15` fragment controls.
  - A banner offers Refresh from the data or Keep as it is.
- **The refreshes' notes reach the status line** (section 12.88).
- **Word fragments** (proposal 46; sections 12.118, 12.119 and 12.124; core-ts 0.8.0).
  - The engine fills a control tagged `od:progid=Word.Document`, or Word's own `w15`-mapped rich text, from a node holding
    Flat OPC. It reconciles styles, lists and notes as Word 365 does, and Word 365 keeps the fill exactly.
  - The Data tab shows a Flat OPC node read-only, and its bind makes Word's `w15` form, for a new control or one already
    there. Every bind and drop refreshes the control bound alone.
  - The Control tab's switch moves a control between Word's form and OpenDoPE's.
  - The Preview tab shows both forms filled.
