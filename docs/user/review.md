# Reviewing documents in the docx4j editor

This guide is for people who review Word documents in the editor: finding and replacing, tracking
changes, reading and resolving other people's changes, commenting, reviewing with others in a shared
document, and working beside an AI agent. The editor writes the same tracked changes and comments
Word does, so a document reviewed here opens in Word with its changes and comments intact, and the
other way round.

Where the editor knowingly behaves differently from Word, `docs/developer/word-divergences.md` says so.

## Find and replace

- **Ctrl+F** opens the find bar; **Ctrl+H** opens it with its Replace row. The arrow at its left
  shows or hides Replace.
- **Enter** or **F3** goes to the next result, **Shift+Enter** or **Shift+F3** to the previous one.
  **Escape** closes the bar and puts you back in the document where the result was.
- **Options:**
  - **Match case**: upper and lower case must match.
  - **Whole words**: only whole words are found.
  - **Wildcards**: Word's own: `?` one character, `*` any run, `[a-z]` a range, `[!x]` anything but,
    `<` and `>` a word's start and end, `@` repeats, `( )` a group, `\1` in the replacement for a
    group. As in Word, `?`, `*` and `[!x]` match a paragraph mark, so a result may run from one
    paragraph into the next, and `^13` is the mark itself.
  - **Regular expression**: a JavaScript regular expression within one paragraph, `$1` in the
    replacement for a group.
- **Word's codes** in a plain search: `^p` a paragraph mark, `^t` a tab, `^l` a line break, `^m` a
  page break, `^f` a footnote reference, `^e` an endnote reference, `^g` a picture, `^~` a non-breaking
  hyphen, `^-` an optional hyphen, `^^` a caret. A result with a paragraph mark in it is selected
  across the mark, into the next paragraph.
- **A result holding a paragraph mark is not replaced.** Replace skips it with a message, and Replace
  all does nothing while any result holds one, as Word's API refuses: delete or retype such a result
  by hand.
- **Order:** the body first, then the notes, then the headers and footers. A result in a header when
  headers are not shown is named in the status line instead.
- **Replace** replaces the current result and goes to the next (Enter in the replacement field).
  **Replace all** replaces every result, the notes and headers included, as one step that Ctrl+Z
  undoes whole (Ctrl+Enter in the replacement field).
- **Deleted text is not searched.** A tracked deletion is still on the page, but find reads the text
  as it would be once the changes were accepted, as Word's API does.
- **With tracking on,** a replacement is a tracked change like any other edit.

## Tracking changes

- **Turn it on** with **Track changes** in the application bar, or **Ctrl+Shift+E** from anywhere in
  the window. The button shows whether it is on.
- **Your name:** the first time, the editor asks for the name your changes carry ("Your name for
  tracked changes"). It is kept as the **Your name** option, which also names you to the people you
  share a document with.
- **It is the document's setting,** saved in the file as Word saves it, whichever perspective you
  are in. A document that arrives with tracking on is tracked from its first keystroke.
- **What is tracked:** typing and deleting, formatting text (bold, a font, a character style),
  changing a paragraph's properties (style, alignment, indents, spacing, a list), adding and deleting
  table rows. Deleting a heading's paragraph mark gives the next paragraph the heading's properties,
  as a tracked change, as Word does.
- **What is not:** comments (Word never tracks them), and the commands with no revision in Word to
  hold them: adding or deleting a table's columns, merging or splitting cells, a table's style, a
  section's settings or a section break, and content controls. Such a command runs untracked, and the
  status line says so.
- **Undo** (Ctrl+Z) takes back an edit and its tracking together.

## The three views

The Review perspective's panel has a **Show** group:

- **All markup**: insertions underlined and deletions struck through, in each author's colour, with a
  change bar in the margin; a formatting change underlined with a dotted line.
- **As accepted**: the document as it would be with every change accepted. You can edit in it.
- **Original**: the document as it was before the changes. It is read-only, and the status line says
  so when you choose it: switch to All markup to edit.

Each author has a colour and a hatch pattern of their own, on their changes and on the swatch beside
their name, so that authors differ by more than colour. The two approximations these views make are in
`docs/developer/word-divergences.md`.

## The review panel

Choose the **Review** perspective. Its **Changes** tab is the review panel, headed **Tracked
changes**. A document with tracked changes opened in another perspective offers the Review perspective
in its banner.

- **The count** says how many changes there are and by how many people, and how many are shown in
  Word only. It counts the entries listed here, as the banner does, which can be far fewer than the
  revisions in the file: ONLYOFFICE, refreshing a Contents with tracking on, wraps every piece of
  every line and leaves empty insertions beside each heading. An insertion with nothing in it is not
  listed and draws no change bar; **Accept all** and **Reject all** clear those too.
- **Previous** and **Next** step through the changes, selecting each in the document. **Alt+PageUp**
  and **Alt+PageDown** do the same from anywhere in the window. At the end they go round and say so.
- **Accept all** and **Reject all** resolve every change. **By author** groups the changes by their
  author, each group with its own **Accept all by** and **Reject all by**. A change whose file gives no
  name is under **Unknown author**, whether the name was left out or left blank.
- **Each entry** shows its kind, its author, its date and its text, and **Accept** and **Reject**:
  - **Inserted**, **Deleted**: text, a tab, a note reference ("Footnote 1"), a paragraph mark (¶).
  - **Formatted**, **Paragraph formatted**: with what changed, as Word's Reviewing Pane words it:
    "Font: Bold", "Font: 16 pt", "Centred", "Left indent: 36 pt", "Numbered + Level: 1 + 1.".
  - **Row inserted**, **Row deleted**: a table's rows. Rows touching other inserted or deleted content
    by the same author are one change with it, as Word's API lists them.
  - **Moved**: a move is one entry for both its halves. **Go to where it was** selects the other half.
    Accepting or rejecting either half resolves the whole move.
  - **Replaced**: a deletion and the same author's insertion right after it, as when someone types
    over a selection: "“7 days” with “14 days”". Its Accept and Reject resolve both halves in one step.
    **Split** shows the two halves apart, each with its own Accept and Reject; **Join** puts them back.
    Word lists a replacement as two revisions.
  - **In Word only**: a revision the editor keeps but does not draw (a table's or a section's
    properties, cells inserted or merged, a list's numbering). Where Word's API lists it, Accept and
    Reject resolve it; otherwise only Accept all and Reject all do.
- **Ctrl+Z** undoes an Accept or a Reject, each one step of its own.
- **Notes, headers and footers** keep their tracked changes and show them, but the panel lists the
  body's only, and says so at its foot.
- **Your own edits:** when tracking is off, the panel says "Your own edits are not tracked." with
  **Track my changes** and **Not now**. It says it once in each browser.

## Comments

- **New comment** (in the application bar, or **Ctrl+Alt+M**) comments on the selection or, at a caret, on
  the word the caret is in or next to, as Word's does. The first time, the editor asks for the name
  your comments carry.
- **The comments margin** shows each comment beside its text. Clicking a comment's head selects its
  text.
  - **Reply** opens a reply field; **Ctrl+Enter** posts it.
  - **Resolve** marks a thread done: it folds to its author and "resolved". Open the fold to read it,
    or to **Reopen** it.
  - **Delete** deletes a reply, or from its first comment the whole thread. **Ctrl+Z** brings it back.
  - **Your own comment's text** is edited in place. Other people's comments are read-only.
- **Comments are never tracked**, in Word or here.
- **A comment on text that a rejected insertion takes** goes with it, reference and all, as Word
  removes it; so does one on text whose deletion you accept. Ctrl+Z brings it back. A comment that a
  file anchored around the change, rather than inside it, stays.

## Reviewing with others

Sharing a document is in `co-editing.md`. In a shared document everyone's changes, comments and
resolutions reach everyone else as they are made.

**Resolving while someone types:** if you accept or reject a change while another person is typing
inside it, what you saw is resolved, and what they typed at the same moment stays theirs, pending.

**Known limitations.** Each needs two people doing something at the same moment, within the time a
message takes to reach the relay and come back:

- **Two people formatting overlapping text the same way** (bold, say) can leave the end of one
  person's range without the formatting. The text is all there: format it again.
- **An undo after such a resolution:** if you accept a change while someone types inside it and then
  undo your accept, the part of the change after their typing can stay accepted. Nothing is lost, and
  nothing is accepted twice.
- **Accept all or Reject all while someone edits:** a change that arrives while Accept all or Reject
  all is resolving the revisions shown in Word only can be resolved with them, without a word.
- **Typing inside text another person is deleting:** words typed into text that someone else deletes
  with tracking on, at that same moment, are taken into the deletion. With tracking off, your words
  cannot be told from their deletion.
- **Two people resolving the same changes opposite ways,** one accepting and the other rejecting at
  the same moment, can lose the text of both: an accepted deletion and a rejected insertion both take
  their text away. Either person's Ctrl+Z restores their own half; when both have undone, resolve it
  again. For a moved paragraph, resolving it again after both have undone can leave an empty
  paragraph where the move was: delete it.

## An agent in the document

When the editor page comes from the agent's companion (see `packages/editor-mcp/README.md` for how
to start it), the application bar has **Invite Claude** (or the agent's own name).

- **Inviting** shares the document into the companion's relay and asks the agent in. The agent can
  also open a file first and give you its link.
- **People** shows the agent with its colour and an **Agent** badge.
- **Every edit the agent makes is its tracked change,** whatever the document's tracking setting, for
  you to accept or reject. A few things have no revision in Word and are not tracked: its comments,
  and the styles, list definitions and links a new edit needs (`docs/developer/agent.md` section 1, rule 1).
- **What the agent did** in the review panel lists its calls, a refused one marked **Refused**.
  Clicking one selects where it acted.
- **If you type in a paragraph while the agent is editing it,** its edit is refused and it is told,
  so it can read the paragraph again and redo it. You see nothing of the refused edit.
- **Stop** beside the agent in People makes it leave the document: it can make no more edits there,
  nor open the file again, until you invite it back.

## Word 2010 and Word 15

A document reviewed here opens in both with every tracked change and comment.

- **Word 2010** has no threads: a reply shows as a comment of its own, and a resolved comment is not
  shown as done.
- **Word 15** keeps threads and the resolved state. Converting a file out of compatibility mode in
  Word 15 clears the resolved state; that is Word's behaviour.
